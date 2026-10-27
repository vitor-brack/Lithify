"""Compressão de OLE compound files do tipo PBrush (Paint).

Quando o usuário cola fotos no Word usando "Colar como bitmap" (ou via outras
fontes), o documento armazena cada imagem como um OLE Embedded Object com
CLSID `0003000A-...` (PBrush). Cada objeto contém:

- `\\x01Ole10Native`: BMP cru (não comprimido) — pode ter ~26 MB para uma foto 4K.
- `\\x03PRINT`: cópia metafile do mesmo bitmap, ~26 MB.

Este módulo reescreve esses OLEs reduzindo a dimensão do BMP em `Ole10Native`
e descartando `\\x03PRINT` (o Word regenera ao salvar como `.doc`). O CLSID
e os streams de metadata (`CompObj`, `Ole`, etc.) são preservados, então o
double-click no Word continua abrindo o Paint normalmente.

Requer Win32 Structured Storage API (pywin32 em Windows).
"""
from __future__ import annotations

import io
import logging
import struct
import tempfile
from dataclasses import dataclass
from pathlib import Path

import olefile
from PIL import Image, UnidentifiedImageError

logger = logging.getLogger(__name__)

PBRUSH_CLSID = "0003000A-0000-0000-C000-000000000046"
EMBEDDINGS_PREFIX = "word/embeddings/"

try:
    import pythoncom
    import pywintypes
    from win32com import storagecon

    _HAS_STG = True
except ImportError:
    pythoncom = None  # type: ignore[assignment]
    pywintypes = None  # type: ignore[assignment]
    storagecon = None  # type: ignore[assignment]
    _HAS_STG = False


@dataclass(frozen=True)
class OleCompressionResult:
    """Resultado da tentativa de comprimir um único OLE bin."""

    was_processed: bool
    skip_reason: str = ""
    orig_size: int = 0
    new_size: int = 0
    orig_dims: tuple[int, int] = (0, 0)
    new_dims: tuple[int, int] = (0, 0)


def is_ole_rewrite_available() -> bool:
    """True se a API Win32 de Structured Storage estiver disponível."""
    return _HAS_STG


def comprimir_ole_pbrush(
    bin_data: bytes,
    max_dim: int = 1920,
) -> tuple[bytes, OleCompressionResult]:
    """Comprime um OLE compound do tipo PBrush.

    Retorna `(novos_bytes, resultado)`. Se o OLE não for PBrush, ou estiver
    corrompido, ou não tiver `Ole10Native` válido, retorna os bytes originais
    intocados com `was_processed=False` e `skip_reason` preenchido.
    """
    if not _HAS_STG:
        return bin_data, OleCompressionResult(
            was_processed=False,
            skip_reason="API Win32 Structured Storage indisponível (Windows + pywin32)",
            orig_size=len(bin_data),
            new_size=len(bin_data),
        )

    parsed = _parse_pbrush_streams(bin_data)
    if parsed is None:
        return bin_data, OleCompressionResult(
            was_processed=False,
            skip_reason="OLE não-PBrush ou inválido",
            orig_size=len(bin_data),
            new_size=len(bin_data),
        )
    clsid, streams = parsed

    native = streams.get("\x01Ole10Native")
    if not native or len(native) < 6:
        return bin_data, OleCompressionResult(
            was_processed=False,
            skip_reason="Sem stream Ole10Native válido",
            orig_size=len(bin_data),
            new_size=len(bin_data),
        )

    declared = struct.unpack("<I", native[:4])[0]
    payload = native[4 : 4 + declared]
    if len(payload) < 6 or payload[:2] != b"BM":
        return bin_data, OleCompressionResult(
            was_processed=False,
            skip_reason="Payload de Ole10Native não é BMP",
            orig_size=len(bin_data),
            new_size=len(bin_data),
        )

    try:
        with Image.open(io.BytesIO(payload)) as img:
            img.load()
            orig_dims = img.size
            resized = _resize_if_needed(img, max_dim)
            if resized.mode != "RGB":
                resized = resized.convert("RGB")
            new_dims = resized.size
            buf = io.BytesIO()
            resized.save(buf, format="BMP")
            new_bmp = buf.getvalue()
    except (UnidentifiedImageError, OSError, ValueError) as exc:
        logger.warning("BMP inválido em OLE: %s", exc)
        return bin_data, OleCompressionResult(
            was_processed=False,
            skip_reason=f"BMP inválido: {exc}",
            orig_size=len(bin_data),
            new_size=len(bin_data),
        )

    new_native = struct.pack("<I", len(new_bmp)) + new_bmp

    try:
        new_bin = _build_pbrush_compound(clsid, streams, new_native)
    except Exception as exc:
        logger.warning("Falha ao reconstruir OLE compound: %s", exc, exc_info=True)
        return bin_data, OleCompressionResult(
            was_processed=False,
            skip_reason=f"Falha ao reconstruir compound: {exc}",
            orig_size=len(bin_data),
            new_size=len(bin_data),
        )

    if len(new_bin) >= len(bin_data):
        return bin_data, OleCompressionResult(
            was_processed=False,
            skip_reason="Reconstrução não reduziu o tamanho",
            orig_size=len(bin_data),
            new_size=len(bin_data),
            orig_dims=orig_dims,
            new_dims=new_dims,
        )

    return new_bin, OleCompressionResult(
        was_processed=True,
        orig_size=len(bin_data),
        new_size=len(new_bin),
        orig_dims=orig_dims,
        new_dims=new_dims,
    )


def _parse_pbrush_streams(bin_data: bytes) -> tuple[str, dict[str, bytes]] | None:
    try:
        ole = olefile.OleFileIO(io.BytesIO(bin_data))
    except Exception:
        return None
    try:
        clsid = str(ole.root.clsid).upper()
        if clsid != PBRUSH_CLSID:
            return None
        streams: dict[str, bytes] = {}
        for entry in ole.listdir(streams=True, storages=False):
            name = entry[0]
            streams[name] = ole.openstream(name).read()
        return clsid, streams
    finally:
        ole.close()


def _resize_if_needed(img: Image.Image, max_dim: int) -> Image.Image:
    w, h = img.size
    longest = max(w, h)
    if longest <= max_dim:
        return img
    ratio = max_dim / longest
    new_size = (max(1, int(round(w * ratio))), max(1, int(round(h * ratio))))
    return img.resize(new_size, Image.Resampling.LANCZOS)


def _build_pbrush_compound(
    clsid: str,
    original_streams: dict[str, bytes],
    new_native: bytes,
) -> bytes:
    """Monta um OLE compound novo preservando metadata e omitindo `\\x03PRINT`."""
    assert _HAS_STG, "Structured Storage API not available"
    pythoncom.CoInitialize()
    try:
        with tempfile.NamedTemporaryFile(suffix=".bin", delete=False) as tmp:
            tmp_path = Path(tmp.name)
        tmp_path.unlink(missing_ok=True)
        try:
            storage = pythoncom.StgCreateDocfile(
                str(tmp_path),
                storagecon.STGM_CREATE
                | storagecon.STGM_READWRITE
                | storagecon.STGM_SHARE_EXCLUSIVE,
            )
            try:
                storage.SetClass(pywintypes.IID("{" + clsid + "}"))
                _write_stream(storage, "\x01CompObj", original_streams.get("\x01CompObj", b""))
                _write_stream(storage, "\x01Ole", original_streams.get("\x01Ole", b""))
                _write_stream(
                    storage,
                    "\x01Ole10ItemName",
                    original_streams.get("\x01Ole10ItemName", b""),
                )
                _write_stream(storage, "\x03ObjInfo", original_streams.get("\x03ObjInfo", b""))
                _write_stream(storage, "\x01Ole10Native", new_native)
                # \x03PRINT intencionalmente omitido — o Word regenera ao salvar como .doc.
                storage.Commit(0)
            finally:
                del storage
            return tmp_path.read_bytes()
        finally:
            tmp_path.unlink(missing_ok=True)
    finally:
        pythoncom.CoUninitialize()


def _write_stream(storage: object, name: str, data: bytes) -> None:
    if not data:
        return
    s = storage.CreateStream(  # type: ignore[attr-defined]
        name,
        storagecon.STGM_CREATE | storagecon.STGM_READWRITE | storagecon.STGM_SHARE_EXCLUSIVE,
        0,
        0,
    )
    s.Write(data)
    s.Commit(0)
