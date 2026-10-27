"""Testes do core.ole_processor.

Testes que dependem de Win32 Structured Storage são pulados automaticamente em
ambientes não-Windows ou sem pywin32.
"""
from __future__ import annotations

import io
import struct
import tempfile
from pathlib import Path

import pytest
from PIL import Image

from core.ole_processor import (
    PBRUSH_CLSID,
    comprimir_ole_pbrush,
    is_ole_rewrite_available,
)

pytestmark = pytest.mark.skipif(
    not is_ole_rewrite_available(),
    reason="Win32 Structured Storage indisponível neste ambiente",
)


def _make_bmp(size: tuple[int, int]) -> bytes:
    img = Image.new("RGB", size, (180, 120, 60))
    buf = io.BytesIO()
    img.save(buf, format="BMP")
    return buf.getvalue()


def _make_pbrush_ole(bmp_bytes: bytes, with_print: bool = True) -> bytes:
    """Cria um OLE compound mínimo no formato PBrush para teste."""
    import pythoncom
    import pywintypes
    from win32com import storagecon

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
                storage.SetClass(pywintypes.IID("{" + PBRUSH_CLSID + "}"))

                def write(name: str, data: bytes) -> None:
                    s = storage.CreateStream(
                        name,
                        storagecon.STGM_CREATE
                        | storagecon.STGM_READWRITE
                        | storagecon.STGM_SHARE_EXCLUSIVE,
                        0,
                        0,
                    )
                    s.Write(data)
                    s.Commit(0)

                write("\x01CompObj", b"\x01\x00\xfe\xff" + b"\x00" * 60)
                write("\x01Ole", b"\x01\x00\x00\x02" + b"\x00" * 16)
                write("\x01Ole10ItemName", b"Bitmap Image\x00")
                write("\x03ObjInfo", b"\x00" * 6)
                native = struct.pack("<I", len(bmp_bytes)) + bmp_bytes
                write("\x01Ole10Native", native)
                if with_print:
                    write("\x03PRINT", b"\x00" * (1024 * 1024))
                storage.Commit(0)
            finally:
                del storage
            return tmp_path.read_bytes()
        finally:
            tmp_path.unlink(missing_ok=True)
    finally:
        pythoncom.CoUninitialize()


def test_pbrush_grande_eh_redimensionado_e_print_descartado():
    bmp = _make_bmp((4000, 3000))
    bin_data = _make_pbrush_ole(bmp, with_print=True)
    assert len(bin_data) > 1024 * 1024  # tem PRINT de 1MB

    new_bin, result = comprimir_ole_pbrush(bin_data, max_dim=1920)

    assert result.was_processed is True
    assert result.orig_dims == (4000, 3000)
    assert max(result.new_dims) == 1920
    assert result.new_size < result.orig_size
    # PRINT (1MB) descartado + BMP redimensionado
    assert new_bin != bin_data


def test_pbrush_pequeno_ainda_processado_para_descartar_print():
    bmp = _make_bmp((400, 300))
    bin_data = _make_pbrush_ole(bmp, with_print=True)

    new_bin, result = comprimir_ole_pbrush(bin_data, max_dim=1920)

    assert result.was_processed is True
    assert result.orig_dims == (400, 300)
    assert result.new_dims == (400, 300)
    assert result.new_size < result.orig_size


def test_ole_nao_pbrush_eh_pulado():
    bin_data = b"isto nao eh um ole valido"
    new_bin, result = comprimir_ole_pbrush(bin_data, max_dim=1920)

    assert result.was_processed is False
    assert "PBrush" in result.skip_reason or "inválido" in result.skip_reason.lower()
    assert new_bin == bin_data


def test_pbrush_sem_ole10native_eh_pulado():
    import pythoncom
    import pywintypes
    from win32com import storagecon

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
                storage.SetClass(pywintypes.IID("{" + PBRUSH_CLSID + "}"))
                s = storage.CreateStream(
                    "\x01CompObj",
                    storagecon.STGM_CREATE
                    | storagecon.STGM_READWRITE
                    | storagecon.STGM_SHARE_EXCLUSIVE,
                    0,
                    0,
                )
                s.Write(b"\x00" * 32)
                s.Commit(0)
                storage.Commit(0)
            finally:
                del storage
            data = tmp_path.read_bytes()
        finally:
            tmp_path.unlink(missing_ok=True)
    finally:
        pythoncom.CoUninitialize()

    new_bin, result = comprimir_ole_pbrush(data, max_dim=1920)
    assert result.was_processed is False
    assert "Ole10Native" in result.skip_reason
    assert new_bin == data


def test_pbrush_com_ole10native_nao_bmp_eh_pulado():
    """Native válido em estrutura, mas conteúdo não é BMP."""
    fake_native_payload = b"\x89PNG\r\n\x1a\n" + b"\x00" * 100  # PNG header
    bin_data = _make_pbrush_ole(fake_native_payload, with_print=False)

    new_bin, result = comprimir_ole_pbrush(bin_data, max_dim=1920)

    assert result.was_processed is False
    assert "BMP" in result.skip_reason


def test_resultado_preserva_pbrush_clsid_e_streams_essenciais():
    """O OLE reescrito deve continuar sendo PBrush e ter o Ole10Native novo."""
    import olefile

    bmp = _make_bmp((2400, 1600))
    bin_data = _make_pbrush_ole(bmp, with_print=True)

    new_bin, result = comprimir_ole_pbrush(bin_data, max_dim=1920)
    assert result.was_processed

    ole = olefile.OleFileIO(io.BytesIO(new_bin))
    try:
        assert str(ole.root.clsid).upper() == PBRUSH_CLSID
        names = {entry[0] for entry in ole.listdir(streams=True, storages=False)}
        assert "\x01Ole10Native" in names
        assert "\x01CompObj" in names
        assert "\x01Ole" in names
        # \x03PRINT deve ter sido descartado
        assert "\x03PRINT" not in names

        new_native = ole.openstream("\x01Ole10Native").read()
        declared = struct.unpack("<I", new_native[:4])[0]
        bmp_payload = new_native[4 : 4 + declared]
        assert bmp_payload[:2] == b"BM"
        with Image.open(io.BytesIO(bmp_payload)) as img:
            assert max(img.size) == 1920
    finally:
        ole.close()
