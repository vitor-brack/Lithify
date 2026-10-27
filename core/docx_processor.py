"""Compressão de imagens embutidas em documentos `.docx`."""
from __future__ import annotations

import logging
import tempfile
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Optional

from core.exceptions import (
    CorruptImageError,
    InvalidDocxError,
    LegacyDocFormatError,
    UnsupportedFormatError,
)
from core.image_compressor import comprimir_imagem
from core.ole_processor import (
    EMBEDDINGS_PREFIX,
    comprimir_ole_pbrush,
    is_ole_rewrite_available,
)

logger = logging.getLogger(__name__)

WORD_MEDIA_PREFIX = "word/media/"
RASTER_EXTENSIONS = {".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff", ".webp"}
DOCX_MAGIC = b"PK\x03\x04"
LEGACY_DOC_MAGIC = b"\xd0\xcf\x11\xe0"


@dataclass(frozen=True)
class DocxCompressionResult:
    """Resultado da compressão de um `.docx`."""

    input_path: Path
    output_path: Path
    original_size: int
    compressed_size: int
    images_processed: int
    images_skipped_already_optimized: int
    images_skipped_non_raster: int
    images_failed: int
    oles_processed: int = 0
    oles_skipped: int = 0

    @property
    def saved_bytes(self) -> int:
        return self.original_size - self.compressed_size

    @property
    def saved_ratio(self) -> float:
        if self.original_size == 0:
            return 0.0
        return self.saved_bytes / self.original_size

    @property
    def total_images(self) -> int:
        return (
            self.images_processed
            + self.images_skipped_already_optimized
            + self.images_skipped_non_raster
            + self.images_failed
        )

    @property
    def total_oles(self) -> int:
        return self.oles_processed + self.oles_skipped


def comprimir_docx(
    input_path: Path | str,
    output_path: Path | str,
    max_dim: int = 2000,
    quality: int = 85,
    preserve_alpha: bool = False,
    progress_callback: Optional[Callable[[int, int, str], None]] = None,
) -> DocxCompressionResult:
    """Comprime as imagens rasterizadas dentro de um `.docx`.

    O arquivo de entrada nunca é modificado. A estrutura interna do `.docx` é
    preservada — apenas os bytes dos arquivos em `word/media/` são substituídos
    por versões comprimidas. Os nomes em `word/media/` são preservados, mesmo
    se o conteúdo binário mudar de PNG para JPEG.
    """
    input_path = Path(input_path)
    output_path = Path(output_path)

    if input_path.resolve() == output_path.resolve():
        raise ValueError("output_path não pode ser igual a input_path")

    _validate_docx(input_path)

    original_size = input_path.stat().st_size
    output_path.parent.mkdir(parents=True, exist_ok=True)

    rasters: list[zipfile.ZipInfo] = []
    non_raster_media: list[zipfile.ZipInfo] = []
    embeddings: list[zipfile.ZipInfo] = []
    with zipfile.ZipFile(input_path, "r") as zin:
        for entry in zin.infolist():
            if entry.is_dir():
                continue
            name = entry.filename
            if name.startswith(WORD_MEDIA_PREFIX):
                ext = Path(name).suffix.lower()
                if ext in RASTER_EXTENSIONS:
                    rasters.append(entry)
                else:
                    non_raster_media.append(entry)
            elif name.startswith(EMBEDDINGS_PREFIX) and name.endswith(".bin"):
                embeddings.append(entry)

    replaced_bytes: dict[str, bytes] = {}
    processed = 0
    skipped_already_optimized = 0
    failed = 0
    oles_processed = 0
    oles_skipped = 0

    total_items = len(rasters) + len(embeddings)
    item_idx = 0

    if rasters:
        with tempfile.TemporaryDirectory(prefix="lithify_") as tmpdir_str:
            tmpdir = Path(tmpdir_str)
            with zipfile.ZipFile(input_path, "r") as zin:
                for entry in rasters:
                    item_idx += 1
                    if progress_callback:
                        progress_callback(item_idx, total_items, Path(entry.filename).name)

                    original_bytes = zin.read(entry.filename)
                    new_bytes = _try_compress_entry(
                        entry=entry,
                        original_bytes=original_bytes,
                        tmpdir=tmpdir,
                        max_dim=max_dim,
                        quality=quality,
                        preserve_alpha=preserve_alpha,
                    )

                    if new_bytes is None:
                        failed += 1
                        replaced_bytes[entry.filename] = original_bytes
                    elif new_bytes is original_bytes:
                        skipped_already_optimized += 1
                        replaced_bytes[entry.filename] = original_bytes
                    else:
                        processed += 1
                        replaced_bytes[entry.filename] = new_bytes

    if embeddings:
        if not is_ole_rewrite_available():
            logger.warning(
                "Win32 Structured Storage indisponível: %d OLE bins não serão processados.",
                len(embeddings),
            )
            oles_skipped = len(embeddings)
        else:
            with zipfile.ZipFile(input_path, "r") as zin:
                for entry in embeddings:
                    item_idx += 1
                    if progress_callback:
                        progress_callback(item_idx, total_items, Path(entry.filename).name)

                    original_bytes = zin.read(entry.filename)
                    new_bytes, ole_result = comprimir_ole_pbrush(
                        bin_data=original_bytes,
                        max_dim=max_dim,
                    )
                    if ole_result.was_processed:
                        oles_processed += 1
                        replaced_bytes[entry.filename] = new_bytes
                        logger.info(
                            "OLE %s: %s -> %s (%d×%d -> %d×%d)",
                            entry.filename,
                            _fmt_mb(ole_result.orig_size),
                            _fmt_mb(ole_result.new_size),
                            *ole_result.orig_dims,
                            *ole_result.new_dims,
                        )
                    else:
                        oles_skipped += 1
                        logger.debug(
                            "OLE %s pulado: %s",
                            entry.filename,
                            ole_result.skip_reason,
                        )

    tmp_output = output_path.with_name(output_path.name + ".tmp")
    try:
        with zipfile.ZipFile(input_path, "r") as zin, zipfile.ZipFile(
            tmp_output, "w", zipfile.ZIP_DEFLATED
        ) as zout:
            for entry in zin.infolist():
                data = (
                    replaced_bytes[entry.filename]
                    if entry.filename in replaced_bytes
                    else zin.read(entry.filename)
                )
                zout.writestr(_clone_zipinfo(entry), data)
        tmp_output.replace(output_path)
    except Exception:
        tmp_output.unlink(missing_ok=True)
        raise

    compressed_size = output_path.stat().st_size

    return DocxCompressionResult(
        input_path=input_path,
        output_path=output_path,
        original_size=original_size,
        compressed_size=compressed_size,
        images_processed=processed,
        images_skipped_already_optimized=skipped_already_optimized,
        images_skipped_non_raster=len(non_raster_media),
        images_failed=failed,
        oles_processed=oles_processed,
        oles_skipped=oles_skipped,
    )


def _fmt_mb(n: int) -> str:
    return f"{n / 1024 / 1024:.1f} MB"


def _try_compress_entry(
    entry: zipfile.ZipInfo,
    original_bytes: bytes,
    tmpdir: Path,
    max_dim: int,
    quality: int,
    preserve_alpha: bool,
) -> bytes | None:
    """Tenta comprimir uma entrada de mídia. Retorna:

    - `original_bytes` (mesmo objeto) se a imagem já estava otimizada (skip);
    - bytes novos se houve compressão real e ficou menor;
    - `original_bytes` se a compressão ficou maior ou igual (mantém original);
    - `None` se houve falha (corrupção etc.).
    """
    name = Path(entry.filename).name
    tmp_in = tmpdir / f"in_{name}"
    tmp_out = tmpdir / f"out_{name}"
    tmp_in.write_bytes(original_bytes)

    try:
        result = comprimir_imagem(
            input_path=tmp_in,
            output_path=tmp_out,
            max_dim=max_dim,
            quality=quality,
            preserve_alpha=preserve_alpha,
        )
    except (CorruptImageError, UnsupportedFormatError, FileNotFoundError) as exc:
        logger.warning("Falha ao comprimir %s: %s", entry.filename, exc)
        return None
    finally:
        tmp_in.unlink(missing_ok=True)

    try:
        if result.was_skipped:
            return original_bytes

        new_bytes = result.output_path.read_bytes()
        if len(new_bytes) >= len(original_bytes):
            return original_bytes
        return new_bytes
    finally:
        if result.output_path.exists():
            result.output_path.unlink(missing_ok=True)


def _validate_docx(path: Path) -> None:
    if not path.exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {path}")

    if path.suffix.lower() == ".doc":
        raise LegacyDocFormatError(
            "Formato .doc antigo não é suportado. Abra no Word, vá em "
            "Arquivo → Salvar como → Documento do Word (.docx), e tente novamente."
        )

    with open(path, "rb") as f:
        magic = f.read(8)

    if magic.startswith(LEGACY_DOC_MAGIC):
        raise LegacyDocFormatError(
            "Formato .doc antigo não é suportado. Abra no Word, vá em "
            "Arquivo → Salvar como → Documento do Word (.docx), e tente novamente."
        )

    if not magic.startswith(DOCX_MAGIC):
        raise InvalidDocxError(f"Arquivo não é um .docx válido: {path.name}")

    try:
        with zipfile.ZipFile(path, "r") as zf:
            bad = zf.testzip()
            if bad is not None:
                raise InvalidDocxError(f".docx corrompido (entrada inválida: {bad})")
    except zipfile.BadZipFile as exc:
        raise InvalidDocxError(f".docx corrompido: {path.name}") from exc


def _clone_zipinfo(src: zipfile.ZipInfo) -> zipfile.ZipInfo:
    new = zipfile.ZipInfo(filename=src.filename, date_time=src.date_time)
    new.compress_type = src.compress_type
    new.external_attr = src.external_attr
    new.create_system = src.create_system
    new.flag_bits = src.flag_bits
    new.internal_attr = src.internal_attr
    return new
