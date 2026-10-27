"""Orquestrador para `.doc` e `.docx`.

Detecta o formato pelo magic byte e dispatcha:
- `.docx` → `comprimir_docx` direto.
- `.doc` → conversão Word→docx → `comprimir_docx` → conversão docx→Word.
"""
from __future__ import annotations

import logging
import tempfile
from pathlib import Path
from typing import Callable, Literal, Optional

from core.doc_converter import is_word_available, word_session
from core.docx_processor import DocxCompressionResult, comprimir_docx
from core.exceptions import InvalidDocxError, WordNotAvailableError

logger = logging.getLogger(__name__)

DOCX_MAGIC = b"PK\x03\x04"
DOC_MAGIC = b"\xd0\xcf\x11\xe0"

DocumentFormat = Literal["docx", "doc"]

StageCallback = Callable[[str], None]
ProgressCallback = Callable[[int, int, str], None]


def detect_format(path: Path) -> DocumentFormat:
    """Detecta o formato pelos magic bytes (com fallback para extensão)."""
    if not path.exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {path}")

    with open(path, "rb") as f:
        magic = f.read(8)

    if magic.startswith(DOCX_MAGIC):
        return "docx"
    if magic.startswith(DOC_MAGIC):
        return "doc"

    ext = path.suffix.lower()
    if ext == ".doc":
        return "doc"
    if ext == ".docx":
        raise InvalidDocxError(
            f"Arquivo {path.name} tem extensão .docx mas o conteúdo não é um ZIP válido."
        )
    raise InvalidDocxError(
        f"Arquivo {path.name} não é um documento Word reconhecível (.doc ou .docx)."
    )


def comprimir_documento(
    input_path: Path | str,
    output_path: Path | str,
    max_dim: int = 2000,
    quality: int = 85,
    preserve_alpha: bool = False,
    progress_callback: Optional[ProgressCallback] = None,
    stage_callback: Optional[StageCallback] = None,
) -> DocxCompressionResult:
    """Comprime imagens em um documento Word (`.doc` ou `.docx`).

    `stage_callback` é chamado com mensagens de status entre as fases
    (conversão, compressão, conversão de volta) — útil para a GUI atualizar
    um label durante operações longas.
    """
    input_path = Path(input_path)
    output_path = Path(output_path)

    if input_path.resolve() == output_path.resolve():
        raise ValueError("output_path não pode ser igual a input_path")

    fmt = detect_format(input_path)

    if fmt == "docx":
        if stage_callback:
            stage_callback("Comprimindo imagens...")
        return comprimir_docx(
            input_path=input_path,
            output_path=output_path,
            max_dim=max_dim,
            quality=quality,
            preserve_alpha=preserve_alpha,
            progress_callback=progress_callback,
        )

    return _comprimir_doc(
        input_path=input_path,
        output_path=output_path,
        max_dim=max_dim,
        quality=quality,
        preserve_alpha=preserve_alpha,
        progress_callback=progress_callback,
        stage_callback=stage_callback,
    )


def _comprimir_doc(
    input_path: Path,
    output_path: Path,
    max_dim: int,
    quality: int,
    preserve_alpha: bool,
    progress_callback: Optional[ProgressCallback],
    stage_callback: Optional[StageCallback],
) -> DocxCompressionResult:
    if not is_word_available():
        raise WordNotAvailableError(
            "Para comprimir arquivos .doc, é necessário ter o Microsoft Word "
            "instalado neste computador.\n\n"
            "Alternativa: abra o arquivo no Word, vá em "
            "Arquivo → Salvar como → Documento do Word (*.docx), "
            "e use o Lithify normalmente com o arquivo .docx."
        )

    original_size = input_path.stat().st_size
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with tempfile.TemporaryDirectory(prefix="lithify_doc_") as tmpdir_str:
        tmpdir = Path(tmpdir_str)
        intermediate_docx = tmpdir / f"{input_path.stem}.docx"
        compressed_docx = tmpdir / f"{input_path.stem}_comp.docx"

        if stage_callback:
            stage_callback("Convertendo .doc para .docx (Word abrindo)...")
        with word_session() as session:
            session.doc_to_docx(input_path, intermediate_docx)

        if stage_callback:
            stage_callback("Comprimindo imagens...")
        inner_result = comprimir_docx(
            input_path=intermediate_docx,
            output_path=compressed_docx,
            max_dim=max_dim,
            quality=quality,
            preserve_alpha=preserve_alpha,
            progress_callback=progress_callback,
        )

        if stage_callback:
            stage_callback("Convertendo de volta para .doc (Word abrindo)...")
        with word_session() as session:
            session.docx_to_doc(compressed_docx, output_path)

    compressed_size = output_path.stat().st_size

    return DocxCompressionResult(
        input_path=input_path,
        output_path=output_path,
        original_size=original_size,
        compressed_size=compressed_size,
        images_processed=inner_result.images_processed,
        images_skipped_already_optimized=inner_result.images_skipped_already_optimized,
        images_skipped_non_raster=inner_result.images_skipped_non_raster,
        images_failed=inner_result.images_failed,
        oles_processed=inner_result.oles_processed,
        oles_skipped=inner_result.oles_skipped,
    )
