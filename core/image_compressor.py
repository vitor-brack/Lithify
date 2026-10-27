"""Núcleo de compressão de imagens rasterizadas."""
from __future__ import annotations

import logging
import shutil
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Iterable, Optional

from PIL import Image, ImageOps, UnidentifiedImageError

from core.exceptions import CorruptImageError, UnsupportedFormatError

logger = logging.getLogger(__name__)

SUPPORTED_EXTENSIONS = {".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff", ".webp"}
JPEG_EXTENSIONS = {".jpg", ".jpeg"}
SKIP_THRESHOLD_BYTES = 150 * 1024


@dataclass(frozen=True)
class CompressionResult:
    """Resultado da compressão de uma imagem."""

    input_path: Path
    output_path: Path
    original_size: int
    compressed_size: int
    original_dims: tuple[int, int]
    final_dims: tuple[int, int]
    output_format: str
    was_skipped: bool

    @property
    def saved_bytes(self) -> int:
        return self.original_size - self.compressed_size

    @property
    def saved_ratio(self) -> float:
        if self.original_size == 0:
            return 0.0
        return self.saved_bytes / self.original_size


def comprimir_imagem(
    input_path: Path | str,
    output_path: Path | str,
    max_dim: int = 2000,
    quality: int = 85,
    preserve_alpha: bool = False,
) -> CompressionResult:
    """Comprime uma imagem rasterizada.

    A extensão de `output_path` pode ser alterada (ex.: PNG sem `preserve_alpha`
    vira `.jpg`). O caminho final efetivamente escrito está em
    `CompressionResult.output_path`.
    """
    input_path = Path(input_path)
    output_path = Path(output_path)

    if not input_path.exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {input_path}")

    ext = input_path.suffix.lower()
    if ext not in SUPPORTED_EXTENSIONS:
        raise UnsupportedFormatError(f"Formato não suportado: {ext}")

    original_size = input_path.stat().st_size

    try:
        with Image.open(input_path) as probe:
            probe.load()
            original_dims = probe.size
            has_alpha = _has_alpha(probe)
    except (UnidentifiedImageError, OSError, ValueError) as exc:
        raise CorruptImageError(f"Imagem ilegível: {input_path.name}") from exc

    if _should_skip(ext, original_size, original_dims, max_dim):
        return _copy_as_skipped(
            input_path=input_path,
            output_path=output_path,
            original_size=original_size,
            original_dims=original_dims,
        )

    output_as_png = preserve_alpha and has_alpha
    final_ext = ".png" if output_as_png else ".jpg"
    final_output = output_path.with_suffix(final_ext)
    final_output.parent.mkdir(parents=True, exist_ok=True)

    try:
        with Image.open(input_path) as img:
            img = ImageOps.exif_transpose(img)
            img = _resize_if_needed(img, max_dim)
            final_dims = img.size

            if output_as_png:
                if img.mode not in ("RGBA", "LA", "P"):
                    img = img.convert("RGBA")
                img.save(final_output, format="PNG", optimize=True)
                output_format = "PNG"
            else:
                if has_alpha:
                    img = _flatten_to_white(img)
                elif img.mode != "RGB":
                    img = img.convert("RGB")
                img.save(
                    final_output,
                    format="JPEG",
                    quality=quality,
                    optimize=True,
                    progressive=True,
                )
                output_format = "JPEG"
    except (UnidentifiedImageError, OSError, ValueError) as exc:
        if final_output.exists():
            final_output.unlink(missing_ok=True)
        raise CorruptImageError(f"Falha ao processar: {input_path.name}") from exc

    compressed_size = final_output.stat().st_size

    return CompressionResult(
        input_path=input_path,
        output_path=final_output,
        original_size=original_size,
        compressed_size=compressed_size,
        original_dims=original_dims,
        final_dims=final_dims,
        output_format=output_format,
        was_skipped=False,
    )


def comprimir_lote(
    inputs: Iterable[Path | str],
    output_dir: Path | str,
    max_dim: int = 2000,
    quality: int = 85,
    preserve_alpha: bool = False,
    progress_callback: Optional[Callable[[int, int, str], None]] = None,
) -> list[CompressionResult]:
    """Comprime múltiplas imagens. Falhas em arquivos individuais são logadas e o lote continua."""
    inputs_list = [Path(p) for p in inputs]
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    total = len(inputs_list)
    results: list[CompressionResult] = []

    for i, input_path in enumerate(inputs_list, start=1):
        if progress_callback:
            progress_callback(i, total, input_path.name)
        try:
            target = output_dir / input_path.name
            result = comprimir_imagem(
                input_path=input_path,
                output_path=target,
                max_dim=max_dim,
                quality=quality,
                preserve_alpha=preserve_alpha,
            )
            results.append(result)
        except (CorruptImageError, UnsupportedFormatError, FileNotFoundError) as exc:
            logger.warning("Pulando %s: %s", input_path.name, exc)
            continue

    return results


def _copy_as_skipped(
    input_path: Path,
    output_path: Path,
    original_size: int,
    original_dims: tuple[int, int],
) -> CompressionResult:
    final_output = output_path.with_suffix(input_path.suffix)
    final_output.parent.mkdir(parents=True, exist_ok=True)
    if final_output.resolve() != input_path.resolve():
        shutil.copyfile(input_path, final_output)
    return CompressionResult(
        input_path=input_path,
        output_path=final_output,
        original_size=original_size,
        compressed_size=original_size,
        original_dims=original_dims,
        final_dims=original_dims,
        output_format="JPEG",
        was_skipped=True,
    )


def _has_alpha(img: Image.Image) -> bool:
    if img.mode in ("RGBA", "LA"):
        return True
    if img.mode == "P" and "transparency" in img.info:
        return True
    return False


def _should_skip(
    ext: str,
    size: int,
    dims: tuple[int, int],
    max_dim: int,
) -> bool:
    if ext not in JPEG_EXTENSIONS:
        return False
    if size >= SKIP_THRESHOLD_BYTES:
        return False
    if max(dims) > max_dim:
        return False
    return True


def _resize_if_needed(img: Image.Image, max_dim: int) -> Image.Image:
    w, h = img.size
    longest = max(w, h)
    if longest <= max_dim:
        return img
    ratio = max_dim / longest
    new_size = (max(1, int(round(w * ratio))), max(1, int(round(h * ratio))))
    return img.resize(new_size, Image.Resampling.LANCZOS)


def _flatten_to_white(img: Image.Image) -> Image.Image:
    if img.mode == "P":
        img = img.convert("RGBA")
    if img.mode == "RGBA":
        background = Image.new("RGB", img.size, (255, 255, 255))
        background.paste(img, mask=img.split()[3])
        return background
    if img.mode == "LA":
        return img.convert("RGB")
    return img.convert("RGB")
