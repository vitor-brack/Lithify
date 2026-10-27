"""Testes do core.image_compressor."""
from __future__ import annotations

from pathlib import Path

import pytest
from PIL import Image

from core.exceptions import CorruptImageError, UnsupportedFormatError
from core.image_compressor import (
    SKIP_THRESHOLD_BYTES,
    comprimir_imagem,
    comprimir_lote,
)


def _make_jpeg(
    path: Path,
    size: tuple[int, int],
    color: tuple[int, int, int] = (120, 80, 200),
    quality: int = 95,
) -> Path:
    img = Image.new("RGB", size, color)
    img.save(path, format="JPEG", quality=quality)
    return path


def _make_noise_jpeg(path: Path, size: tuple[int, int], quality: int = 95) -> Path:
    """JPEG com conteúdo realista (ruído) para garantir tamanho > threshold de skip."""
    import secrets

    raw = secrets.token_bytes(size[0] * size[1] * 3)
    img = Image.frombytes("RGB", size, raw)
    img.save(path, format="JPEG", quality=quality)
    return path


def _make_png_with_alpha(path: Path, size: tuple[int, int]) -> Path:
    img = Image.new("RGBA", size, (255, 0, 0, 128))
    img.save(path, format="PNG")
    return path


def test_jpeg_grande_eh_redimensionado_e_comprimido(tmp_path: Path):
    src = _make_noise_jpeg(tmp_path / "big.jpg", (4000, 3000))
    out = tmp_path / "out.jpg"

    result = comprimir_imagem(src, out, max_dim=2000, quality=85)

    assert result.was_skipped is False
    assert result.compressed_size < result.original_size
    assert max(result.final_dims) == 2000
    assert result.output_path.exists()
    assert result.output_path.suffix == ".jpg"
    assert result.output_format == "JPEG"


def test_png_com_alpha_vira_jpeg_com_fundo_branco(tmp_path: Path):
    src = _make_png_with_alpha(tmp_path / "alpha.png", (1000, 1000))
    out = tmp_path / "out.png"

    result = comprimir_imagem(src, out, preserve_alpha=False)

    assert result.output_format == "JPEG"
    assert result.output_path.suffix == ".jpg"
    with Image.open(result.output_path) as img:
        assert img.mode == "RGB"


def test_png_com_alpha_preservado(tmp_path: Path):
    src = _make_png_with_alpha(tmp_path / "alpha.png", (1000, 1000))
    out = tmp_path / "out.png"

    result = comprimir_imagem(src, out, preserve_alpha=True)

    assert result.output_format == "PNG"
    assert result.output_path.suffix == ".png"
    with Image.open(result.output_path) as img:
        assert img.mode in ("RGBA", "LA", "P")


def test_jpeg_pequeno_eh_skipped_e_copiado_byte_a_byte(tmp_path: Path):
    src = _make_jpeg(tmp_path / "small.jpg", (800, 600), quality=85)
    assert src.stat().st_size < SKIP_THRESHOLD_BYTES
    out = tmp_path / "out.jpg"

    result = comprimir_imagem(src, out, max_dim=2000, quality=85)

    assert result.was_skipped is True
    assert result.original_size == result.compressed_size
    assert result.output_path.read_bytes() == src.read_bytes()


def test_arquivo_corrompido_lanca_corruptimage(tmp_path: Path):
    src = tmp_path / "corrupted.jpg"
    src.write_bytes(b"not a real jpeg, just some random bytes" * 10)
    out = tmp_path / "out.jpg"

    with pytest.raises(CorruptImageError):
        comprimir_imagem(src, out)


def test_formato_nao_suportado_lanca_excecao(tmp_path: Path):
    src = tmp_path / "doc.pdf"
    src.write_bytes(b"%PDF-1.4")
    out = tmp_path / "out.jpg"

    with pytest.raises(UnsupportedFormatError):
        comprimir_imagem(src, out)


def test_arquivo_inexistente_lanca_filenotfounderror(tmp_path: Path):
    src = tmp_path / "ghost.jpg"
    out = tmp_path / "out.jpg"

    with pytest.raises(FileNotFoundError):
        comprimir_imagem(src, out)


def test_imagem_menor_que_max_dim_nao_eh_upscaled(tmp_path: Path):
    src = _make_noise_jpeg(tmp_path / "medium.jpg", (1500, 1000))
    out = tmp_path / "out.jpg"

    result = comprimir_imagem(src, out, max_dim=3000)

    assert result.final_dims == (1500, 1000)


def test_extensao_de_saida_eh_corrigida_para_jpg(tmp_path: Path):
    """Mesmo se chamarem com '.png' no output_path, conversão para JPEG ajusta a extensão."""
    src = _make_png_with_alpha(tmp_path / "in.png", (1200, 900))
    out = tmp_path / "saida.png"

    result = comprimir_imagem(src, out, preserve_alpha=False)

    assert result.output_path == tmp_path / "saida.jpg"
    assert result.output_path.exists()
    assert not (tmp_path / "saida.png").exists() or result.output_path != (tmp_path / "saida.png")


def test_metadados_exif_sao_removidos(tmp_path: Path):
    """Saída JPEG não deve carregar EXIF herdado."""
    from PIL import Image as PilImage

    src = tmp_path / "with_exif.jpg"
    img = PilImage.new("RGB", (3000, 2000), (50, 50, 50))
    exif = img.getexif()
    exif[0x010F] = "FakeCamera"
    img.save(src, format="JPEG", exif=exif, quality=95)
    out = tmp_path / "out.jpg"

    result = comprimir_imagem(src, out, max_dim=2000)

    with PilImage.open(result.output_path) as out_img:
        out_exif = out_img.getexif()
        assert 0x010F not in out_exif


def test_comprimir_lote_continua_apos_falha_individual(tmp_path: Path):
    good = _make_noise_jpeg(tmp_path / "good.jpg", (3000, 2000))
    bad = tmp_path / "bad.jpg"
    bad.write_bytes(b"corrupted bytes")

    output_dir = tmp_path / "out"
    progress: list[tuple[int, int, str]] = []

    def cb(cur: int, total: int, name: str) -> None:
        progress.append((cur, total, name))

    results = comprimir_lote([good, bad], output_dir, progress_callback=cb)

    assert len(results) == 1
    assert results[0].input_path == good
    assert len(progress) == 2
    assert progress[0] == (1, 2, "good.jpg")
    assert progress[1] == (2, 2, "bad.jpg")
