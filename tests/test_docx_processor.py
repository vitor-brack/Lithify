"""Testes do core.docx_processor."""
from __future__ import annotations

import io
import secrets
import zipfile
from pathlib import Path

import pytest
from PIL import Image

from core.docx_processor import comprimir_docx
from core.exceptions import InvalidDocxError, LegacyDocFormatError


CONTENT_TYPES = b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="jpg" ContentType="image/jpeg"/>
  <Default Extension="png" ContentType="image/png"/>
  <Default Extension="emf" ContentType="image/x-emf"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>
"""

ROOT_RELS = b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"""

DOCUMENT_XML = b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body><w:p><w:r><w:t>Laudo de teste</w:t></w:r></w:p></w:body>
</w:document>
"""

DOCUMENT_RELS = b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.jpg"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image2.png"/>
</Relationships>
"""


def _noise_jpeg_bytes(size: tuple[int, int], quality: int = 95) -> bytes:
    raw = secrets.token_bytes(size[0] * size[1] * 3)
    img = Image.frombytes("RGB", size, raw)
    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=quality)
    return buf.getvalue()


def _png_with_alpha_bytes(size: tuple[int, int]) -> bytes:
    img = Image.new("RGBA", size, (200, 50, 50, 200))
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()


def _make_docx(
    path: Path,
    media: dict[str, bytes] | None = None,
) -> Path:
    """Monta um .docx mínimo. `media` mapeia nomes (ex.: 'image1.jpg') -> bytes."""
    media = media or {}
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES)
        z.writestr("_rels/.rels", ROOT_RELS)
        z.writestr("word/document.xml", DOCUMENT_XML)
        z.writestr("word/_rels/document.xml.rels", DOCUMENT_RELS)
        for name, data in media.items():
            z.writestr(f"word/media/{name}", data)
    return path


def test_docx_com_imagens_eh_comprimido_e_estrutura_preservada(tmp_path: Path):
    img1 = _noise_jpeg_bytes((4000, 3000))
    img2 = _png_with_alpha_bytes((1500, 1500))
    src = _make_docx(
        tmp_path / "laudo.docx",
        media={"image1.jpg": img1, "image2.png": img2},
    )
    original_bytes = src.read_bytes()
    out = tmp_path / "comprimidos" / "laudo.docx"

    result = comprimir_docx(src, out, max_dim=2000, quality=82)

    assert out.exists()
    assert result.compressed_size < result.original_size
    assert result.images_processed >= 1
    assert result.images_failed == 0

    assert src.read_bytes() == original_bytes

    with zipfile.ZipFile(out, "r") as zout:
        names = set(zout.namelist())
        assert "[Content_Types].xml" in names
        assert "_rels/.rels" in names
        assert "word/document.xml" in names
        assert "word/_rels/document.xml.rels" in names
        assert "word/media/image1.jpg" in names
        assert "word/media/image2.png" in names

        assert zout.read("word/document.xml") == DOCUMENT_XML
        assert zout.read("[Content_Types].xml") == CONTENT_TYPES
        assert zout.read("_rels/.rels") == ROOT_RELS
        assert zout.read("word/_rels/document.xml.rels") == DOCUMENT_RELS

        new_img1 = zout.read("word/media/image1.jpg")
        assert len(new_img1) < len(img1)


def test_doc_antigo_por_extensao_lanca_legacy(tmp_path: Path):
    src = tmp_path / "antigo.doc"
    src.write_bytes(b"\xd0\xcf\x11\xe0" + b"\x00" * 100)
    out = tmp_path / "out.docx"

    with pytest.raises(LegacyDocFormatError):
        comprimir_docx(src, out)


def test_doc_antigo_por_magic_bytes_lanca_legacy(tmp_path: Path):
    src = tmp_path / "renamed.docx"
    src.write_bytes(b"\xd0\xcf\x11\xe0" + b"\x00" * 100)
    out = tmp_path / "out.docx"

    with pytest.raises(LegacyDocFormatError):
        comprimir_docx(src, out)


def test_arquivo_nao_zip_lanca_invaliddocx(tmp_path: Path):
    src = tmp_path / "fake.docx"
    src.write_bytes(b"este nao e um zip")
    out = tmp_path / "out.docx"

    with pytest.raises(InvalidDocxError):
        comprimir_docx(src, out)


def test_zip_corrompido_lanca_invaliddocx(tmp_path: Path):
    src = tmp_path / "broken.docx"
    valid = _make_docx(tmp_path / "tmp.docx", media={})
    data = valid.read_bytes()
    src.write_bytes(data[:50] + b"\x00" * 200)
    out = tmp_path / "out.docx"

    with pytest.raises(InvalidDocxError):
        comprimir_docx(src, out)


def test_docx_sem_imagens_resulta_em_estrutura_identica(tmp_path: Path):
    src = _make_docx(tmp_path / "vazio.docx", media={})
    out = tmp_path / "out" / "vazio.docx"

    result = comprimir_docx(src, out)

    assert result.images_processed == 0
    assert result.images_skipped_already_optimized == 0
    assert result.images_skipped_non_raster == 0

    with zipfile.ZipFile(src, "r") as zin, zipfile.ZipFile(out, "r") as zout:
        in_names = set(zin.namelist())
        out_names = set(zout.namelist())
        assert in_names == out_names
        for name in in_names:
            assert zin.read(name) == zout.read(name)


def test_formatos_vetoriais_sao_pulados(tmp_path: Path):
    img = _noise_jpeg_bytes((3000, 2000))
    src = tmp_path / "com_emf.docx"
    with zipfile.ZipFile(src, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES)
        z.writestr("_rels/.rels", ROOT_RELS)
        z.writestr("word/document.xml", DOCUMENT_XML)
        z.writestr("word/_rels/document.xml.rels", DOCUMENT_RELS)
        z.writestr("word/media/image1.jpg", img)
        z.writestr("word/media/image2.emf", b"FAKE EMF DATA" * 100)

    out = tmp_path / "out.docx"
    result = comprimir_docx(src, out, max_dim=2000)

    assert result.images_skipped_non_raster == 1
    assert result.images_processed == 1

    with zipfile.ZipFile(out, "r") as zout:
        assert zout.read("word/media/image2.emf") == b"FAKE EMF DATA" * 100


def test_progress_callback_eh_chamado_para_cada_raster(tmp_path: Path):
    src = _make_docx(
        tmp_path / "lote.docx",
        media={
            "image1.jpg": _noise_jpeg_bytes((3000, 2000)),
            "image2.jpg": _noise_jpeg_bytes((2500, 2000)),
            "image3.png": _png_with_alpha_bytes((1500, 1000)),
            "vetor.emf": b"vetor",
        },
    )
    out = tmp_path / "out.docx"

    progress: list[tuple[int, int, str]] = []
    comprimir_docx(src, out, progress_callback=lambda c, t, n: progress.append((c, t, n)))

    assert len(progress) == 3
    assert all(t == 3 for _, t, _ in progress)
    assert [c for c, _, _ in progress] == [1, 2, 3]


def test_output_igual_input_levanta_valueerror(tmp_path: Path):
    src = _make_docx(tmp_path / "x.docx", media={})

    with pytest.raises(ValueError):
        comprimir_docx(src, src)


def test_arquivo_inexistente_lanca_filenotfounderror(tmp_path: Path):
    src = tmp_path / "nao_existe.docx"
    out = tmp_path / "out.docx"

    with pytest.raises(FileNotFoundError):
        comprimir_docx(src, out)


def test_imagem_corrompida_dentro_do_docx_eh_contada_como_failed(tmp_path: Path):
    src = tmp_path / "com_lixo.docx"
    with zipfile.ZipFile(src, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES)
        z.writestr("_rels/.rels", ROOT_RELS)
        z.writestr("word/document.xml", DOCUMENT_XML)
        z.writestr("word/_rels/document.xml.rels", DOCUMENT_RELS)
        z.writestr("word/media/lixo.jpg", b"isto nao e um jpeg" * 50)

    out = tmp_path / "out.docx"
    result = comprimir_docx(src, out)

    assert result.images_failed == 1
    assert result.images_processed == 0
    with zipfile.ZipFile(out, "r") as zout:
        assert zout.read("word/media/lixo.jpg") == b"isto nao e um jpeg" * 50
