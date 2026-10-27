"""Testes do core.document_processor (orquestrador .doc + .docx)."""
from __future__ import annotations

import io
import secrets
import zipfile
from pathlib import Path

import pytest
from PIL import Image

from core import doc_converter, document_processor
from core.docx_processor import DocxCompressionResult
from core.exceptions import InvalidDocxError, WordNotAvailableError


CONTENT_TYPES = b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="jpg" ContentType="image/jpeg"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>
"""

ROOT_RELS = b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"""

DOCUMENT_XML = b"""<?xml version="1.0"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body><w:p><w:r><w:t>x</w:t></w:r></w:p></w:body>
</w:document>
"""

DOCUMENT_RELS = b"""<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.jpg"/>
</Relationships>
"""


def _noise_jpeg_bytes(size: tuple[int, int]) -> bytes:
    raw = secrets.token_bytes(size[0] * size[1] * 3)
    img = Image.frombytes("RGB", size, raw)
    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=95)
    return buf.getvalue()


def _make_docx(path: Path, with_image: bool = True) -> Path:
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES)
        z.writestr("_rels/.rels", ROOT_RELS)
        z.writestr("word/document.xml", DOCUMENT_XML)
        z.writestr("word/_rels/document.xml.rels", DOCUMENT_RELS)
        if with_image:
            z.writestr("word/media/image1.jpg", _noise_jpeg_bytes((3000, 2000)))
    return path


def _fake_legacy_doc(path: Path) -> Path:
    path.write_bytes(b"\xd0\xcf\x11\xe0" + b"\x00" * 200)
    return path


@pytest.fixture(autouse=True)
def _reset_word_cache():
    doc_converter.reset_word_available_cache()
    yield
    doc_converter.reset_word_available_cache()


def test_detect_format_docx(tmp_path: Path):
    src = _make_docx(tmp_path / "ok.docx")
    assert document_processor.detect_format(src) == "docx"


def test_detect_format_doc_via_magic(tmp_path: Path):
    src = _fake_legacy_doc(tmp_path / "antigo.doc")
    assert document_processor.detect_format(src) == "doc"


def test_detect_format_doc_via_extensao_quando_magic_indefinido(tmp_path: Path):
    src = tmp_path / "estranho.doc"
    src.write_bytes(b"alguma coisa que nao e ole nem zip")
    assert document_processor.detect_format(src) == "doc"


def test_detect_format_invalido(tmp_path: Path):
    src = tmp_path / "qualquer.bin"
    src.write_bytes(b"lixo aleatorio sem extensao reconhecida")
    with pytest.raises(InvalidDocxError):
        document_processor.detect_format(src)


def test_detect_format_arquivo_inexistente(tmp_path: Path):
    with pytest.raises(FileNotFoundError):
        document_processor.detect_format(tmp_path / "ghost.docx")


def test_comprimir_documento_despacha_para_docx(tmp_path: Path):
    src = _make_docx(tmp_path / "x.docx")
    out = tmp_path / "out" / "x.docx"

    result = document_processor.comprimir_documento(src, out, max_dim=2000, quality=82)

    assert isinstance(result, DocxCompressionResult)
    assert out.exists()
    assert result.images_processed >= 1


def test_comprimir_documento_chama_stage_callback_para_docx(tmp_path: Path):
    src = _make_docx(tmp_path / "x.docx")
    out = tmp_path / "out.docx"
    stages: list[str] = []

    document_processor.comprimir_documento(
        src, out, stage_callback=stages.append
    )

    assert any("Comprimindo" in s for s in stages)


def test_comprimir_documento_doc_sem_word_disponivel_falha_clara(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
):
    src = _fake_legacy_doc(tmp_path / "antigo.doc")
    out = tmp_path / "out.doc"

    monkeypatch.setattr(document_processor, "is_word_available", lambda: False)

    with pytest.raises(WordNotAvailableError) as exc_info:
        document_processor.comprimir_documento(src, out)

    msg = str(exc_info.value)
    assert "Word" in msg
    assert ".docx" in msg


def test_output_igual_input_levanta_valueerror(tmp_path: Path):
    src = _make_docx(tmp_path / "x.docx")
    with pytest.raises(ValueError):
        document_processor.comprimir_documento(src, src)


@pytest.mark.skipif(
    not doc_converter.is_word_available(),
    reason="Microsoft Word não disponível neste ambiente",
)
def test_integracao_doc_round_trip_produz_doc_valido(tmp_path: Path):
    """Integração: pipeline .doc → .docx → compressão → .doc roda end-to-end e produz um .doc com magic bytes corretos.

    Não assertimos sobre `images_processed` porque o Word descarta imagens
    não referenciadas em `document.xml` durante o round-trip do fixture mínimo.
    Em arquivos reais com imagens referenciadas no corpo, a fidelidade é preservada.
    Pulado automaticamente em ambientes sem Word.
    """
    intermediate_docx = _make_docx(tmp_path / "fonte.docx")
    src_doc = tmp_path / "fonte.doc"
    out_doc = tmp_path / "comprimidos" / "fonte.doc"

    with doc_converter.word_session() as session:
        session.docx_to_doc(intermediate_docx, src_doc)

    assert src_doc.exists()

    result = document_processor.comprimir_documento(
        src_doc, out_doc, max_dim=2000, quality=82
    )

    assert out_doc.exists()
    assert out_doc.stat().st_size > 0
    assert isinstance(result, DocxCompressionResult)
    assert result.input_path == src_doc
    assert result.output_path == out_doc

    with open(out_doc, "rb") as f:
        magic = f.read(4)
    assert magic == b"\xd0\xcf\x11\xe0"
