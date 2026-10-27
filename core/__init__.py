from core.doc_converter import is_word_available
from core.document_processor import comprimir_documento, detect_format
from core.docx_processor import DocxCompressionResult, comprimir_docx
from core.exceptions import (
    CompressionError,
    CorruptImageError,
    InvalidDocxError,
    LegacyDocFormatError,
    UnsupportedFormatError,
    WordNotAvailableError,
)
from core.image_compressor import (
    CompressionResult,
    comprimir_imagem,
    comprimir_lote,
)
from core.ole_processor import (
    OleCompressionResult,
    comprimir_ole_pbrush,
    is_ole_rewrite_available,
)

__all__ = [
    "CompressionError",
    "CorruptImageError",
    "InvalidDocxError",
    "LegacyDocFormatError",
    "UnsupportedFormatError",
    "WordNotAvailableError",
    "CompressionResult",
    "DocxCompressionResult",
    "OleCompressionResult",
    "comprimir_imagem",
    "comprimir_lote",
    "comprimir_docx",
    "comprimir_documento",
    "comprimir_ole_pbrush",
    "detect_format",
    "is_word_available",
    "is_ole_rewrite_available",
]
