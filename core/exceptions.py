"""Exceções específicas do núcleo de compressão."""


class CompressionError(Exception):
    """Erro base para falhas de compressão."""


class CorruptImageError(CompressionError):
    """Imagem ilegível ou corrompida."""


class LegacyDocFormatError(CompressionError):
    """Arquivo `.doc` antigo recebido por uma rotina que só aceita `.docx`."""


class WordNotAvailableError(CompressionError):
    """Microsoft Word não está disponível para converter `.doc`."""


class InvalidDocxError(CompressionError):
    """Arquivo `.docx` corrompido ou que não é um ZIP válido."""


class UnsupportedFormatError(CompressionError):
    """Formato de imagem não suportado para compressão."""
