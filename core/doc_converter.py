"""Conversão entre `.doc` e `.docx` via automação MS Word (Windows-only).

Este módulo isola toda a interação com o Word/COM. Se `pywin32` não estiver
disponível ou Word não estiver instalado, `is_word_available()` retorna False
e os usuários devem cair no fluxo de erro claro em vez de tentar converter.
"""
from __future__ import annotations

import logging
from contextlib import contextmanager
from pathlib import Path
from typing import Iterator, Optional

from core.exceptions import WordNotAvailableError

logger = logging.getLogger(__name__)


try:
    import pythoncom
    import win32com.client

    _HAS_WIN32COM = True
except ImportError:
    pythoncom = None  # type: ignore[assignment]
    win32com = None  # type: ignore[assignment]
    _HAS_WIN32COM = False


# Word FileFormat constants (WdSaveFormat enum)
_WD_FORMAT_DOCUMENT = 0  # .doc legado
_WD_FORMAT_DOCX = 12  # wdFormatXMLDocument

_word_available_cache: Optional[bool] = None


def is_word_available() -> bool:
    """True se for possível instanciar o Word via COM. Resultado é cacheado."""
    global _word_available_cache
    if _word_available_cache is not None:
        return _word_available_cache
    _word_available_cache = _check_word_available()
    return _word_available_cache


def reset_word_available_cache() -> None:
    """Limpa o cache (útil para testes)."""
    global _word_available_cache
    _word_available_cache = None


def _check_word_available() -> bool:
    if not _HAS_WIN32COM:
        logger.debug("pywin32 não está disponível.")
        return False
    try:
        pythoncom.CoInitialize()
        try:
            word = win32com.client.DispatchEx("Word.Application")
            word.Quit()
            return True
        finally:
            pythoncom.CoUninitialize()
    except Exception:
        logger.debug("Word não acessível via COM.", exc_info=True)
        return False


@contextmanager
def word_session() -> Iterator["WordSession"]:
    """Abre uma sessão Word, garante teardown.

    Use o objeto WordSession retornado para chamar `doc_to_docx` / `docx_to_doc`
    múltiplas vezes sem reabrir o Word.
    """
    if not _HAS_WIN32COM:
        raise WordNotAvailableError(
            "pywin32 não está instalado. O suporte a .doc requer Microsoft Word."
        )
    pythoncom.CoInitialize()
    word = None
    try:
        try:
            word = win32com.client.DispatchEx("Word.Application")
        except Exception as exc:
            raise WordNotAvailableError(
                "Não foi possível abrir o Microsoft Word. Verifique se está instalado."
            ) from exc
        word.Visible = False
        word.DisplayAlerts = 0
        try:
            word.AutomationSecurity = 3
        except Exception:
            pass
        try:
            yield WordSession(word)
        finally:
            try:
                word.Quit()
            except Exception:
                logger.warning("Falha ao encerrar Word.", exc_info=True)
    finally:
        pythoncom.CoUninitialize()


class WordSession:
    """Wrapper fino sobre uma instância COM do Word."""

    def __init__(self, word_app: object) -> None:
        self._word = word_app

    def doc_to_docx(self, src: Path, dst: Path) -> None:
        """Converte `.doc` em `.docx` no caminho destino."""
        _save_as(self._word, src, dst, _WD_FORMAT_DOCX)

    def docx_to_doc(self, src: Path, dst: Path) -> None:
        """Converte `.docx` em `.doc` no caminho destino."""
        _save_as(self._word, src, dst, _WD_FORMAT_DOCUMENT)


def _save_as(word: object, src: Path, dst: Path, file_format: int) -> None:
    src_str = str(src.resolve())
    dst_str = str(dst.resolve())
    dst.parent.mkdir(parents=True, exist_ok=True)
    if dst.exists():
        dst.unlink()
    logger.debug("Word SaveAs: %s -> %s (fmt=%s)", src_str, dst_str, file_format)
    doc = word.Documents.Open(src_str, ReadOnly=True, ConfirmConversions=False)  # type: ignore[attr-defined]
    try:
        doc.SaveAs2(dst_str, FileFormat=file_format)
    finally:
        doc.Close(SaveChanges=False)
