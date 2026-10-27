"""Ponto de entrada do Lithify."""
from __future__ import annotations

import argparse
import logging
import os
import sys
from logging.handlers import RotatingFileHandler
from pathlib import Path

from gui.main_window import MainWindow, make_root


APP_NAME = "Lithify"
LOG_FILE_NAME = "lithify.log"
LOG_FORMAT = "%(asctime)s [%(levelname)s] %(name)s: %(message)s"


def _log_dir() -> Path:
    """Resolve a pasta para logs.

    Windows: %APPDATA%/Lithify/. Outros: ~/.lithify/.
    Cai para um subdir ao lado do executável se nada disso funcionar.
    """
    appdata = os.environ.get("APPDATA")
    if appdata:
        return Path(appdata) / APP_NAME
    home = Path.home()
    if home.exists():
        return home / f".{APP_NAME.lower()}"
    return Path(sys.argv[0]).resolve().parent / "logs"


def _configure_logging(debug: bool) -> Path:
    log_dir = _log_dir()
    try:
        log_dir.mkdir(parents=True, exist_ok=True)
    except OSError:
        log_dir = Path.cwd()
    log_path = log_dir / LOG_FILE_NAME

    root = logging.getLogger()
    root.setLevel(logging.DEBUG if debug else logging.INFO)
    for h in list(root.handlers):
        root.removeHandler(h)

    file_handler = RotatingFileHandler(
        log_path,
        maxBytes=1_000_000,
        backupCount=3,
        encoding="utf-8",
    )
    file_handler.setLevel(logging.DEBUG if debug else logging.INFO)
    file_handler.setFormatter(logging.Formatter(LOG_FORMAT))
    root.addHandler(file_handler)

    if debug:
        console = logging.StreamHandler()
        console.setLevel(logging.DEBUG)
        console.setFormatter(logging.Formatter(LOG_FORMAT))
        root.addHandler(console)

    logging.getLogger("PIL").setLevel(logging.INFO)
    return log_path


def _parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(prog=APP_NAME, description="Compressor de imagens para laudos.")
    parser.add_argument("--debug", action="store_true", help="Habilita logging detalhado no console.")
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = _parse_args(argv if argv is not None else sys.argv[1:])
    log_path = _configure_logging(debug=args.debug)
    logger = logging.getLogger(__name__)
    logger.info("Iniciando %s (log em %s)", APP_NAME, log_path)

    try:
        root = make_root()
        MainWindow(root)
        root.mainloop()
    except Exception:
        logger.exception("Erro fatal não tratado")
        return 1

    logger.info("Encerrando %s", APP_NAME)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
