import logging
import os
import sys

# Resolve o diretório base: pasta do .exe (PyInstaller frozen) ou do script
if getattr(sys, "frozen", False):
    _BASE_DIR = os.path.dirname(sys.executable)
else:
    _BASE_DIR = os.path.dirname(os.path.abspath(__file__))
_LOG_DIR = os.path.join(_BASE_DIR, "logs")
_LOG_FILE = os.path.join(_LOG_DIR, "migracao.log")

os.makedirs(_LOG_DIR, exist_ok=True)

logger = logging.getLogger("eproc_migrador")

if not logger.handlers:
    logger.setLevel(logging.INFO)

    _fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s")

    # Arquivo — modo append para acumular execuções de dias diferentes
    _fh = logging.FileHandler(_LOG_FILE, mode="a", encoding="utf-8")
    _fh.setFormatter(_fmt)
    logger.addHandler(_fh)

    # Console — exibe o mesmo conteúdo no terminal
    _ch = logging.StreamHandler()
    _ch.setFormatter(_fmt)
    logger.addHandler(_ch)
