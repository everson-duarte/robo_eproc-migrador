import logging
import os

_LOG_DIR = "logs"
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
