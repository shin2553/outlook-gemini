import configparser
import sys
from pathlib import Path


def get_app_dir() -> Path:
    """실행 파일 기준 디렉토리 반환. PyInstaller .exe와 일반 스크립트 모두 호환."""
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).parent
    return Path(__file__).parent


_CONFIG_PATH = get_app_dir() / "config.ini"

_config = configparser.ConfigParser()
_config.read(_CONFIG_PATH, encoding="utf-8")

def _resolve(raw: str) -> str:
    """상대 경로는 app_dir 기준 절대 경로로 변환."""
    if not raw:
        return raw
    p = Path(raw)
    if not p.is_absolute():
        p = get_app_dir() / p
    return str(p)


GEMINI_API_KEY = _config.get("gemini", "api_key", fallback="")
MANUAL_INDEX_PATH = _resolve(_config.get("paths", "manual_index", fallback=""))
MANUAL_DIR = _resolve(_config.get("paths", "manual_dir", fallback=""))
PDF_DIR = _resolve(_config.get("paths", "pdf_dir", fallback=""))
MAX_MANUAL_CONTEXT = int(_config.get("settings", "max_manual_context", fallback="8000"))
LANGUAGE = _config.get("settings", "language", fallback="ko")

COMPANY_ROLE = _config.get("company", "role", fallback="")
