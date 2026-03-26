"""
매뉴얼 검색 모듈
기존 manual_index.py의 search_manuals()를 활용하여
Gemini API에 전달할 컨텍스트 텍스트와 PDF 경로를 반환합니다.
"""

import sys
from pathlib import Path

import config


def _find_and_load_manual_index():
    """
    manual_index.py를 자동 탐색하여 search_manuals 함수 반환.
    탐색 순서:
      1. config.MANUAL_INDEX_PATH (명시된 경로)
      2. manual_dir 의 상위 폴더
      3. manual_dir 자체
    """
    candidates = []
    if config.MANUAL_INDEX_PATH:
        candidates.append(Path(config.MANUAL_INDEX_PATH))
    if config.MANUAL_DIR:
        base = Path(config.MANUAL_DIR)
        candidates.append(base.parent / "manual_index.py")
        candidates.append(base / "manual_index.py")

    for p in candidates:
        if p.exists():
            sys.path.insert(0, str(p.parent))
            from manual_index import search_manuals as _sm
            return _sm

    raise FileNotFoundError(
        "manual_index.py를 찾을 수 없습니다.\n"
        "설정 > 매뉴얼 관리에서 txt 매뉴얼 폴더를 올바르게 지정하세요."
    )


search_manuals = _find_and_load_manual_index()


def _find_pdf(txt_path: Path) -> Path | None:
    """
    txt 파일에 대응하는 PDF 원본을 찾습니다.
    탐색 순서:
      1. config.PDF_DIR 폴더 (별도 PDF 보관 폴더)
      2. txt 파일과 같은 폴더 (같은 이름의 .pdf)
    """
    stem = txt_path.stem

    if config.PDF_DIR:
        candidate = Path(config.PDF_DIR) / (stem + ".pdf")
        if candidate.exists():
            return candidate

    candidate = txt_path.parent / (stem + ".pdf")
    if candidate.exists():
        return candidate

    return None


def build_manual_context(query: str, max_chars: int = config.MAX_MANUAL_CONTEXT) -> str:
    """txt 기반 매뉴얼 컨텍스트 문자열 반환."""
    matched_files = search_manuals(query)
    if not matched_files:
        return ""

    sections = []
    total_chars = 0

    for path in matched_files:
        if total_chars >= max_chars:
            break
        try:
            text = path.read_text(encoding="utf-8", errors="ignore")
        except Exception:
            continue

        remaining = max_chars - total_chars
        if len(text) > remaining:
            text = text[:remaining] + "\n... (이하 생략)"

        header = f"### [{path.name}]\n"
        sections.append(header + text)
        total_chars += len(header) + len(text)

    return "\n\n".join(sections)


def get_matched_pdfs(query: str) -> list[Path]:
    """
    매칭된 txt 파일 각각에 대응하는 PDF 원본 경로 목록 반환.
    PDF가 없는 항목은 제외됩니다.
    """
    result = []
    for txt_path in search_manuals(query):
        pdf = _find_pdf(txt_path)
        if pdf:
            result.append(pdf)
    return result


def get_matched_file_names(query: str) -> list[str]:
    """매칭된 매뉴얼 txt 파일명 목록 반환 (UI 표시용)."""
    return [p.name for p in search_manuals(query)]
