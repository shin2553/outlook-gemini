"""
PDF 페이지를 이미지(PNG)로 추출
pymupdf(fitz) 사용: pip install pymupdf
"""

import tempfile
from pathlib import Path


def extract_page_as_png(pdf_path: Path, page_num: int, dpi: int = 150) -> Path:
    """
    PDF의 특정 페이지를 PNG 임시 파일로 렌더링하여 경로 반환.
    page_num: 1-based
    """
    try:
        import fitz
    except ImportError:
        raise ImportError("pymupdf 패키지가 필요합니다: pip install pymupdf")

    doc = fitz.open(str(pdf_path))
    total = len(doc)
    if page_num < 1 or page_num > total:
        doc.close()
        raise ValueError(f"페이지 번호 범위 초과: {page_num} (총 {total}페이지)")

    page = doc[page_num - 1]
    mat = fitz.Matrix(dpi / 72, dpi / 72)
    pix = page.get_pixmap(matrix=mat)

    stem = Path(pdf_path).stem
    tmp = tempfile.NamedTemporaryFile(
        suffix=f"_{stem}_p{page_num}.png",
        delete=False,
        dir=tempfile.gettempdir(),
    )
    tmp_path = Path(tmp.name)
    tmp.close()  # Windows: 핸들 닫은 후 저장해야 권한 오류 없음
    pix.save(str(tmp_path))
    doc.close()
    return tmp_path


def find_pdf(filename: str, search_dirs: list[Path]) -> Path | None:
    """
    파일명으로 PDF를 지정된 폴더들에서 탐색.
    1) 정확한 파일명
    2) stem + .pdf / .txt
    3) 핵심 키워드가 모두 포함된 PDF (대소문자 무시)
    """
    stem = Path(filename).stem
    exact_names = [Path(filename).name, stem + ".pdf", stem + ".txt"]

    for d in search_dirs:
        if not d.exists():
            continue

        # 정확한 이름 우선
        for name in exact_names:
            candidate = d / name
            if candidate.exists():
                return candidate

        # 핵심 키워드 추출 (2자 이상 단어)
        import re as _re
        keywords = [w.lower() for w in _re.split(r'[\s_\-\.]+', stem) if len(w) >= 2]

        # 스코어: 키워드 매칭 수 (동점 시 stem 길이 짧은 것 우선 — 더 전용 파일)
        candidates = []
        for pdf in d.glob("*.pdf"):
            pdf_lower = pdf.stem.lower()
            score = sum(1 for kw in keywords if kw in pdf_lower)
            if score > 0:
                candidates.append((score, len(pdf.stem), pdf))

        candidates.sort(key=lambda x: (-x[0], x[1]))

        # 키워드 절반 이상 매칭된 경우만 반환
        if candidates and candidates[0][0] >= max(1, len(keywords) // 2):
            return candidates[0][2]

    return None
