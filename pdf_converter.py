"""
PDF → txt 변환 모듈
pypdf를 사용해 텍스트를 추출하고 매뉴얼 폴더에 저장합니다.
스캔본(이미지 PDF)은 텍스트 추출이 안 될 수 있습니다.
"""

from pathlib import Path
from pypdf import PdfReader


def convert_pdf_to_txt(pdf_path: Path, output_dir: Path) -> tuple[Path, int]:
    """
    PDF 파일을 텍스트로 변환하여 output_dir에 저장합니다.

    Returns:
        (저장된 txt 경로, 추출된 페이지 수)

    Raises:
        ValueError: 텍스트를 추출할 수 없는 경우 (스캔본 등)
    """
    reader = PdfReader(str(pdf_path))
    total_pages = len(reader.pages)

    lines = []
    extracted = 0
    for i, page in enumerate(reader.pages):
        try:
            text = page.extract_text() or ""
        except Exception:
            text = ""
        if text.strip():
            lines.append(f"--- Page {i + 1} ---")
            lines.append(text.strip())
            extracted += 1

    if extracted == 0:
        raise ValueError(
            f"텍스트를 추출할 수 없습니다.\n"
            f"스캔된 이미지 PDF이거나 보안이 걸린 파일일 수 있습니다.\n"
            f"(전체 {total_pages}페이지 중 0페이지 추출)"
        )

    content = "\n\n".join(lines)
    out_path = output_dir / (pdf_path.stem + ".txt")
    out_path.write_text(content, encoding="utf-8")

    return out_path, extracted


if __name__ == "__main__":
    import sys
    if len(sys.argv) < 3:
        print("사용법: python pdf_converter.py <pdf경로> <출력폴더>")
        sys.exit(1)
    pdf = Path(sys.argv[1])
    out = Path(sys.argv[2])
    result, pages = convert_pdf_to_txt(pdf, out)
    print(f"변환 완료: {result} ({pages}페이지)")
