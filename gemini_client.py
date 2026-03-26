"""
Gemini API 클라이언트
- 텍스트, 이미지, PDF 멀티모달 지원
- 2단계 분석: txt 우선 → 시각 정보 필요 시 PDF 원본으로 재답변
- 응답 구조: [분석] + [이메일 초안] 분리
"""

import mimetypes
import re
from dataclasses import dataclass
from pathlib import Path

from google import genai
from google.genai import types

import config

_client = genai.Client(api_key=config.GEMINI_API_KEY)
MODEL = "gemini-2.5-flash"

_VISUAL_NEEDED_MARKER = "VISUAL_NEEDED"
_ANALYSIS_SEP = "===분석==="
_DRAFT_SEP    = "===이메일초안==="
_REF_SEP      = "===페이지참조==="


@dataclass
class ReplyResult:
    analysis: str        # 문의 분석 + 근거
    draft: str           # 이메일 초안
    page_refs: list[tuple[str, int]] = None  # [(파일명, 페이지번호), ...]
    used_pdf: bool = False  # PDF 원본으로 분석했는지 여부

    def __post_init__(self):
        if self.page_refs is None:
            self.page_refs = []


def _build_system_prompt(responder_company: str = "", responder_department: str = "") -> str:
    company = responder_company
    dept    = responder_department
    role    = config.COMPANY_ROLE

    if company:
        identity = f"당신은 {company} {dept} 소속 기술 지원 담당자입니다. {company}는 {role}입니다.".strip()
    else:
        identity = f"당신은 {role} 기술 지원 담당자입니다." if role else "당신은 기술 지원 담당자입니다."

    return f"""{identity}
고객의 기술 문의에 대해 매뉴얼을 참고하여 분석하고 답변 초안을 작성합니다.

답변 원칙:
- 한국어, 정중한 비즈니스 어체
- 고객에게 그대로 발송 가능한 이메일 형태
- 제공된 매뉴얼에서 해당 제조사를 파악하여 자연스럽게 활용
- 발신자는 {company or '당사'} 소속임을 유지
- 코드 예시는 주석 포함 실제 코드 형태
- 관련 상위 옵션(SDC 등)이 있으면 함께 비교 제안

정보 출처 표기 원칙 (반드시 준수):
- 매뉴얼에 명확한 근거가 있는 내용: 그대로 서술
- 매뉴얼에 없지만 일반 기술 지식으로 보충한 내용: 문장 뒤에 [추정] 태그 추가
- 매뉴얼에도 없고 확신할 수 없는 내용: [확인 필요] 태그 추가 후 "제조사 기술팀에 추가 확인 예정"으로 표현
- [추정] 또는 [확인 필요] 태그가 붙은 내용은 이메일 초안에서도 반드시 동일하게 표기

반드시 아래 형식을 정확히 지켜서 출력하세요:

{_ANALYSIS_SEP}
**문의 요약**
(고객이 묻는 핵심 내용 1~3줄 요약)

**관련 매뉴얼 및 근거**
(참조한 매뉴얼 파일명, 챕터/페이지, 해당 내용 요약. 매뉴얼 근거가 없는 경우 "해당 내용 매뉴얼 미수록" 명시)

**답변 방향**
(어떤 내용으로 답변할지 핵심 포인트. 매뉴얼 근거 있는 항목과 [추정]/[확인 필요] 항목을 구분하여 기술)

**주의 / 확인 필요 사항**
(불확실하거나 발송 전 반드시 검토가 필요한 부분 명시)

{_DRAFT_SEP}
(고객에게 발송할 이메일 초안 — 인사말부터 서명까지 완성된 형태)
※ 이메일 초안 작성 규칙:
- 마크다운 문법 절대 사용 금지 (**, *, #, `, __ 등 모두 금지)
- 강조가 필요한 경우 꺾쇠 없이 글자 앞뒤로 공백이나 줄바꿈으로 구분
- 번호 목록은 "1.", "2." 형태로, 항목은 "-" 또는 "•" 사용
- 코드는 별도 들여쓰기 없이 줄바꿈 후 그대로 작성
- 실제 비즈니스 이메일을 작성하듯 자연스러운 평문으로 작성

{_REF_SEP}
매뉴얼에서 참조한 페이지 번호를 반드시 아래 형식으로 나열하세요. (참조 페이지가 없는 경우에도 섹션은 유지하고 빈 줄로 남기세요)
파일명은 위 매뉴얼 헤더(### [...])에 표시된 이름을 그대로 사용하세요.
[REF] 파일명 | 페이지번호
예시:
[REF] RTC6_Manual.txt | 23
[REF] RTC6_Manual.txt | 45
"""


def _make_part(path: Path) -> types.Part:
    mime, _ = mimetypes.guess_type(str(path))
    if mime is None:
        mime = "application/pdf" if path.suffix.lower() == ".pdf" else "image/jpeg"
    return types.Part.from_bytes(data=path.read_bytes(), mime_type=mime)


def _call(contents: list, system_prompt: str) -> str:
    response = _client.models.generate_content(
        model=MODEL,
        contents=contents,
        config=types.GenerateContentConfig(
            system_instruction=system_prompt,
            thinking_config=types.ThinkingConfig(thinking_budget=10000),
        ),
    )
    return response.text


def _parse(raw: str, used_pdf: bool = False) -> ReplyResult:
    """Gemini 응답을 분석/초안/페이지참조로 분리."""
    analysis = ""
    draft = ""
    refs_text = ""

    if _ANALYSIS_SEP in raw and _DRAFT_SEP in raw:
        after_analysis = raw.split(_ANALYSIS_SEP, 1)[1]
        parts = after_analysis.split(_DRAFT_SEP, 1)
        analysis = parts[0].strip()
        remainder = parts[1].strip() if len(parts) > 1 else ""
    elif _DRAFT_SEP in raw:
        parts = raw.split(_DRAFT_SEP, 1)
        analysis = parts[0].strip()
        remainder = parts[1].strip()
    else:
        return ReplyResult(analysis="", draft=raw.strip())

    # 페이지 참조 섹션 분리
    if _REF_SEP in remainder:
        draft_part, refs_text = remainder.split(_REF_SEP, 1)
        draft = draft_part.strip()
    else:
        draft = remainder

    return ReplyResult(analysis=analysis, draft=draft, page_refs=parse_page_refs(refs_text), used_pdf=used_pdf)


def parse_page_refs(text: str) -> list[tuple[str, int]]:
    """[REF] 파일명.pdf | 페이지번호 형식 파싱."""
    refs = []
    for m in re.finditer(r'\[REF\]\s+(.+?)\s*\|\s*(\d+)', text):
        refs.append((m.group(1).strip(), int(m.group(2))))
    return refs


def generate_reply(
    mail_body: str,
    manual_context: str,
    attachments: list[Path] | None = None,
    manual_pdfs: list[Path] | None = None,
    responder_name: str = "",
    responder_title: str = "",
    responder_company: str = "",
    responder_department: str = "",
    mode: str = "auto",
    extra_instructions: str = "",
) -> ReplyResult:
    """
    Gemini API로 문의 분석 + 이메일 초안 생성.

    mode:
      "auto"      - txt 우선, VISUAL_NEEDED 시 PDF로 자동 전환
      "text_only" - 항상 txt만 사용 (빠름, 이미지 미분석)
      "pdf_only"  - 처음부터 PDF 원본 직접 분석 (정확, 느림)
    """
    system_prompt = _build_system_prompt(responder_company, responder_department)

    # 공통 헤더
    header = ""
    if responder_name:
        header += f"## 답변 작성자\n이름: {responder_name}"
        if responder_title:
            header += f"  /  직책: {responder_title}"
        header += "\n이메일 초안 말미 서명은 위 작성자 정보로 작성하세요.\n\n"

    mail_section = f"## 고객 문의 메일\n{mail_body}\n\n"
    extra_section = (
        f"## 추가 지시사항 (최우선 적용 — 아래 지시사항은 기본 답변 원칙보다 우선합니다)\n"
        f"{extra_instructions}\n\n"
    ) if extra_instructions else ""
    used_pdf = False

    # ── PDF 원본 모드: 바로 2단계 ────────────────────────────
    if mode == "pdf_only" and manual_pdfs:
        hints = _page_hints_from_context(manual_context)
        raw = _call_with_pdf(header, mail_section, manual_pdfs, attachments, system_prompt, hints, extra_section)
        return _parse(raw, used_pdf=True)

    # ── 1단계: txt 기반 ──────────────────────────────────────
    step1_text = header
    if manual_context:
        visual_needed_instruction = (
            "## 중요 지시\n"
            f"텍스트 매뉴얼만으로 도면·회로도·핀배열·타이밍차트 등 시각 정보가 "
            f"반드시 필요한 경우, 답변 대신 딱 한 줄만 출력하세요: {_VISUAL_NEEDED_MARKER}\n"
            "그 외에는 지정된 형식으로 분석과 이메일 초안을 작성하세요.\n\n"
        ) if mode == "auto" else ""
        step1_text += (
            visual_needed_instruction
            + f"## 참고 매뉴얼 (텍스트)\n{manual_context}\n\n"
        )
    step1_text += mail_section + extra_section + "## 위 형식에 맞춰 분석과 이메일 초안을 작성해주세요:"

    contents: list = [step1_text]
    if attachments:
        for att in attachments:
            try:
                contents.append(_make_part(att))
            except Exception as e:
                contents.append(f"[첨부파일 로드 실패: {att.name} - {e}]")

    raw = _call(contents, system_prompt)

    # ── 2단계: VISUAL_NEEDED → PDF 원본 (auto 모드만) ────────
    if mode == "auto" and _VISUAL_NEEDED_MARKER in raw and manual_pdfs:
        hints = _page_hints_from_context(manual_context)
        raw = _call_with_pdf(header, mail_section, manual_pdfs, attachments, system_prompt, hints, extra_section)
        used_pdf = True

    return _parse(raw, used_pdf=used_pdf)


def _page_hints_from_context(manual_context: str) -> list[int]:
    """txt 컨텍스트에서 언급된 페이지 번호 추출 (API 호출 없음)."""
    pages = set()
    for m in re.finditer(r'(?:page|p\.?|페이지)\s*(\d+)', manual_context, re.IGNORECASE):
        n = int(m.group(1))
        if 1 <= n <= 5000:
            pages.add(n)
    return sorted(pages)


def _make_pdf_subset(pdf_path: Path, page_hints: list[int],
                     context: int = 10, max_pages: int = 200) -> Path:
    """
    1000페이지 초과 PDF를 관련 페이지만 추출한 서브셋으로 반환.
    page_hints: 1-based 페이지 번호 목록
    """
    try:
        import fitz
    except ImportError:
        return pdf_path

    doc = fitz.open(str(pdf_path))
    total = len(doc)

    if total <= 900:
        doc.close()
        return pdf_path  # 제한 이하면 원본 그대로

    # 관련 페이지 범위 수집 (0-based)
    pages: set[int] = set()
    for p in page_hints:
        p0 = p - 1
        for i in range(max(0, p0 - context), min(total, p0 + context + 1)):
            pages.add(i)

    if not pages:
        # 힌트 없으면 균등 샘플링
        step = max(1, total // max_pages)
        pages = set(range(0, total, step))

    sorted_pages = sorted(pages)[:max_pages]

    new_doc = fitz.open()
    for p in sorted_pages:
        new_doc.insert_pdf(doc, from_page=p, to_page=p)

    import tempfile
    tmp = tempfile.NamedTemporaryFile(
        suffix=f"_{pdf_path.stem}_subset.pdf", delete=False
    )
    tmp_path = Path(tmp.name)
    tmp.close()
    new_doc.save(str(tmp_path))
    new_doc.close()
    doc.close()
    print(f"[INFO] PDF 서브셋: {pdf_path.name} {total}p → {len(sorted_pages)}p")
    return tmp_path


def _call_with_pdf(
    header: str,
    mail_section: str,
    manual_pdfs: list[Path],
    attachments: list[Path] | None,
    system_prompt: str,
    page_hints: list[int] | None = None,
    extra_section: str = "",
) -> str:
    """PDF 원본을 직접 Gemini에 전달하여 답변 생성. 대용량 PDF는 자동 서브셋."""
    hints = page_hints or []
    text = (
        header
        + "## 참고 매뉴얼 (PDF 원본 — 도면·이미지 포함)\n"
        "첨부된 PDF를 직접 참고하여 지정된 형식으로 분석과 이메일 초안을 작성하세요.\n\n"
        + mail_section
        + extra_section
        + "## 위 형식에 맞춰 분석과 이메일 초안을 작성해주세요:"
    )
    contents: list = [text]
    for pdf in manual_pdfs:
        try:
            subset = _make_pdf_subset(pdf, hints)
            contents.append(_make_part(subset))
        except Exception as e:
            contents.append(f"[PDF 로드 실패: {pdf.name} - {e}]")
    if attachments:
        for att in attachments:
            try:
                contents.append(_make_part(att))
            except Exception:
                pass
    return _call(contents, system_prompt)


if __name__ == "__main__":
    test_body = "안녕하세요. RTC6 보드를 사용 중인데, scan_speed 설정이 적용이 안 됩니다. 확인 부탁드립니다."
    test_manual = "RTC6 Manual: set_scanner_speed() 함수로 갈바노 스캔 속도를 설정합니다. 단위는 bits/ms입니다."

    print("Gemini API 테스트 중...")
    result = generate_reply(test_body, test_manual)
    print("\n[분석]")
    print(result.analysis)
    print("\n[이메일 초안]")
    print(result.draft)
