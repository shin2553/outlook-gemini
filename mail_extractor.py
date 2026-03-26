"""
Outlook COM으로 현재 선택된/열린 메일에서 본문과 첨부파일을 추출합니다.
"""

import re
import tempfile
from dataclasses import dataclass, field
from pathlib import Path

import win32com.client


SUPPORTED_EXTENSIONS = {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tiff", ".pdf"}


@dataclass
class MailData:
    subject: str
    sender: str
    body: str
    attachments: list[Path] = field(default_factory=list)
    temp_dir: tempfile.TemporaryDirectory = field(default=None, repr=False)

    def cleanup(self):
        """임시 첨부파일 디렉토리 삭제"""
        if self.temp_dir:
            self.temp_dir.cleanup()


def _html_to_text(html: str) -> str:
    """HTML 태그 제거하여 순수 텍스트 추출"""
    text = re.sub(r"<style[^>]*>.*?</style>", "", html, flags=re.DOTALL | re.IGNORECASE)
    text = re.sub(r"<script[^>]*>.*?</script>", "", text, flags=re.DOTALL | re.IGNORECASE)
    text = re.sub(r"<br\s*/?>", "\n", text, flags=re.IGNORECASE)
    text = re.sub(r"<p[^>]*>", "\n", text, flags=re.IGNORECASE)
    text = re.sub(r"</p>", "\n", text, flags=re.IGNORECASE)
    text = re.sub(r"<[^>]+>", "", text)
    text = re.sub(r"&nbsp;", " ", text)
    text = re.sub(r"&lt;", "<", text)
    text = re.sub(r"&gt;", ">", text)
    text = re.sub(r"&amp;", "&", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def _get_active_mail_item():
    """Outlook에서 현재 활성화된 MailItem 반환."""
    outlook = win32com.client.Dispatch("Outlook.Application")

    # 1) 열린 Inspector(편집창) 우선
    inspector = outlook.ActiveInspector()
    if inspector and inspector.CurrentItem:
        item = inspector.CurrentItem
        if item.Class == 43:  # olMail
            return item

    # 2) Explorer(목록창)에서 선택된 메일
    explorer = outlook.ActiveExplorer()
    if explorer:
        selection = explorer.Selection
        if selection.Count > 0:
            item = selection.Item(1)
            if item.Class == 43:
                return item

    return None


def extract_mail() -> MailData:
    """
    현재 Outlook에서 선택된 메일의 본문과 첨부파일을 추출합니다.

    Returns:
        MailData 객체

    Raises:
        RuntimeError: 활성 메일을 찾을 수 없는 경우
    """
    item = _get_active_mail_item()
    if item is None:
        raise RuntimeError("Outlook에서 선택된 메일을 찾을 수 없습니다.\n메일을 열거나 선택한 후 다시 시도하세요.")

    subject = item.Subject or "(제목 없음)"
    sender = item.SenderName or item.SenderEmailAddress or "알 수 없음"

    # 본문: HTML → 텍스트 변환 (HTML 없으면 일반 텍스트)
    html_body = item.HTMLBody
    if html_body:
        body = _html_to_text(html_body)
    else:
        body = item.Body or ""

    # 첨부파일 저장
    temp_dir = tempfile.TemporaryDirectory(prefix="outlook_gemini_")
    saved_attachments: list[Path] = []

    for i in range(1, item.Attachments.Count + 1):
        att = item.Attachments.Item(i)
        att_name = att.FileName
        ext = Path(att_name).suffix.lower()

        if ext not in SUPPORTED_EXTENSIONS:
            continue

        save_path = Path(temp_dir.name) / att_name
        att.SaveAsFile(str(save_path))
        saved_attachments.append(save_path)

    return MailData(
        subject=subject,
        sender=sender,
        body=body,
        attachments=saved_attachments,
        temp_dir=temp_dir,
    )


if __name__ == "__main__":
    print("Outlook 메일 추출 테스트...")
    try:
        mail = extract_mail()
        print(f"제목: {mail.subject}")
        print(f"발신자: {mail.sender}")
        print(f"본문 (앞 300자):\n{mail.body[:300]}")
        print(f"첨부파일 ({len(mail.attachments)}개):")
        for att in mail.attachments:
            print(f"  - {att.name} ({att.stat().st_size // 1024} KB)")
        mail.cleanup()
    except RuntimeError as e:
        print(f"[오류] {e}")
