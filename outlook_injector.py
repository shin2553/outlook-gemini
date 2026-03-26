"""
Outlook 답장 창에 생성된 초안을 자동 입력합니다.
"""

import win32com.client


def _get_active_mail_item():
    """Outlook에서 현재 활성화된 MailItem 반환."""
    outlook = win32com.client.Dispatch("Outlook.Application")

    inspector = outlook.ActiveInspector()
    if inspector and inspector.CurrentItem:
        item = inspector.CurrentItem
        if item.Class == 43:
            return item

    explorer = outlook.ActiveExplorer()
    if explorer:
        selection = explorer.Selection
        if selection.Count > 0:
            item = selection.Item(1)
            if item.Class == 43:
                return item

    return None


def _text_to_html(text: str) -> str:
    """일반 텍스트를 기본 HTML로 변환 (줄바꿈 처리)."""
    import html
    escaped = html.escape(text)
    lines = escaped.split("\n")
    html_lines = []
    for line in lines:
        if line.strip() == "":
            html_lines.append("<br>")
        else:
            html_lines.append(f"<p style='margin:0;font-family:맑은 고딕,Calibri,sans-serif;font-size:10pt'>{line}</p>")
    return "\n".join(html_lines)


def inject_reply(draft_text: str, reply_all: bool = False,
                 image_paths: list | None = None) -> None:
    """
    현재 선택된 메일에 대한 답장 창을 열고 초안을 입력합니다.

    Args:
        draft_text:   Gemini가 생성한 답변 초안
        reply_all:    True면 전체 답장, False면 답장
        image_paths:  인라인으로 삽입할 이미지 경로 목록 (PNG)

    Raises:
        RuntimeError: 활성 메일을 찾을 수 없는 경우
    """
    item = _get_active_mail_item()
    if item is None:
        raise RuntimeError("Outlook에서 선택된 메일을 찾을 수 없습니다.")

    reply = item.ReplyAll() if reply_all else item.Reply()
    existing_html = reply.HTMLBody or ""
    draft_html = _text_to_html(draft_text)

    # 인라인 이미지 HTML (cid 참조)
    img_html = ""
    if image_paths:
        img_html = "<div style='margin:12px 0'>\n"
        for i, img_path in enumerate(image_paths):
            cid = f"manualpage_{i}.png"
            img_html += (
                f"<p style='margin:4px 0;font-size:8pt;color:#888'>"
                f"[매뉴얼 참조 이미지 {i+1}]</p>\n"
                f"<img src='cid:{cid}' style='max-width:600px;"
                f"border:1px solid #ddd;margin-bottom:8px'><br>\n"
            )
        img_html += "</div>\n"

    separator = "<hr style='border:none;border-top:1px solid #ccc;margin:12px 0'>"

    if "<body" in existing_html.lower():
        insert_pos = existing_html.lower().find("<body")
        tag_end = existing_html.find(">", insert_pos) + 1
        new_html = (
            existing_html[:tag_end]
            + "\n" + draft_html + "\n"
            + img_html
            + separator + "\n"
            + existing_html[tag_end:]
        )
    else:
        new_html = draft_html + img_html + separator + existing_html

    reply.HTMLBody = new_html

    # 이미지를 인라인 첨부(CID)로 추가
    if image_paths:
        _CONTENT_ID_PROP = "http://schemas.microsoft.com/mapi/proptag/0x3712001E"
        for i, img_path in enumerate(image_paths):
            cid = f"manualpage_{i}.png"
            att = reply.Attachments.Add(str(img_path))
            try:
                att.PropertyAccessor.SetProperty(_CONTENT_ID_PROP, cid)
            except Exception:
                pass  # CID 설정 실패 시 일반 첨부로 폴백

    reply.Display()


if __name__ == "__main__":
    test_draft = """안녕하세요, SCANLAB 기술지원팀입니다.

RTC6 보드의 scan_speed 설정 관련하여 문의해 주셨습니다.

set_scanner_speed() 함수의 단위는 bits/ms입니다.

추가 문의 사항이 있으시면 언제든지 연락 주시기 바랍니다.

감사합니다."""

    print("Outlook 답장 주입 테스트...")
    try:
        inject_reply(test_draft)
        print("답장 창이 열렸습니다. 확인해주세요.")
    except RuntimeError as e:
        print(f"[오류] {e}")
