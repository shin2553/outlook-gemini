"""
Outlook + Gemini API 자동 답변 초안 생성
VBA에서 호출되는 메인 진입점
"""

import sys
import logging
import traceback
from pathlib import Path

# 로그 파일 설정
LOG_PATH = Path(__file__).parent / "outlook_gemini.log"
logging.basicConfig(
    filename=str(LOG_PATH),
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    encoding="utf-8",
)
log = logging.getLogger(__name__)


def run():
    log.info("=== 자동 답변 초안 생성 시작 ===")

    # 1. 메일 추출
    log.info("1. 메일 추출 중...")
    from mail_extractor import extract_mail
    try:
        mail = extract_mail()
    except RuntimeError as e:
        log.error(f"메일 추출 실패: {e}")
        _show_error(str(e))
        return

    log.info(f"   제목: {mail.subject}")
    log.info(f"   발신자: {mail.sender}")
    log.info(f"   첨부파일: {[a.name for a in mail.attachments]}")

    # 2. 매뉴얼 검색
    log.info("2. 매뉴얼 검색 중...")
    from manual_searcher import build_manual_context, get_matched_file_names
    query = mail.subject + " " + mail.body
    matched = get_matched_file_names(query)
    log.info(f"   매칭된 매뉴얼: {matched}")
    manual_context = build_manual_context(query)

    # 3. Gemini API 호출
    log.info("3. Gemini API 호출 중...")
    from gemini_client import generate_reply
    try:
        draft = generate_reply(
            mail_body=f"제목: {mail.subject}\n발신자: {mail.sender}\n\n{mail.body}",
            manual_context=manual_context,
            attachments=mail.attachments if mail.attachments else None,
        )
    except Exception as e:
        log.error(f"Gemini API 호출 실패: {e}")
        _show_error(f"Gemini API 오류:\n{e}")
        mail.cleanup()
        return

    log.info(f"   초안 생성 완료 ({len(draft)} chars)")

    # 4. Outlook 답장 창에 주입
    log.info("4. Outlook 답장 창 열기...")
    from outlook_injector import inject_reply
    try:
        inject_reply(draft)
    except Exception as e:
        log.error(f"답장 주입 실패: {e}")
        _show_error(f"답장 창 열기 오류:\n{e}")
        mail.cleanup()
        return

    mail.cleanup()
    log.info("=== 완료 ===")


def _show_error(message: str):
    """사용자에게 오류 메시지를 팝업으로 표시."""
    try:
        import ctypes
        ctypes.windll.user32.MessageBoxW(
            0,
            message,
            "Outlook Gemini - 오류",
            0x10,  # MB_ICONERROR
        )
    except Exception:
        pass  # 팝업 실패 시 로그만 남김


if __name__ == "__main__":
    try:
        run()
    except Exception:
        log.critical(f"예상치 못한 오류:\n{traceback.format_exc()}")
        _show_error(f"예상치 못한 오류가 발생했습니다.\n로그 파일을 확인하세요:\n{LOG_PATH}")
        sys.exit(1)
