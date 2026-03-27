# PyInstaller 런타임 훅: win32com gencache 경로 수정
# 실행 파일 내에서는 gen_py 캐시를 사용자 임시 폴더로 리다이렉트합니다.
import os
import sys
import tempfile


def _fix_win32com_gencache():
    if not getattr(sys, "frozen", False):
        return
    try:
        import win32com.client.gencache as gencache

        gen_py_dir = os.path.join(tempfile.gettempdir(), "outlook_gemini_gen_py")
        os.makedirs(gen_py_dir, exist_ok=True)

        # gencache가 임시 폴더를 사용하도록 경로 덮어쓰기
        gencache.GetGeneratePath = lambda: gen_py_dir

        import win32com

        win32com.__gen_path__ = gen_py_dir
    except Exception:
        pass


_fix_win32com_gencache()
