"""
Outlook Gemini - 자동 답변 초안 생성 UI
"""

import configparser
import json
import threading
import tkinter as tk
from pathlib import Path
from tkinter import ttk, messagebox, filedialog

# ── 경로 설정 ──────────────────────────────────────────────────
import os
from config import get_app_dir

_APP_DIR = get_app_dir()
os.chdir(str(_APP_DIR))

PROFILES_PATH = _APP_DIR / "profiles.json"
CONFIG_PATH   = _APP_DIR / "config.ini"


# ── config.ini 읽기/쓰기 ───────────────────────────────────────

def _read_config() -> configparser.ConfigParser:
    cfg = configparser.ConfigParser()
    cfg.read(str(CONFIG_PATH), encoding="utf-8")
    return cfg


def _write_config(cfg: configparser.ConfigParser):
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        cfg.write(f)


def load_api_key() -> str:
    return _read_config().get("gemini", "api_key", fallback="")


def save_api_key(key: str):
    cfg = _read_config()
    if not cfg.has_section("gemini"):
        cfg.add_section("gemini")
    cfg.set("gemini", "api_key", key)
    _write_config(cfg)


def load_paths() -> dict:
    cfg = _read_config()
    return {
        "manual_dir": cfg.get("paths", "manual_dir", fallback=""),
        "pdf_dir":    cfg.get("paths", "pdf_dir",    fallback=""),
    }


def save_paths(data: dict):
    cfg = _read_config()
    if not cfg.has_section("paths"):
        cfg.add_section("paths")
    for k, v in data.items():
        cfg.set("paths", k, v)
    _write_config(cfg)


def load_company() -> dict:
    cfg = _read_config()
    return {"role": cfg.get("company", "role", fallback="")}


def save_company(data: dict):
    cfg = _read_config()
    if not cfg.has_section("company"):
        cfg.add_section("company")
    for k, v in data.items():
        cfg.set("company", k, v)
    _write_config(cfg)


# ── 설정 다이얼로그 ────────────────────────────────────────────

class SettingsDialog(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("설정")
        self.geometry("640x400")
        self.resizable(False, False)
        self.grab_set()
        self._pending_pdfs: list = []
        self._build_ui()

    def _build_ui(self):
        notebook = ttk.Notebook(self)
        notebook.pack(fill=tk.BOTH, expand=True, padx=12, pady=12)

        # ── 탭1: API 키 ──────────────────────────────────────
        tab_api = ttk.Frame(notebook, padding=12)
        notebook.add(tab_api, text="  Gemini API 키  ")

        tk.Label(tab_api, text="API 키:", font=("맑은 고딕", 9, "bold")).grid(row=0, column=0, sticky="w", pady=6)
        self._ent_key = ttk.Entry(tab_api, width=44, font=("맑은 고딕", 9), show="*")
        self._ent_key.grid(row=0, column=1, sticky="ew", padx=(8, 0), pady=6)
        self._ent_key.insert(0, load_api_key())

        self._show_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            tab_api, text="키 표시", variable=self._show_var,
            command=lambda: self._ent_key.config(show="" if self._show_var.get() else "*"),
        ).grid(row=1, column=1, sticky="w", padx=(8, 0))

        api_btn = tk.Frame(tab_api)
        api_btn.grid(row=2, column=0, columnspan=2, sticky="e", pady=(16, 0))
        ttk.Button(api_btn, text="연결 테스트", command=self._test).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(api_btn, text="저장", command=self._save_api, width=8).pack(side=tk.LEFT)
        tab_api.columnconfigure(1, weight=1)

        # ── 탭2: 회사 정보 ───────────────────────────────────
        tab_co = ttk.Frame(notebook, padding=12)
        notebook.add(tab_co, text="  회사 정보  ")

        co = load_company()
        fields = [
            ("관계 설명:", "role", co.get("role", "")),
        ]
        self._co_entries: dict[str, ttk.Entry] = {}
        for i, (label, key, val) in enumerate(fields):
            tk.Label(tab_co, text=label, font=("맑은 고딕", 9, "bold")).grid(row=i, column=0, sticky="w", pady=5, padx=(0, 8))
            ent = ttk.Entry(tab_co, width=38, font=("맑은 고딕", 9))
            ent.insert(0, val)
            ent.grid(row=i, column=1, sticky="ew", pady=5)
            self._co_entries[key] = ent

        tk.Label(
            tab_co,
            text='예) "여러 제조사 제품의 국내 공식 대리점"',
            font=("맑은 고딕", 8), fg="#888",
        ).grid(row=len(fields), column=0, columnspan=2, sticky="w", pady=(4, 0))

        co_btn = tk.Frame(tab_co)
        co_btn.grid(row=len(fields)+1, column=0, columnspan=2, sticky="e", pady=(16, 0))
        ttk.Button(co_btn, text="저장", command=self._save_company, width=8).pack(side=tk.LEFT)
        tab_co.columnconfigure(1, weight=1)

        # ── 탭3: 매뉴얼 관리 ─────────────────────────────────
        tab_manual = ttk.Frame(notebook, padding=10)
        notebook.add(tab_manual, text="  매뉴얼 관리  ")

        paths = load_paths()

        # 경로 설정
        path_frame = ttk.LabelFrame(tab_manual, text=" 경로 설정 ", padding=8)
        path_frame.pack(fill=tk.X, pady=(0, 8))

        tk.Label(path_frame, text="txt 매뉴얼 폴더:", font=("맑은 고딕", 9, "bold")).grid(row=0, column=0, sticky="w", pady=4, padx=(0, 6))
        self._ent_dir = ttk.Entry(path_frame, width=38, font=("맑은 고딕", 8))
        self._ent_dir.insert(0, paths.get("manual_dir", ""))
        self._ent_dir.grid(row=0, column=1, sticky="ew", pady=4)
        ttk.Button(path_frame, text="찾기", width=5,
                   command=lambda: self._browse_dir(self._ent_dir)).grid(row=0, column=2, padx=(4, 0))
        tk.Label(path_frame, text="※ manual_index.py가 같은 폴더 또는 상위 폴더에 있어야 합니다.",
                 font=("맑은 고딕", 7), fg="#888").grid(row=1, column=0, columnspan=3, sticky="w")

        tk.Label(path_frame, text="PDF 원본 폴더:", font=("맑은 고딕", 9, "bold")).grid(row=2, column=0, sticky="w", pady=4, padx=(0, 6))
        self._ent_pdf_dir = ttk.Entry(path_frame, width=38, font=("맑은 고딕", 8))
        self._ent_pdf_dir.insert(0, paths.get("pdf_dir", ""))
        self._ent_pdf_dir.grid(row=2, column=1, sticky="ew", pady=4)
        ttk.Button(path_frame, text="찾기", width=5,
                   command=lambda: self._browse_dir(self._ent_pdf_dir)).grid(row=2, column=2, padx=(4, 0))
        tk.Label(path_frame, text="※ 비워두면 txt만 사용. PDF 있으면 도면 필요 시 자동으로 원본 참조.",
                 font=("맑은 고딕", 7), fg="#888").grid(row=3, column=0, columnspan=3, sticky="w")

        ttk.Button(path_frame, text="경로 저장", command=self._save_paths).grid(
            row=4, column=0, columnspan=3, sticky="e", pady=(6, 0))
        path_frame.columnconfigure(1, weight=1)

        # 매뉴얼 파일 목록 + PDF 변환
        file_frame = ttk.LabelFrame(tab_manual, text=" 매뉴얼 파일 목록 ", padding=8)
        file_frame.pack(fill=tk.BOTH, expand=True)

        list_scroll = ttk.Scrollbar(file_frame, orient=tk.VERTICAL)
        self._lst_files = tk.Listbox(
            file_frame, font=("맑은 고딕", 8), bg="#fafafa", relief=tk.FLAT,
            yscrollcommand=list_scroll.set, height=6,
        )
        list_scroll.config(command=self._lst_files.yview)
        self._lst_files.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        list_scroll.pack(side=tk.LEFT, fill=tk.Y)

        file_btn = tk.Frame(file_frame)
        file_btn.pack(side=tk.LEFT, padx=(8, 0), anchor="n")
        ttk.Button(file_btn, text="목록 새로고침", width=16, command=self._refresh_file_list).pack(pady=(0, 4))
        ttk.Button(file_btn, text="새 PDF 자동 변환", width=16, command=self._sync_pdfs).pack(pady=(0, 4))
        ttk.Button(file_btn, text="PDF 직접 선택 변환", width=16, command=self._add_pdf).pack(pady=(0, 4))
        ttk.Button(file_btn, text="선택 파일 삭제", width=16, command=self._delete_manual).pack()

        self._lbl_convert = tk.Label(tab_manual, text="", font=("맑은 고딕", 8), fg="#1a73e8")
        self._lbl_convert.pack(anchor="w", pady=(4, 0))

        self._refresh_file_list()

        # ── 공통 닫기 버튼 ────────────────────────────────────
        ttk.Button(self, text="닫기", command=self.destroy, width=8).pack(pady=(0, 10))

    # ── API 키 ────────────────────────────────────────────────

    def _save_api(self):
        key = self._ent_key.get().strip()
        if not key:
            messagebox.showwarning("입력 오류", "API 키를 입력하세요.", parent=self)
            return
        save_api_key(key)
        try:
            import importlib
            import config as cfg_module
            importlib.reload(cfg_module)
            import gemini_client
            from google import genai
            gemini_client._client = genai.Client(api_key=cfg_module.GEMINI_API_KEY)
        except Exception:
            pass
        messagebox.showinfo("저장 완료", "API 키가 저장되었습니다.", parent=self)

    def _test(self):
        key = self._ent_key.get().strip()
        if not key:
            messagebox.showwarning("입력 오류", "API 키를 입력하세요.", parent=self)
            return
        self._ent_key.config(state=tk.DISABLED)
        threading.Thread(target=self._do_test, args=(key,), daemon=True).start()

    def _do_test(self, key: str):
        try:
            from google import genai
            from google.genai import types
            client = genai.Client(api_key=key)
            resp = client.models.generate_content(
                model="gemini-2.5-flash",
                contents="ping",
                config=types.GenerateContentConfig(max_output_tokens=8),
            )
            resp.text
            self.after(0, lambda: self._test_result(True, "연결 성공 ✓"))
        except Exception as e:
            self.after(0, lambda: self._test_result(False, str(e)))

    def _test_result(self, ok: bool, msg: str):
        self._ent_key.config(state=tk.NORMAL)
        if ok:
            messagebox.showinfo("테스트 결과", msg, parent=self)
        else:
            messagebox.showerror("테스트 실패", msg, parent=self)

    # ── 매뉴얼 관리 ───────────────────────────────────────────

    def _refresh_file_list(self):
        manual_dir = self._ent_dir.get().strip()
        pdf_dir    = self._ent_pdf_dir.get().strip()
        self._lst_files.delete(0, tk.END)
        if not manual_dir:
            return
        folder = Path(manual_dir)
        if not folder.exists():
            return

        # 변환된 txt 목록
        existing_stems = set()
        for f in sorted(folder.glob("*.txt")):
            size_kb = f.stat().st_size // 1024
            self._lst_files.insert(tk.END, f"  {f.name}  ({size_kb} KB)")
            existing_stems.add(f.stem)

        # PDF 원본 폴더에서 미변환 파일 자동 탐지
        self._pending_pdfs = []
        if pdf_dir:
            pdf_folder = Path(pdf_dir)
            if pdf_folder.exists():
                seen = set()
                all_pdfs = sorted(pdf_folder.glob("*.pdf")) + sorted(pdf_folder.glob("*.PDF"))
                for pdf in all_pdfs:
                    if pdf.stem in seen:
                        continue
                    seen.add(pdf.stem)
                    if pdf.stem not in existing_stems:
                        self._pending_pdfs.append(pdf)

        if self._pending_pdfs:
            self._lbl_convert.config(
                text=f"⚠  미변환 PDF {len(self._pending_pdfs)}개 발견 — '새 PDF 자동 변환' 버튼으로 일괄 변환",
                fg="#e67e00",
            )
        else:
            self._lbl_convert.config(text="✓  모든 PDF가 변환되어 있습니다.", fg="#1a73e8")

    def _sync_pdfs(self):
        """PDF 원본 폴더에서 미변환 파일을 자동으로 일괄 변환."""
        manual_dir = self._ent_dir.get().strip()
        if not manual_dir:
            messagebox.showwarning("경로 없음", "먼저 txt 매뉴얼 폴더를 지정하세요.", parent=self)
            return
        if not hasattr(self, "_pending_pdfs") or not self._pending_pdfs:
            messagebox.showinfo("알림", "변환할 새 PDF가 없습니다.\n'목록 새로고침'을 먼저 눌러주세요.", parent=self)
            return

        count = len(self._pending_pdfs)
        if not messagebox.askyesno(
            "자동 변환 확인",
            f"미변환 PDF {count}개를 txt로 변환합니다.\n\n"
            + "\n".join(f"  • {p.name}" for p in self._pending_pdfs[:10])
            + (f"\n  ... 외 {count-10}개" if count > 10 else ""),
            parent=self,
        ):
            return

        self._lbl_convert.config(text="변환 중...", fg="#555")
        threading.Thread(
            target=self._do_convert,
            args=([str(p) for p in self._pending_pdfs], manual_dir),
            daemon=True,
        ).start()

    def _add_pdf(self):
        manual_dir = self._ent_dir.get().strip()
        if not manual_dir:
            messagebox.showwarning("경로 없음", "먼저 매뉴얼 폴더를 지정하세요.", parent=self)
            return

        pdf_paths = filedialog.askopenfilenames(
            parent=self,
            title="변환할 PDF 파일 선택 (복수 선택 가능)",
            filetypes=[("PDF 파일", "*.pdf"), ("모든 파일", "*.*")],
        )
        if not pdf_paths:
            return

        self._lbl_convert.config(text="변환 중...")
        threading.Thread(
            target=self._do_convert, args=(list(pdf_paths), manual_dir), daemon=True
        ).start()

    def _do_convert(self, pdf_paths: list, manual_dir: str):
        from pathlib import Path
        from pdf_converter import convert_pdf_to_txt

        results = []
        for pdf_path in pdf_paths:
            pdf = Path(pdf_path)
            try:
                out, pages = convert_pdf_to_txt(pdf, Path(manual_dir))
                results.append(f"✓ {pdf.name} → {out.name} ({pages}페이지)")
            except Exception as e:
                results.append(f"✗ {pdf.name}: {e}")

        def _done():
            self._refresh_file_list()
            self._lbl_convert.config(text=f"완료: {len(pdf_paths)}개 처리")
            messagebox.showinfo("변환 결과", "\n".join(results), parent=self)

        self.after(0, _done)

    def _delete_manual(self):
        sel = self._lst_files.curselection()
        if not sel:
            messagebox.showwarning("선택 없음", "삭제할 파일을 선택하세요.", parent=self)
            return
        manual_dir = self._ent_dir.get().strip()
        if not manual_dir:
            return
        # 목록 항목에서 파일명 추출 ("  filename.txt  (123 KB)" 형식)
        item = self._lst_files.get(sel[0]).strip().split("  ")[0]
        from pathlib import Path
        target = Path(manual_dir) / item
        if not target.exists():
            messagebox.showerror("오류", f"파일을 찾을 수 없습니다:\n{item}", parent=self)
            return
        if not messagebox.askyesno("삭제 확인", f"'{item}' 파일을 삭제하시겠습니까?", parent=self):
            return
        target.unlink()
        self._refresh_file_list()
        self._lbl_convert.config(text=f"'{item}' 삭제됨")

    def _browse_dir(self, entry: ttk.Entry):
        path = filedialog.askdirectory(parent=self, title="매뉴얼 폴더 선택")
        if path:
            entry.delete(0, tk.END)
            entry.insert(0, path)

    def _save_paths(self):
        data = {
            "manual_dir": self._ent_dir.get().strip(),
            "pdf_dir":    self._ent_pdf_dir.get().strip(),
        }
        save_paths(data)
        try:
            import importlib
            import config as cfg_module
            importlib.reload(cfg_module)
            import manual_searcher
            importlib.reload(manual_searcher)
        except Exception:
            pass
        messagebox.showinfo("저장 완료", "매뉴얼 경로가 저장되었습니다.\n다음 메일 불러오기부터 반영됩니다.", parent=self)

    # ── 회사 정보 ─────────────────────────────────────────────

    def _save_company(self):
        data = {k: ent.get().strip() for k, ent in self._co_entries.items()}
        save_company(data)
        try:
            import importlib
            import config as cfg_module
            importlib.reload(cfg_module)
        except Exception:
            pass
        messagebox.showinfo("저장 완료", "회사 정보가 저장되었습니다.\n다음 초안 생성부터 반영됩니다.", parent=self)


# ── 프로필 관리 ────────────────────────────────────────────────

def load_profiles() -> list[dict]:
    if PROFILES_PATH.exists():
        try:
            return json.loads(PROFILES_PATH.read_text(encoding="utf-8"))
        except Exception:
            pass
    return []


def save_profiles(profiles: list[dict]):
    PROFILES_PATH.write_text(
        json.dumps(profiles, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


# ── 프로필 관리 다이얼로그 ─────────────────────────────────────

class ProfileDialog(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("답변 주체 프로필 관리")
        self.geometry("420x320")
        self.resizable(False, False)
        self.grab_set()  # 모달

        self._profiles = load_profiles()
        self._build_ui()
        self._refresh_list()

    def _build_ui(self):
        # 목록
        list_frame = ttk.LabelFrame(self, text=" 프로필 목록 ", padding=8)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=12, pady=(12, 6))

        self._listbox = tk.Listbox(list_frame, font=("맑은 고딕", 10), height=7, relief=tk.FLAT, bg="#fafafa")
        self._listbox.pack(fill=tk.BOTH, expand=True)
        self._listbox.bind("<<ListboxSelect>>", self._on_select)

        # 입력 폼
        form = ttk.LabelFrame(self, text=" 추가 / 수정 ", padding=8)
        form.pack(fill=tk.X, padx=12, pady=6)

        tk.Label(form, text="이름:", font=("맑은 고딕", 9)).grid(row=0, column=0, sticky="w", padx=4, pady=3)
        self._ent_name = ttk.Entry(form, width=14, font=("맑은 고딕", 9))
        self._ent_name.grid(row=0, column=1, sticky="w", padx=4)

        tk.Label(form, text="직책:", font=("맑은 고딕", 9)).grid(row=0, column=2, sticky="w", padx=(8, 4))
        self._ent_title = ttk.Entry(form, width=14, font=("맑은 고딕", 9))
        self._ent_title.grid(row=0, column=3, sticky="w", padx=4)

        tk.Label(form, text="회사:", font=("맑은 고딕", 9)).grid(row=1, column=0, sticky="w", padx=4, pady=3)
        self._ent_company = ttk.Entry(form, width=14, font=("맑은 고딕", 9))
        self._ent_company.grid(row=1, column=1, sticky="w", padx=4)

        tk.Label(form, text="부서:", font=("맑은 고딕", 9)).grid(row=1, column=2, sticky="w", padx=(8, 4))
        self._ent_dept = ttk.Entry(form, width=14, font=("맑은 고딕", 9))
        self._ent_dept.grid(row=1, column=3, sticky="w", padx=4)

        self._chk_default_var = tk.BooleanVar()
        ttk.Checkbutton(form, text="기본값으로 설정", variable=self._chk_default_var).grid(
            row=2, column=0, columnspan=4, sticky="w", padx=4, pady=(4, 0)
        )

        # 버튼
        btn_frame = tk.Frame(self, bg="#f0f0f0")
        btn_frame.pack(fill=tk.X, padx=12, pady=(4, 12))

        ttk.Button(btn_frame, text="추가", command=self._add, width=8).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="수정", command=self._update, width=8).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="삭제", command=self._delete, width=8).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="닫기", command=self._close, width=8).pack(side=tk.RIGHT, padx=2)

    def _refresh_list(self):
        self._listbox.delete(0, tk.END)
        for p in self._profiles:
            star = " ★" if p.get("default") else ""
            company = p.get("company", "")
            dept    = p.get("department", "")
            affil   = f"  [{company} {dept}]".strip() if company or dept else ""
            self._listbox.insert(tk.END, f"  {p['name']}  /  {p['title']}{affil}{star}")

    def _on_select(self, _event=None):
        sel = self._listbox.curselection()
        if not sel:
            return
        p = self._profiles[sel[0]]
        self._ent_name.delete(0, tk.END)
        self._ent_name.insert(0, p.get("name", ""))
        self._ent_title.delete(0, tk.END)
        self._ent_title.insert(0, p.get("title", ""))
        self._ent_company.delete(0, tk.END)
        self._ent_company.insert(0, p.get("company", ""))
        self._ent_dept.delete(0, tk.END)
        self._ent_dept.insert(0, p.get("department", ""))
        self._chk_default_var.set(p.get("default", False))

    def _add(self):
        name = self._ent_name.get().strip()
        if not name:
            messagebox.showwarning("입력 오류", "이름을 입력하세요.", parent=self)
            return
        if self._chk_default_var.get():
            for p in self._profiles:
                p["default"] = False
        self._profiles.append({
            "name": name, "title": self._ent_title.get().strip(),
            "company": self._ent_company.get().strip(),
            "department": self._ent_dept.get().strip(),
            "default": self._chk_default_var.get(),
        })
        save_profiles(self._profiles)
        self._refresh_list()

    def _update(self):
        sel = self._listbox.curselection()
        if not sel:
            messagebox.showwarning("선택 오류", "수정할 프로필을 선택하세요.", parent=self)
            return
        name = self._ent_name.get().strip()
        if not name:
            messagebox.showwarning("입력 오류", "이름을 입력하세요.", parent=self)
            return
        if self._chk_default_var.get():
            for p in self._profiles:
                p["default"] = False
        self._profiles[sel[0]] = {
            "name": name, "title": self._ent_title.get().strip(),
            "company": self._ent_company.get().strip(),
            "department": self._ent_dept.get().strip(),
            "default": self._chk_default_var.get(),
        }
        save_profiles(self._profiles)
        self._refresh_list()

    def _delete(self):
        sel = self._listbox.curselection()
        if not sel:
            messagebox.showwarning("선택 오류", "삭제할 프로필을 선택하세요.", parent=self)
            return
        if not messagebox.askyesno("삭제 확인", "선택한 프로필을 삭제하시겠습니까?", parent=self):
            return
        del self._profiles[sel[0]]
        save_profiles(self._profiles)
        self._refresh_list()

    def _close(self):
        self.destroy()
        # 부모 앱에 변경 알림
        if hasattr(self.master, "_reload_profiles"):
            self.master._reload_profiles()


# ── 메인 앱 ────────────────────────────────────────────────────

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Outlook Gemini - 자동 답변 초안 생성")
        self.geometry("1200x800")
        self.minsize(800, 600)
        self.state("zoomed")  # 시작 시 최대화
        self.configure(bg="#f0f0f0")

        self._mail = None
        self._manual_context = ""
        self._page_ref_vars: list[tuple[str, int, tk.BooleanVar]] = []

        self._build_ui()
        self._reload_profiles()

    # ── UI 구성 ────────────────────────────────────────────────

    def _build_ui(self):
        # ── 헤더 ─────────────────────────────────────────────
        header = tk.Frame(self, bg="#1a73e8", height=48)
        header.pack(fill=tk.X)
        header.pack_propagate(False)
        tk.Label(
            header,
            text="  🤖  Gemini 자동 답변 초안 생성",
            bg="#1a73e8", fg="white",
            font=("맑은 고딕", 13, "bold"),
            anchor="w",
        ).pack(side=tk.LEFT, padx=12, pady=10)
        tk.Button(
            header,
            text="⚙  설정",
            command=lambda: SettingsDialog(self),
            bg="#1557b0", fg="white",
            font=("맑은 고딕", 9),
            relief=tk.FLAT, cursor="hand2",
            padx=10, pady=4,
        ).pack(side=tk.RIGHT, padx=12, pady=10)

        # 메인 영역
        main = tk.Frame(self, bg="#f0f0f0")
        main.pack(fill=tk.BOTH, expand=True, padx=14, pady=(8, 6))

        # ── 답변 주체 (한 줄) ────────────────────────────────
        responder_frame = ttk.LabelFrame(main, text=" 답변 주체 ", padding=6)
        responder_frame.pack(fill=tk.X, pady=(0, 6))

        tk.Label(responder_frame, text="프로필:", font=("맑은 고딕", 9, "bold")).pack(side=tk.LEFT, padx=(0, 4))
        self._cmb_profile = ttk.Combobox(responder_frame, state="readonly", width=20, font=("맑은 고딕", 9))
        self._cmb_profile.pack(side=tk.LEFT, padx=(0, 10))
        self._cmb_profile.bind("<<ComboboxSelected>>", self._on_profile_selected)

        tk.Label(responder_frame, text="이름:", font=("맑은 고딕", 9, "bold")).pack(side=tk.LEFT, padx=(0, 3))
        self._ent_resp_name = ttk.Entry(responder_frame, width=10, font=("맑은 고딕", 9))
        self._ent_resp_name.pack(side=tk.LEFT, padx=(0, 8))

        tk.Label(responder_frame, text="직책:", font=("맑은 고딕", 9, "bold")).pack(side=tk.LEFT, padx=(0, 3))
        self._ent_resp_title = ttk.Entry(responder_frame, width=8, font=("맑은 고딕", 9))
        self._ent_resp_title.pack(side=tk.LEFT, padx=(0, 8))

        tk.Label(responder_frame, text="회사:", font=("맑은 고딕", 9, "bold")).pack(side=tk.LEFT, padx=(0, 3))
        self._ent_resp_company = ttk.Entry(responder_frame, width=10, font=("맑은 고딕", 9))
        self._ent_resp_company.pack(side=tk.LEFT, padx=(0, 8))

        tk.Label(responder_frame, text="부서:", font=("맑은 고딕", 9, "bold")).pack(side=tk.LEFT, padx=(0, 3))
        self._ent_resp_dept = ttk.Entry(responder_frame, width=8, font=("맑은 고딕", 9))
        self._ent_resp_dept.pack(side=tk.LEFT, padx=(0, 10))

        ttk.Button(responder_frame, text="프로필 관리", command=self._open_profile_dialog).pack(side=tk.LEFT)

        self._btn_reload = tk.Button(
            responder_frame,
            text="📥  메일 불러오기",
            command=self._load_mail,
            bg="#1a73e8", fg="white",
            font=("맑은 고딕", 10, "bold"),
            relief=tk.FLAT, cursor="hand2",
            padx=16, pady=5,
        )
        self._btn_reload.pack(side=tk.RIGHT)

        # ── 3단 중간 영역: 메일 정보 | 매뉴얼 | 모드+지시사항 ──
        mid = tk.Frame(main, bg="#f0f0f0")
        mid.pack(fill=tk.X, pady=(0, 0))
        mid.columnconfigure(0, weight=2, uniform="mid")
        mid.columnconfigure(1, weight=1, uniform="mid")
        mid.columnconfigure(2, weight=1, uniform="mid")

        # 왼쪽: 메일 정보
        mail_frame = ttk.LabelFrame(mid, text=" 메일 정보 ", padding=8)
        mail_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 5))
        mail_frame.columnconfigure(1, weight=1)
        mail_frame.rowconfigure(3, weight=1)

        def _update_wraplength(event):
            w = max(80, event.width - 90)
            self._lbl_subject.config(wraplength=w)
            self._lbl_attach.config(wraplength=w)
        mail_frame.bind("<Configure>", _update_wraplength)

        grid_opts = {"sticky": "w", "padx": 4, "pady": 2}
        tk.Label(mail_frame, text="제목:", font=("맑은 고딕", 9, "bold")).grid(row=0, column=0, **grid_opts)
        self._lbl_subject = tk.Label(mail_frame, text="-", wraplength=400, justify="left", font=("맑은 고딕", 9))
        self._lbl_subject.grid(row=0, column=1, sticky="ew", padx=4, pady=2)

        tk.Label(mail_frame, text="발신자:", font=("맑은 고딕", 9, "bold")).grid(row=1, column=0, **grid_opts)
        self._lbl_sender = tk.Label(mail_frame, text="-", font=("맑은 고딕", 9))
        self._lbl_sender.grid(row=1, column=1, sticky="w", padx=4, pady=2)

        tk.Label(mail_frame, text="첨부파일:", font=("맑은 고딕", 9, "bold")).grid(row=2, column=0, **grid_opts)
        self._lbl_attach = tk.Label(mail_frame, text="-", wraplength=400, justify="left",
                                    font=("맑은 고딕", 9), fg="#555")
        self._lbl_attach.grid(row=2, column=1, sticky="ew", padx=4, pady=2)

        tk.Label(mail_frame, text="본문\n(미리보기):", font=("맑은 고딕", 9, "bold"), justify="left").grid(
            row=3, column=0, sticky="nw", padx=4, pady=2)
        self._txt_body = tk.Text(
            mail_frame, height=4, state=tk.DISABLED,
            font=("맑은 고딕", 8), bg="#fafafa", relief=tk.FLAT,
            wrap=tk.WORD, bd=1,
        )
        self._txt_body.grid(row=3, column=1, sticky="nsew", padx=4, pady=2)

        # 중간: 매칭된 매뉴얼
        manual_col = tk.Frame(mid, bg="#f0f0f0")
        manual_col.grid(row=0, column=1, sticky="nsew", padx=5)
        manual_col.rowconfigure(0, weight=1)

        manual_frame = ttk.LabelFrame(manual_col, text=" 매칭된 매뉴얼 ", padding=8)
        manual_frame.pack(fill=tk.BOTH, expand=True)

        self._lst_manuals = tk.Listbox(
            manual_frame, height=6,
            font=("맑은 고딕", 8), bg="#fafafa",
            selectmode=tk.BROWSE, relief=tk.FLAT,
        )
        self._lst_manuals.pack(fill=tk.BOTH, expand=True)

        # 오른쪽: 분석 모드 + 추가 지시사항
        ctrl_col = tk.Frame(mid, bg="#f0f0f0")
        ctrl_col.grid(row=0, column=2, sticky="nsew", padx=(5, 0))

        mode_frame = ttk.LabelFrame(ctrl_col, text=" 분석 모드 ", padding=8)
        mode_frame.pack(fill=tk.X, pady=(0, 5))

        self._mode_var = tk.StringVar(value="auto")
        modes = [
            ("자동  (txt→PDF 자동전환)", "auto"),
            ("텍스트만  (빠름)", "text_only"),
            ("PDF 원본  (정확, 느림)", "pdf_only"),
        ]
        for label, val in modes:
            tk.Radiobutton(
                mode_frame, text=label, variable=self._mode_var, value=val,
                font=("맑은 고딕", 8), bg="#f0f0f0",
            ).pack(anchor="w")

        extra_frame = ttk.LabelFrame(ctrl_col, text=" 추가 지시사항 (선택) ", padding=6)
        extra_frame.pack(fill=tk.BOTH, expand=True)

        tk.Label(
            extra_frame,
            text="예) 답변을 짧게, 영어로 작성, 코드 예시 필수 포함 등",
            font=("맑은 고딕", 7), fg="#aaa",
        ).pack(anchor="w", pady=(0, 3))

        self._txt_extra = tk.Text(
            extra_frame, height=3,
            font=("맑은 고딕", 9), relief=tk.FLAT, bg="white",
            wrap=tk.WORD, padx=6, pady=4, bd=1,
        )
        self._txt_extra.pack(fill=tk.BOTH, expand=True)

        # ── 생성 툴바 ────────────────────────────────────────
        gen_frame = tk.Frame(main, bg="#f0f0f0")
        gen_frame.pack(fill=tk.X, pady=(8, 4))

        self._btn_generate = tk.Button(
            gen_frame,
            text="✨  Gemini로 답변 초안 생성",
            command=self._generate,
            bg="#1a73e8", fg="white",
            font=("맑은 고딕", 11, "bold"),
            relief=tk.FLAT, cursor="hand2",
            padx=20, pady=8,
        )
        self._btn_generate.pack(side=tk.LEFT, padx=(0, 10))

        self._progress = ttk.Progressbar(gen_frame, mode="indeterminate", length=130)
        self._progress.pack(side=tk.LEFT, pady=6)

        self._lbl_status = tk.Label(gen_frame, text="", bg="#f0f0f0", font=("맑은 고딕", 9), fg="#555")
        self._lbl_status.pack(side=tk.LEFT, padx=8)

        self._btn_insert = tk.Button(
            gen_frame,
            text="📨  Outlook 답장에 삽입",
            command=self._insert_reply,
            bg="#34a853", fg="white",
            font=("맑은 고딕", 10, "bold"),
            relief=tk.FLAT, cursor="hand2",
            padx=14, pady=5,
            state=tk.DISABLED,
        )
        self._btn_insert.pack(side=tk.RIGHT, padx=(6, 0))
        ttk.Button(gen_frame, text="클립보드 복사", command=self._copy_clipboard).pack(side=tk.RIGHT, padx=4)
        ttk.Button(gen_frame, text="초안 지우기", command=self._clear_draft).pack(side=tk.RIGHT)

        # ── 참조 페이지 패널 (가로 스크롤) ──────────────────
        ref_outer = ttk.LabelFrame(main, text=" 📄  참조 페이지 (메일에 포함할 항목 체크) ", padding=4)
        ref_outer.pack(fill=tk.X, pady=(0, 4))

        self._lbl_ref_hint = tk.Label(
            ref_outer, text="초안 생성 후 참조 페이지가 표시됩니다.",
            font=("맑은 고딕", 8), fg="#aaa",
        )
        self._lbl_ref_hint.pack(side=tk.LEFT, padx=(0, 8))

        ref_canvas = tk.Canvas(ref_outer, height=30, highlightthickness=0, bg="#f0f0f0")
        ref_scroll = ttk.Scrollbar(ref_outer, orient=tk.HORIZONTAL, command=ref_canvas.xview)
        ref_canvas.configure(xscrollcommand=ref_scroll.set)
        ref_scroll.pack(side=tk.BOTTOM, fill=tk.X)
        ref_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self._ref_inner = tk.Frame(ref_canvas, bg="#f0f0f0")
        self._ref_canvas_win = ref_canvas.create_window((0, 0), window=self._ref_inner, anchor="nw")

        def _on_ref_inner_resize(_event):
            ref_canvas.configure(scrollregion=ref_canvas.bbox("all"))
        self._ref_inner.bind("<Configure>", _on_ref_inner_resize)

        # ── 분석 + 초안 영역 (좌우 PanedWindow) ─────────────
        paned = tk.PanedWindow(main, orient=tk.HORIZONTAL, sashrelief=tk.FLAT,
                               sashwidth=6, bg="#d0d0d0")
        paned.pack(fill=tk.BOTH, expand=True)

        analysis_frame = ttk.LabelFrame(paned, text=" 📋  문의 분석 및 답변 근거 ", padding=8)
        paned.add(analysis_frame, minsize=200, stretch="always")

        scroll_a = ttk.Scrollbar(analysis_frame, orient=tk.VERTICAL)
        scroll_a.pack(side=tk.RIGHT, fill=tk.Y)
        self._txt_analysis = tk.Text(
            analysis_frame,
            font=("맑은 고딕", 9),
            wrap=tk.WORD,
            yscrollcommand=scroll_a.set,
            relief=tk.FLAT, bg="#f8f9fa",
            padx=8, pady=6,
            state=tk.DISABLED,
        )
        self._txt_analysis.pack(fill=tk.BOTH, expand=True)
        scroll_a.config(command=self._txt_analysis.yview)

        draft_frame = ttk.LabelFrame(paned, text=" ✉️  이메일 초안 (편집 가능) ", padding=8)
        paned.add(draft_frame, minsize=200, stretch="always")

        scroll_y = ttk.Scrollbar(draft_frame, orient=tk.VERTICAL)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        self._txt_draft = tk.Text(
            draft_frame,
            font=("맑은 고딕", 10),
            wrap=tk.WORD,
            yscrollcommand=scroll_y.set,
            relief=tk.FLAT, bg="white",
            padx=8, pady=8,
        )
        self._txt_draft.pack(fill=tk.BOTH, expand=True)
        scroll_y.config(command=self._txt_draft.yview)

    # ── 프로필 ─────────────────────────────────────────────────

    def _reload_profiles(self):
        profiles = load_profiles()
        labels = [f"{p['name']}  /  {p['title']}" for p in profiles]
        self._cmb_profile["values"] = labels

        # 기본값 프로필 자동 선택
        default_idx = next((i for i, p in enumerate(profiles) if p.get("default")), 0 if profiles else None)
        if default_idx is not None and profiles:
            self._cmb_profile.current(default_idx)
            self._apply_profile(profiles[default_idx])

    def _on_profile_selected(self, _event=None):
        profiles = load_profiles()
        idx = self._cmb_profile.current()
        if 0 <= idx < len(profiles):
            self._apply_profile(profiles[idx])

    def _apply_profile(self, profile: dict):
        self._ent_resp_name.delete(0, tk.END)
        self._ent_resp_name.insert(0, profile.get("name", ""))
        self._ent_resp_title.delete(0, tk.END)
        self._ent_resp_title.insert(0, profile.get("title", ""))
        self._ent_resp_company.delete(0, tk.END)
        self._ent_resp_company.insert(0, profile.get("company", ""))
        self._ent_resp_dept.delete(0, tk.END)
        self._ent_resp_dept.insert(0, profile.get("department", ""))

    def _open_profile_dialog(self):
        ProfileDialog(self)

    # ── 메일 로드 ──────────────────────────────────────────────

    def _load_mail(self):
        self._set_status("Outlook 메일 불러오는 중...")
        self._btn_reload.config(state=tk.DISABLED)

        def _task():
            import pythoncom
            pythoncom.CoInitialize()
            try:
                from mail_extractor import extract_mail
                mail = self._mail = extract_mail()

                from manual_searcher import build_manual_context, get_matched_file_names
                query = mail.subject + " " + mail.body
                matched = get_matched_file_names(query)
                self._manual_context = build_manual_context(query)

                self.after(0, lambda: self._on_mail_loaded(mail, matched))
            except RuntimeError as e:
                self.after(0, lambda: self._on_mail_error(str(e)))
            finally:
                pythoncom.CoUninitialize()

        threading.Thread(target=_task, daemon=True).start()

    def _on_mail_loaded(self, mail, matched_files):
        self._lbl_subject.config(text=mail.subject)
        self._lbl_sender.config(text=mail.sender)

        if mail.attachments:
            names = [a.name for a in mail.attachments]
            if len(names) <= 3:
                attach_text = ", ".join(names)
            else:
                attach_text = ", ".join(names[:3]) + f"  외 {len(names) - 3}개"
            self._lbl_attach.config(text=attach_text, fg="#1a73e8")
        else:
            self._lbl_attach.config(text="없음", fg="#aaa")

        self._txt_body.config(state=tk.NORMAL)
        self._txt_body.delete("1.0", tk.END)
        preview_len = 2000
        preview = mail.body[:preview_len] + (f"\n\n... (이하 {len(mail.body)-preview_len}자 생략, Gemini에는 전문 전달)" if len(mail.body) > preview_len else "")
        self._txt_body.insert("1.0", preview)
        self._txt_body.config(state=tk.DISABLED)

        self._lst_manuals.delete(0, tk.END)
        for f in matched_files:
            self._lst_manuals.insert(tk.END, f"  {f}")

        self._set_status(f"메일 로드 완료  |  매뉴얼 {len(matched_files)}개 매칭")
        self._btn_reload.config(state=tk.NORMAL, text="⟳  메일 다시 불러오기")

    def _on_mail_error(self, msg):
        self._set_status("메일 로드 실패")
        self._btn_reload.config(state=tk.NORMAL)
        messagebox.showwarning("메일 없음", msg)

    # ── 초안 생성 ──────────────────────────────────────────────

    def _generate(self):
        if self._mail is None:
            messagebox.showwarning("알림", "먼저 메일을 불러오세요.")
            return

        self._btn_generate.config(state=tk.DISABLED)
        self._btn_insert.config(state=tk.DISABLED)
        self._progress.start(10)

        mode = self._mode_var.get()
        mode_label = {"auto": "자동", "text_only": "텍스트만", "pdf_only": "PDF 원본"}.get(mode, mode)
        self._set_status(f"Gemini API 호출 중 [{mode_label} 모드]...")

        responder_name    = self._ent_resp_name.get().strip()
        responder_title   = self._ent_resp_title.get().strip()
        responder_company = self._ent_resp_company.get().strip()
        responder_dept    = self._ent_resp_dept.get().strip()
        extra_instructions = self._txt_extra.get("1.0", tk.END).strip()

        def _task():
            try:
                from gemini_client import generate_reply
                mail = self._mail
                from manual_searcher import get_matched_pdfs
                query = self._mail.subject + " " + self._mail.body
                manual_pdfs = get_matched_pdfs(query) or None

                result = generate_reply(
                    mail_body=f"제목: {mail.subject}\n발신자: {mail.sender}\n\n{mail.body}",
                    manual_context=self._manual_context,
                    attachments=mail.attachments if mail.attachments else None,
                    manual_pdfs=manual_pdfs,
                    responder_name=responder_name,
                    responder_title=responder_title,
                    responder_company=responder_company,
                    responder_department=responder_dept,
                    mode=mode,
                    extra_instructions=extra_instructions,
                )
                self.after(0, self._on_draft_ready, result)
            except Exception as e:
                self.after(0, lambda: self._on_generate_error(str(e)))

        threading.Thread(target=_task, daemon=True).start()

    def _on_draft_ready(self, result):
        from gemini_client import ReplyResult
        self._progress.stop()
        self._btn_generate.config(state=tk.NORMAL)
        self._btn_insert.config(state=tk.NORMAL)
        self._set_status("초안 생성 완료  ✓")

        if isinstance(result, ReplyResult):
            self._txt_analysis.config(state=tk.NORMAL)
            self._txt_analysis.delete("1.0", tk.END)
            self._txt_analysis.insert("1.0", result.analysis)
            self._txt_analysis.config(state=tk.DISABLED)

            self._txt_draft.delete("1.0", tk.END)
            self._txt_draft.insert("1.0", result.draft)

            if result.used_pdf:
                self._populate_refs(result.page_refs)
            else:
                self._populate_refs([])  # 텍스트 모드: 참조 페이지 없음으로 표시
                self._lbl_ref_hint.config(
                    text="PDF 원본 모드에서만 참조 페이지가 표시됩니다.  "
                         "(현재: 텍스트 분석)"
                )
        else:
            self._txt_draft.delete("1.0", tk.END)
            self._txt_draft.insert("1.0", str(result))

        self._txt_draft.focus_set()

    def _populate_refs(self, refs: list[tuple[str, int]]):
        """참조 페이지 패널을 갱신."""
        for w in self._ref_inner.winfo_children():
            w.destroy()
        self._page_ref_vars.clear()

        if not refs:
            self._lbl_ref_hint.config(text="참조 페이지 없음  (도면·이미지 인용 불필요)")
            return

        self._lbl_ref_hint.config(text="")
        for filename, page_num in refs:
            var = tk.BooleanVar(value=True)
            chip = tk.Frame(self._ref_inner, bg="#e8edf5", bd=1, relief=tk.GROOVE)
            chip.pack(side=tk.LEFT, padx=(0, 6), pady=2)
            ttk.Checkbutton(chip, variable=var).pack(side=tk.LEFT, padx=(4, 0))
            tk.Label(
                chip, text=f"{filename}  p.{page_num}",
                font=("맑은 고딕", 8), bg="#e8edf5",
            ).pack(side=tk.LEFT, padx=(2, 4))
            ttk.Button(
                chip, text="미리보기", width=6,
                command=lambda f=filename, p=page_num: self._preview_page(f, p),
            ).pack(side=tk.LEFT, padx=(0, 4), pady=2)
            self._page_ref_vars.append((filename, page_num, var))

    def _preview_page(self, filename: str, page_num: int):
        """PDF 페이지를 새 창에서 미리보기."""
        import base64
        try:
            import config
            search_dirs = []
            if config.PDF_DIR:
                search_dirs.append(Path(config.PDF_DIR))
            if config.MANUAL_DIR:
                search_dirs.append(Path(config.MANUAL_DIR))

            from pdf_image_extractor import find_pdf, extract_page_as_png
            pdf_path = find_pdf(filename, search_dirs)
            if not pdf_path:
                messagebox.showwarning("파일 없음", f"PDF를 찾을 수 없습니다:\n{filename}\n\n설정 > 매뉴얼 관리에서 PDF 원본 폴더를 지정하세요.")
                return

            img_path = extract_page_as_png(pdf_path, page_num)

            with open(img_path, "rb") as f:
                png_bytes = f.read()
            photo = tk.PhotoImage(data=base64.b64encode(png_bytes))

            win = tk.Toplevel(self)
            win.title(f"{filename}  —  p.{page_num}")

            canvas = tk.Canvas(win, width=photo.width(), height=min(photo.height(), 720))
            scroll_y = ttk.Scrollbar(win, orient=tk.VERTICAL, command=canvas.yview)
            canvas.configure(
                yscrollcommand=scroll_y.set,
                scrollregion=(0, 0, photo.width(), photo.height()),
            )
            scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
            canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            canvas.create_image(0, 0, anchor="nw", image=photo)
            canvas._photo = photo  # GC 방지

        except ImportError as e:
            messagebox.showerror("패키지 없음", str(e))
        except Exception as e:
            messagebox.showerror("미리보기 오류", str(e))

    def _on_generate_error(self, msg: str):
        self._progress.stop()
        self._btn_generate.config(state=tk.NORMAL)
        self._set_status("오류 발생")
        messagebox.showerror("Gemini API 오류", msg)

    # ── Outlook 답장 삽입 ──────────────────────────────────────

    def _insert_reply(self):
        draft = self._txt_draft.get("1.0", tk.END).strip()
        if not draft:
            messagebox.showwarning("알림", "초안 내용이 비어 있습니다.")
            return

        # 체크된 참조 페이지를 PNG로 추출
        image_paths = []
        selected = [(f, p) for f, p, var in self._page_ref_vars if var.get()]
        if selected:
            try:
                import config
                search_dirs = []
                if config.PDF_DIR:
                    search_dirs.append(Path(config.PDF_DIR))
                if config.MANUAL_DIR:
                    search_dirs.append(Path(config.MANUAL_DIR))

                from pdf_image_extractor import find_pdf, extract_page_as_png
                for filename, page_num in selected:
                    pdf_path = find_pdf(filename, search_dirs)
                    if pdf_path:
                        image_paths.append(extract_page_as_png(pdf_path, page_num))
            except Exception as e:
                if not messagebox.askyesno(
                    "이미지 추출 오류",
                    f"페이지 이미지 추출 실패:\n{e}\n\n이미지 없이 삽입하시겠습니까?",
                ):
                    return
                image_paths = []

        try:
            from outlook_injector import inject_reply
            inject_reply(draft, image_paths=image_paths or None)
            self._set_status("답장 창에 삽입 완료  ✓")
        except Exception as e:
            messagebox.showerror("삽입 오류", str(e))

    # ── 유틸 ───────────────────────────────────────────────────

    def _clear_draft(self):
        self._txt_analysis.config(state=tk.NORMAL)
        self._txt_analysis.delete("1.0", tk.END)
        self._txt_analysis.config(state=tk.DISABLED)
        self._txt_draft.delete("1.0", tk.END)
        self._btn_insert.config(state=tk.DISABLED)

    def _copy_clipboard(self):
        draft = self._txt_draft.get("1.0", tk.END).strip()
        if draft:
            self.clipboard_clear()
            self.clipboard_append(draft)
            self._set_status("클립보드에 복사됨  ✓")

    def _set_status(self, msg: str):
        self._lbl_status.config(text=msg)


if __name__ == "__main__":
    app = App()
    app.mainloop()
