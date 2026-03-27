# Outlook + Gemini API 자동 답변 초안 생성 프로젝트

## 프로젝트 목적
여러 제조사 제품(SCANLAB 등)의 고객 기술 문의 메일을 Outlook에서 선택하면,
Gemini API가 매뉴얼 기반으로 답변 초안을 자동 생성하여 Outlook 답장 창에 입력해주는 도구.

---

## 개발 환경
- OS: Windows 11
- Python: 설치됨
- Outlook: Microsoft Office Business 2021 (독립 설치형)
- Gemini API 키: 발급 완료 (Google AI Studio)
- 모델: `gemini-2.5-flash` (thinking_budget=10000)

---

## 파일 구성

| 파일 | 역할 |
|---|---|
| `ui.py` | 메인 tkinter UI |
| `gemini_client.py` | Gemini API 호출, 응답 파싱 |
| `mail_extractor.py` | Outlook COM으로 메일 추출 |
| `manual_searcher.py` | 매뉴얼 txt 검색, PDF 경로 반환 |
| `outlook_injector.py` | Outlook 답장창에 초안 + 이미지 삽입 |
| `pdf_converter.py` | PDF → txt 변환 |
| `pdf_image_extractor.py` | PDF 페이지 → PNG 렌더링 (pymupdf) |
| `config.py` | config.ini 값 로드, 상대경로 해석 |
| `config.ini` | API 키, 경로, 설정값 (상대경로 지원) |
| `config_template.ini` | 배포용 빈 설정 템플릿 |
| `profiles.json` | 답변 주체 프로필 목록 |
| `manual_index.py` | 키워드→매뉴얼 매핑 인덱스, search_manuals() 제공 |
| `hook_win32com.py` | PyInstaller 런타임 훅 — win32com gencache 경로 수정 |
| `build.bat` | PyInstaller 빌드 스크립트 (CRLF, 매뉴얼 복사 포함) |
| `outlook_gemini.spec` | PyInstaller 빌드 설정 |
| `Manuals_txt/` | 텍스트 변환된 매뉴얼 (.txt) |
| `Manuals/` | PDF 원본 매뉴얼 |

---

## 동작 흐름
1. 답변 주체 프로필 선택 → 메일 불러오기 버튼 클릭
2. Outlook COM으로 선택된 메일 본문 + 첨부파일(이미지/PDF) 추출
3. manual_index.py로 관련 매뉴얼 txt 검색 → 컨텍스트 구성
4. (선택) 추가 지시사항 입력, 분석 모드 선택
5. Gemini API 호출 → 분석 + 이메일 초안 생성
6. 참조 페이지 체크 → Outlook 답장 창에 초안 + 인용 이미지 삽입
7. 담당자가 검토 후 발송

---

## 기술 스택
- **Python**: 메일 추출 (win32com/pythoncom), Gemini API 호출, 초안 Outlook 주입
- **Gemini API**: 텍스트 + 이미지 + PDF 멀티모달, thinking 모드
- **매뉴얼 시스템**: manual_index.py + Manuals_txt/ txt 파일
- **pymupdf(fitz)**: PDF 페이지 → PNG 렌더링, PDF 서브셋 추출
- **PyInstaller**: 단일 폴더 실행 파일 빌드 (onedir)

---

## 분석 모드 (3가지)
| 모드 | 설명 |
|---|---|
| `auto` | txt 우선 → 시각 정보 필요 시 PDF 자동 전환 (`VISUAL_NEEDED` 마커) |
| `text_only` | 항상 txt만 사용 (빠름) |
| `pdf_only` | 처음부터 PDF 원본 직접 분석 (정확, 느림) |

- PDF가 900페이지 초과 시 `_make_pdf_subset()`으로 관련 페이지(최대 200p)만 추출하여 전달
- PDF 서브셋 임시 파일은 API 호출 완료 후 자동 삭제

---

## Gemini 응답 구조
```
===분석===
**문의 요약** / **관련 매뉴얼 및 근거** / **답변 방향** / **주의/확인 필요 사항**

===이메일초안===
(마크다운 없는 평문 이메일)

===페이지참조===
[REF] 파일명.txt | 페이지번호
```
- `ReplyResult` dataclass: `analysis`, `draft`, `page_refs: list[tuple[str,int]]`, `used_pdf: bool`
- `page_refs` / 참조 페이지 UI는 `used_pdf=True`(PDF 모드)일 때만 표시

---

## 정보 출처 태그 원칙
- 매뉴얼 근거 있음 → 그대로 서술
- 일반 기술 지식으로 보충 → `[추정]` 태그
- 불확실 → `[확인 필요]` 태그 (이메일 초안에도 동일 표기)

---

## 추가 지시사항
- UI에서 자유 텍스트로 입력
- `## 추가 지시사항 (최우선 적용 — 기본 답변 원칙보다 우선합니다)` 섹션으로 프롬프트에 포함
- 텍스트/PDF 모드 모두 적용

---

## 지원 첨부파일 형식
- 이미지 (jpg, png 등) ✅ — Gemini 멀티모달로 직접 전달
- PDF ✅ — Gemini에 직접 전달 또는 서브셋 추출
- Word/Excel ❌ (미지원)

---

## 인라인 이미지 삽입 (Outlook)
- 참조 페이지 체크박스에서 선택한 페이지를 PNG로 렌더링
- `outlook_injector.inject_reply(image_paths=[...])` 로 CID 방식 인라인 삽입
- MAPI 속성: `PR_ATTACH_CONTENT_ID = "http://schemas.microsoft.com/mapi/proptag/0x3712001E"`
- win32com 백그라운드 스레드에서 반드시 `pythoncom.CoInitialize()` / `CoUninitialize()` 필요

---

## PDF 변환 / 관리
- PDF → txt: `pdf_converter.convert_pdf_to_txt()` → `manual_dir`에 동일 stem으로 저장
- 미변환 PDF 자동 탐지: `pdf_dir` 스캔 후 txt 없는 파일 목록 표시
- 설정 > 매뉴얼 관리에서 일괄 자동 변환 또는 직접 선택 변환 가능
- 중복 탐지: `glob("*.pdf") + glob("*.PDF")` 후 `seen` set으로 stem 기준 중복 제거
- 페이지별 텍스트 추출 실패는 무시하고 계속 진행 (스캔 혼합 PDF 대응)

---

## PDF 파일 탐색 (pdf_image_extractor.find_pdf)
1. 정확한 파일명 매칭
2. stem + `.pdf` / `.txt`
3. 키워드 스코어링 (2자 이상 단어 매칭 수) + stem 길이 오름차순 tiebreaker

---

## UI 레이아웃 구조
```
[헤더: 타이틀 + 설정 버튼]
[답변 주체: 프로필/이름/직책/회사/부서/프로필관리 + 메일불러오기(우측)]
[3단: 메일정보(2) | 매칭된매뉴얼(1) | 분석모드+추가지시사항(1)]
[생성 툴바: 초안생성버튼 + 진행바 + 상태 | 지우기/복사/Outlook삽입]
[참조 페이지: 가로 스크롤 chip 목록 (PDF 모드 시만 표시)]
[PanedWindow: 문의분석(좌) | 이메일초안(편집가능)(우)]
```
- 3단 컬럼 비율: `columnconfigure(uniform="mid")` weight=2:1:1 강제
- 본문 미리보기: `rowconfigure(3, weight=1)` + `sticky="nsew"` 로 세로 확장
- 프로필 관리 다이얼로그: `420x400`

---

## 설정 (config.ini)
```ini
[gemini]
api_key = ...

[paths]
manual_index = manual_index.py      ← 상대경로 또는 절대경로 모두 지원
manual_dir = Manuals_txt
pdf_dir = Manuals

[settings]
max_manual_context = 80000
language = ko

[company]
role = 여러 제조사 제품의 국내 공식 대리점
```
- 상대경로는 `config.py`의 `_resolve()`가 `get_app_dir()` 기준 절대경로로 변환
- `manual_index.py` 자동 탐색: `MANUAL_INDEX_PATH` → `manual_dir` 상위 → `manual_dir` 내부 순

---

## 매뉴얼 경로 (프로젝트 내 관리)
- 매뉴얼 인덱스: `outlook_gemini/manual_index.py`
- 텍스트 매뉴얼: `outlook_gemini/Manuals_txt/`
- PDF 원본: `outlook_gemini/Manuals/`
- 빌드 시 `build.bat`이 자동으로 `dist/OutlookGemini/` 하위에 복사

---

## manual_index.py 구조
- `EXPORT_DIR`: 매뉴얼 txt 폴더 경로 (manual_searcher.py가 config.MANUAL_DIR로 덮어씀)
- `INDEX`: 키워드 → txt 파일명 목록 딕셔너리
- `search_manuals(query)`: 1) INDEX 키워드 매칭 → 2) 미등록 파일 파일명 기반 자동 탐지
- 새 제품 매뉴얼 추가: txt 파일을 Manuals_txt/에 넣고 INDEX에 키워드 등록

---

## manual_searcher.py 로딩 방식
- `importlib.util.spec_from_file_location()`으로 외부 .py 파일 직접 로드
  → PyInstaller 실행 파일 환경에서 `sys.path` 방식이 동작하지 않는 문제 우회
- 로드 후 `mod.EXPORT_DIR = Path(config.MANUAL_DIR)` 주입 → 폴더명 변경에 유연하게 대응

---

## gemini_client.py 클라이언트 관리
- `_client: list = []` — lazy initialization (모듈 로드 시 API 키 불필요)
- `_get_client()`: API 키 없으면 명확한 ValueError 발생
- `reset_client()`: API 키 변경 시 호출하여 클라이언트 재초기화
- `ui.py`의 `_save_api()`에서 저장 후 `reset_client()` 호출

---

## PyInstaller 빌드
- `build.bat` 실행 (CRLF 인코딩 필수 — LF이면 CMD 오파싱 발생)
- `hook_win32com.py`: win32com gencache를 `%TEMP%/outlook_gemini_gen_py`로 리다이렉트
- `outlook_gemini.spec` hidden imports: `win32com.client.dynamic`, `win32com.client.gencache` 등
- `build.bat` 빌드 후 복사 항목:
  - `config_template.ini` → `dist/config.ini` (API 키·경로 빈칸)
  - `manual_index.py`, `Manuals_txt/`, `Manuals/` → dist 하위
  - `profiles.json` (기존 파일 있으면 덮어쓰지 않음)
- 빌드 전 `OutlookGemini.exe` 자동 종료 (`taskkill`)

---

## 주의사항
- Gemini API 키는 config.ini로 관리 (코드에 하드코딩 금지)
- config_template.ini는 배포용 — API 키와 경로 모두 빈칸 유지
- 답변은 항상 검토 후 발송 (완전 자동발송 아님)
- win32com 사용 스레드에서 반드시 `pythoncom.CoInitialize()` 호출
- Gemini PDF 업로드 한도: 1000페이지 → 초과 시 자동 서브셋 추출로 처리
- 메일 재로드 시 이전 `MailData.cleanup()` 자동 호출 (임시 파일 정리)
- 앱 종료 후 백그라운드 스레드 콜백은 `winfo_exists()` 체크로 방지
