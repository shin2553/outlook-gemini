# Outlook Gemini - 자동 답변 초안 생성

Outlook에서 고객 기술 문의 메일을 선택하면, Google Gemini AI가 매뉴얼을 참고하여 답변 초안을 자동으로 생성해주는 도구입니다.

## 주요 기능

- **Outlook 연동**: 현재 열린 메일을 자동으로 불러오기
- **매뉴얼 기반 답변**: txt 매뉴얼 검색 후 Gemini API로 분석
- **3가지 분석 모드**: 자동(txt→PDF 전환) / 텍스트만 / PDF 원본
- **참조 페이지 인용**: PDF 원본 페이지를 이미지로 메일에 삽입
- **추가 지시사항**: 답변 생성 전 조건/요청사항 자유 입력
- **출처 태그**: `[추정]` / `[확인 필요]` 자동 표기

## 요구사항

- Windows 10/11
- Microsoft Outlook 설치됨 (Office 2016 이상)
- Google Gemini API 키 ([발급 방법](https://aistudio.google.com/))

## 설치 (스크립트 실행 방식)

```bash
pip install -r requirements.txt
python ui.py
```

## 설치 (exe 빌드 방식)

```bash
build.bat
```
빌드 완료 후 `dist\OutlookGemini\` 폴더를 배포합니다.

## 초기 설정

1. `config_template.ini`를 복사하여 `config.ini`로 저장
2. `config.ini`에 Gemini API 키 입력
3. 프로그램 실행 후 **설정 > 매뉴얼 관리**에서 경로 지정

```ini
[gemini]
api_key = 여기에_API_키_입력

[paths]
manual_dir = C:/path/to/manual_txts
pdf_dir    = C:/path/to/pdf_originals
```

## 매뉴얼 시스템

- `manual_dir`: NotebookLM 등으로 변환한 txt 파일 폴더
- `manual_dir` 상위 또는 동일 폴더에 `manual_index.py` 필요
- PDF 원본은 설정 > 매뉴얼 관리에서 자동 변환 가능

## Outlook VBA 연동

`ribbon/` 폴더의 `GeminiReply.bas` 및 `설치방법.md` 참고

## 파일 구성

```
ui.py                  # 메인 UI
gemini_client.py       # Gemini API 클라이언트
mail_extractor.py      # Outlook 메일 추출
manual_searcher.py     # 매뉴얼 검색
outlook_injector.py    # Outlook 답장 삽입
pdf_converter.py       # PDF → txt 변환
pdf_image_extractor.py # PDF 페이지 → 이미지
config.py              # 설정 로더
config_template.ini    # 설정 템플릿 (config.ini로 복사하여 사용)
requirements.txt       # 필요 패키지
build.bat              # exe 빌드 스크립트
```

## 라이선스

MIT
