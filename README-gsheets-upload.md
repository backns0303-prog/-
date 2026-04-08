# Google Sheets Upload

현재 폴더의 `.xls` 파일을 읽어서 Google Spreadsheet에 새 워크시트로 업로드합니다.

워크시트 이름 규칙:

- `공정 진행정보_YYYY-MM-DD_HHMM`
- `수주내역정보_YYYY-MM-DD_HHMM`
- `수주관리_YYYY-MM-DD_HHMM`

파일 분류 규칙:

- `관리번호`, `포장계획일`, `진행률`이 있으면 `공정 진행정보`
- `수주번호`, `주문일자`, `확정납기`, `단품코드`, `수주량`이 있으면 `수주내역정보`
- `대리점`, `CRM 고객코드`, `수주건명`이 있으면 `수주관리`

## 1. 패키지 설치

```powershell
python -m pip install -r requirements-gsheets.txt
```

## 2. 먼저 업로드 계획만 확인

```powershell
python upload_xls_to_gsheets.py `
  --credentials .\streamlit-sheets-upload-34b193fd0a59.json `
  --spreadsheet-id "여기에_구글시트_ID" `
  --pattern "*.xls" `
  --dry-run
```

## 3. 실제 업로드

```powershell
python upload_xls_to_gsheets.py `
  --credentials .\streamlit-sheets-upload-34b193fd0a59.json `
  --spreadsheet-id "여기에_구글시트_ID" `
  --pattern "*.xls"
```

## 4. 참고

- 같은 시점에 같은 이름의 워크시트가 이미 있으면 내용을 비우고 다시 채웁니다.
- 업로드 시점마다 새 시트를 남기려면 실행 시간을 다르게 주면 됩니다.
- 필요하면 `--uploaded-at 2026-03-26_1745`처럼 수동 지정도 가능합니다.
