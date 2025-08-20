# Bulk DOCX → PDF → Email (Daum)

윈도우에 파이썬 설치 없이 실행 가능한 자동화 도구입니다.  
맥에서 개발/세팅 → GitHub Actions로 **윈도우용 단일 exe**를 빌드하여 배포합니다.

## 기능
- DOCX 템플릿(`{{name}}`, `{{line1}}` 등) 치환
- LibreOffice headless로 PDF 변환
- Daum(한메일) SMTP를 사용해 대량 이메일 발송

## 폴더 구조
```
bulk-docx-to-pdf-mail/
├─ src/                 # 파이썬 소스
├─ data/                # 템플릿/CSV
├─ windows/             # 윈도우 실행 스크립트/설정
└─ .github/workflows/   # GitHub Actions
```

## 준비물
1. **data/template.docx**: Word 템플릿. 바꿔야 하는 줄을 `{{line1}}`, `{{line2}}` 처럼 플레이스홀더로 작성
2. **data/data.csv**: 수신자/치환값 목록 (샘플 제공)
3. **windows/config.ini**: SMTP 계정/LibreOffice 경로 설정
4. **LibreOffice Portable** (무설치) 다운로드 후 exe와 같은 폴더에 두기  
   - 경로 예시: `LibreOfficePortable/App/libreoffice/program/soffice.exe`

## GitHub Actions로 윈도우 exe 빌드
1. 이 리포를 GitHub에 올립니다.
2. `Actions` 탭에서 **build-windows-exe** 워크플로 실행
3. Artifacts에서 `bulk_mailer_windows` 다운로드
4. 압축 풀면 `bulk_mailer.exe`, `data/`, `windows/` 포함

## 윈도우에서 실행
1. `bulk_mailer.exe`, `data/`, `windows/`, `LibreOfficePortable/`를 같은 폴더에 둡니다.
2. `windows/config.example.ini`를 `config.ini`로 복사 후 편집
3. `windows/run.bat` 실행(더블클릭)

## SMTP 설정 (Daum/hanmail)
- `HOST`: `smtp.daum.net` 또는 `smtp.hanmail.net`
- `PORT`: `465`
- `USER`: 전체 이메일 주소(예: `your_id@daum.net`)
- `PASS`: 계정 비밀번호
- `USE_SSL`: `true`

> 주의: 장기간 미접속 계정은 POP/IMAP/SMTP가 막혀 있을 수 있습니다. 먼저 웹메일 로그인으로 활성화하세요.

## 테스트 모드(DRY_RUN)
- `windows/config.ini`에서 `DRY_RUN=true`로 설정하면 **PDF 변환/메일 발송을 생략**하고 흐름만 검증합니다.

## CSV 예시
```csv
email,name,line1,line2
user1@daum.net,김민수,1번 줄 치환,2번 줄 치환
user2@hanmail.net,이서연,다른 내용1,다른 내용2
```

## 템플릿 권장 작성법
- 바꾸려는 “줄” 전체를 `{{lineN}}`로 둡니다.
- 줄 바꿈/문단 스타일은 템플릿에서 고정.

## 트러블슈팅
- `template.docx 없음` → `data/template.docx` 추가
- PDF 변환 실패 → `LibreOfficePortable` 경로 확인, 파일명에 한글/공백 이슈 시 짧은 경로 사용
- 메일 전송 실패 → 계정 활성화, HOST/PORT/USER/PASS/SSL 재확인, 과도 발송 시 `SLEEP_BETWEEN_MS` 1000~2000
