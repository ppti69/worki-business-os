# WORKI Production V3

## 포함
- 실제 로그인/권한
- SQLite DB 저장
- 고객사/사용자/자료/CRM/공고/문의/보고/자동화 CRUD
- PC 파일 업로드/다운로드
- 최대 업로드 용량 환경설정
- 고객사별 OpenAI Vector Store
- Responses API + file_search RAG
- AI 사용량 기록
- 감사로그
- Docker / Render 배포파일

## 로컬 Docker
1. `.env.example`을 `.env`로 복사하고 값 입력
2. `docker build -t worki .`
3. `docker run --env-file .env -p 3000:3000 -v worki-data:/data worki`
4. http://localhost:3000

## Render
1. 이 폴더를 GitHub 저장소에 올림
2. Render에서 Blueprint/New Web Service로 저장소 연결
3. `OPENAI_API_KEY`, `ADMIN_PASSWORD` 입력
4. Persistent Disk가 `/data`에 연결됐는지 확인
5. 배포 후 관리자 이메일로 로그인

## 기본 관리자
ADMIN_EMAIL / ADMIN_PASSWORD 환경변수 값.
환경변수를 넣지 않은 로컬 기본값:
admin@worki.local / ChangeMe123!

## 샘플 고객
ceo@abcbuilder.demo / 1234

## 중요
실제 서비스 공개 전 관리자 비밀번호를 반드시 변경하세요.
