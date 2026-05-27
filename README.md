# 장기재고 소진계획 관리 시스템 - API 백엔드

Python FastAPI 백엔드. Render 또는 Railway에 배포됩니다.

## 로컬 실행 (개발 시)

```cmd
py -m pip install -r requirements.txt
py migrate.py
py -m uvicorn main:app --port 8000
```

## Render 배포 방법

1. https://render.com 접속 → GitHub로 로그인
2. "New +" → "Web Service"
3. GitHub 저장소 연결
4. 설정:
   - **Runtime**: Python 3
   - **Build Command**: `pip install -r requirements.txt && python migrate.py`
   - **Start Command**: `uvicorn main:app --host 0.0.0.0 --port $PORT`
5. Environment Variables 추가:
   - `SECRET_KEY`: 랜덤 문자열 (생성 클릭)
   - `FRONTEND_URL`: Vercel 배포 URL
6. **Add Disk** (데이터 영구 저장):
   - Name: `inventory-data`
   - Mount Path: `/opt/render/project/src/data`
   - Size: 1 GB
7. Deploy 클릭

## Railway 배포 방법

1. https://railway.app 접속 → GitHub로 로그인
2. "New Project" → "Deploy from GitHub repo"
3. 저장소 선택 → 자동 배포
4. Variables 탭에서 환경변수 추가
5. 볼륨 추가 (데이터 영구 저장)

## 환경변수

| 변수 | 설명 | 예시 |
|------|------|------|
| `SECRET_KEY` | 세션 암호화 키 | 랜덤 32자 문자열 |
| `FRONTEND_URL` | Vercel 배포 URL | `https://app.vercel.app` |
| `PORT` | 서버 포트 (자동 주입) | `8000` |
