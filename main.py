"""
main.py - FastAPI 서버 (Render 클라우드 배포용)
"""
import os, logging
from pathlib import Path
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse
from fastapi.requests import Request

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s - %(message)s")
logger = logging.getLogger(__name__)

app = FastAPI(title="장기재고 소진계획 관리 시스템 API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=[
        "https://inventory-pwa-81en.vercel.app",  # 실제 Vercel 배포 주소
        "http://localhost:3000",                   # 로컬 개발용
        "http://127.0.0.1:3000",
    ],
    allow_credentials=True,
    allow_methods=["GET","POST","PUT","PATCH","DELETE","OPTIONS"],
    allow_headers=["*"],
    expose_headers=["Content-Disposition"],
)

try:
    from starlette.middleware.sessions import SessionMiddleware
    SECRET_KEY = os.environ.get("SECRET_KEY", "inventory-secret-key-2024-please-change")
    HTTPS_ONLY = os.environ.get("SESSION_HTTPS_ONLY", "true").lower() == "true"
    app.add_middleware(SessionMiddleware, secret_key=SECRET_KEY, max_age=28800,
                       https_only=HTTPS_ONLY, same_site="none" if HTTPS_ONLY else "lax")
    logger.info("세션 미들웨어 활성화")
except Exception as e:
    logger.warning(f"세션 미들웨어 비활성: {e}")

from backend.api import router as api_router
app.include_router(api_router, prefix="/api")

from backend.database import init_db
init_db()
logger.info("DB 초기화 완료")

@app.get("/health")
async def health():
    return {"status": "ok"}

@app.exception_handler(Exception)
async def global_exc(request: Request, exc: Exception):
    import traceback
    logger.error(traceback.format_exc())
    return JSONResponse(status_code=500, content={"detail": str(exc), "ok": False})
