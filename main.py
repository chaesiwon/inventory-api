"""
main.py - FastAPI 서버 (Render 클라우드 배포용)
인증: X-User-Id / X-Auth-Token 헤더 기반
CORS: allow_origins=["*"] (Vercel 등 모든 도메인 허용)
"""
import os, logging
from pathlib import Path
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse
from fastapi.requests import Request

logging.basicConfig(level=logging.INFO,
                    format="%(asctime)s %(levelname)s - %(message)s")
logger = logging.getLogger(__name__)

app = FastAPI(title="장기재고 소진계획 관리 시스템 API")

# ── CORS ─────────────────────────────────────
# allow_origins=["*"] → 모든 도메인 허용 (Vercel 포함)
# allow_credentials=False 필수 (wildcard 사용 시 규칙)
# 인증은 쿠키 대신 X-User-Id / X-Auth-Token 헤더로 처리
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["GET","POST","PUT","PATCH","DELETE","OPTIONS"],
    allow_headers=["*"],
    expose_headers=["Content-Disposition"],
)

# ── 세션 (선택적 - 헤더 인증 보조용) ──────────
try:
    from starlette.middleware.sessions import SessionMiddleware
    SECRET_KEY = os.environ.get(
        "SECRET_KEY", "inventory-secret-key-2024-please-change")
    app.add_middleware(SessionMiddleware,
                       secret_key=SECRET_KEY,
                       max_age=28800,
                       https_only=True,
                       same_site="none")
    logger.info("세션 미들웨어 활성화")
except Exception as e:
    logger.warning(f"세션 미들웨어 비활성 (헤더 인증으로 대체): {e}")

# ── API 라우터 ────────────────────────────────
from backend.api import router as api_router
app.include_router(api_router, prefix="/api")

# ── DB 초기화 ─────────────────────────────────
from backend.database import init_db
init_db()
logger.info("DB 초기화 완료")

# ── 헬스체크 ──────────────────────────────────
@app.get("/health")
async def health():
    return {"status": "ok"}

# ── 글로벌 예외 → 항상 JSON 반환 ───────────────
@app.exception_handler(Exception)
async def global_exc(request: Request, exc: Exception):
    import traceback
    logger.error(traceback.format_exc())
    return JSONResponse(
        status_code=500,
        content={"detail": str(exc), "ok": False}
    )
