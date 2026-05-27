"""
main.py - FastAPI (Render/Railway 클라우드 배포 전용)
- CORS 설정: Vercel 프론트엔드 도메인 허용
- 세션: itsdangerous 기반 (쿠키)
- 헬스체크: /health (Render 가동 확인)
"""
import os, logging
from pathlib import Path
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse
from fastapi.requests import Request

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(name)s - %(message)s"
)
logger = logging.getLogger(__name__)

BASE_DIR = Path(__file__).parent

app = FastAPI(
    title="장기재고 소진계획 관리 시스템 API",
    version="1.0.0",
)

# ── CORS (프론트 ↔ 백엔드 크로스 도메인 허용) ──
FRONTEND_URL = os.environ.get("FRONTEND_URL", "")
ALLOWED_ORIGINS = [
    "http://localhost:3000",
    "http://localhost:8000",
    "http://127.0.0.1:8000",
]
for o in FRONTEND_URL.split(","):
    o = o.strip().rstrip("/")
    if o and o not in ALLOWED_ORIGINS:
        ALLOWED_ORIGINS.append(o)

logger.info(f"CORS 허용: {ALLOWED_ORIGINS}")

app.add_middleware(
    CORSMiddleware,
    allow_origins=ALLOWED_ORIGINS,
    allow_credentials=True,
    allow_methods=["GET", "POST", "PUT", "PATCH", "DELETE", "OPTIONS"],
    allow_headers=["*"],
    expose_headers=["Content-Disposition"],  # 다운로드 파일명 노출
)

# ── 세션 미들웨어 ──
try:
    from starlette.middleware.sessions import SessionMiddleware
    SECRET_KEY = os.environ.get(
        "SECRET_KEY",
        "inventory-secret-default-change-in-prod-32chars!!"
    )
    app.add_middleware(
        SessionMiddleware,
        secret_key=SECRET_KEY,
        max_age=28800,
        https_only=False,
        same_site="none",  # 크로스 도메인 쿠키
    )
    logger.info("세션 미들웨어 활성화")
except Exception as e:
    logger.warning(f"세션 미들웨어 비활성화: {e}")

# ── API 라우터 ──
from backend.api import router as api_router
app.include_router(api_router, prefix="/api")

# ── 정적 파일 (선택적) ──
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

STATIC_DIR = BASE_DIR / "frontend" / "static"
TMPL_DIR   = BASE_DIR / "frontend" / "templates"

if STATIC_DIR.exists():
    app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")

templates = None
if TMPL_DIR.exists():
    templates = Jinja2Templates(directory=str(TMPL_DIR))

# ── DB 초기화 ──
from backend.database import init_db
try:
    init_db()
    logger.info("DB 초기화 완료")
except Exception as e:
    logger.error(f"DB 초기화 실패: {e}")
    raise

# ── 헬스체크 (Render 가동 확인용) ──
@app.get("/health")
async def health():
    return {"status": "ok", "message": "장기재고 관리 시스템 정상 운영 중"}

# ── 글로벌 예외 핸들러 ──
@app.exception_handler(Exception)
async def global_exception_handler(request: Request, exc: Exception):
    import traceback
    logger.error(f"처리되지 않은 오류:\n{traceback.format_exc()}")
    return JSONResponse(
        status_code=500,
        content={"detail": f"서버 내부 오류: {str(exc)}"}
    )

# ── SPA 라우팅 (프론트 통합 시만 사용) ──
if templates:
    from fastapi.responses import HTMLResponse

    @app.get("/", response_class=HTMLResponse)
    async def root(request: Request):
        return templates.TemplateResponse("index.html", {"request": request})

    @app.get("/{full_path:path}", response_class=HTMLResponse)
    async def spa(request: Request, full_path: str):
        skip = ("api/", "static/", "health", "docs", "openapi")
        if any(full_path.startswith(s) for s in skip):
            return JSONResponse(status_code=404, content={"detail": "Not Found"})
        return templates.TemplateResponse("index.html", {"request": request})
