"""
main.py - FastAPI 앱 (클라우드 배포 버전)
Render / Railway 등 클라우드 플랫폼에서 실행
CORS: 프론트엔드 도메인(Vercel)에서 API 호출 허용
"""
import os, logging
from pathlib import Path
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from fastapi.requests import Request
from fastapi.responses import HTMLResponse, JSONResponse

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
logger = logging.getLogger(__name__)

BASE_DIR = Path(__file__).parent

app = FastAPI(
    title="장기재고 소진계획 관리 시스템",
    version="1.0.0",
    docs_url="/docs",           # Swagger UI (개발 시 유용)
    redoc_url=None,
)

# ── CORS 설정 ─────────────────────────────────
# FRONTEND_URL: Vercel 배포 URL (환경변수로 주입)
# 예) https://inventory-app.vercel.app
FRONTEND_ORIGINS_ENV = os.environ.get("FRONTEND_URL", "")
ALLOWED_ORIGINS = [
    "http://localhost:3000",
    "http://localhost:8000",
    "http://127.0.0.1:8000",
]
# 환경변수에서 추가 도메인 읽기 (쉼표 구분)
for origin in FRONTEND_ORIGINS_ENV.split(","):
    o = origin.strip()
    if o and o not in ALLOWED_ORIGINS:
        ALLOWED_ORIGINS.append(o)

app.add_middleware(
    CORSMiddleware,
    allow_origins=ALLOWED_ORIGINS,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)
logger.info(f"CORS 허용 도메인: {ALLOWED_ORIGINS}")

# ── 세션 미들웨어 ──────────────────────────────
try:
    from starlette.middleware.sessions import SessionMiddleware
    SECRET_KEY = os.environ.get("SECRET_KEY", "inventory-secret-key-2024-change-in-prod-!!!")
    app.add_middleware(
        SessionMiddleware,
        secret_key=SECRET_KEY,
        max_age=28800,
        https_only=False,
        same_site="none",       # 크로스 도메인 쿠키 허용 (Vercel↔Render)
    )
except ImportError:
    logger.warning("starlette 미설치 - 세션 비활성화")

# ── API 라우터 ────────────────────────────────
from backend.api import router as api_router
app.include_router(api_router, prefix="/api")

# ── 정적 파일 (프론트 빌드 산출물 - 선택적) ──
STATIC_DIR = BASE_DIR / "frontend" / "static"
if STATIC_DIR.exists():
    app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")

# ── 템플릿 (선택적 - 프론트가 분리된 경우 불필요) ──
TMPL_DIR = BASE_DIR / "frontend" / "templates"
templates = None
if TMPL_DIR.exists():
    templates = Jinja2Templates(directory=str(TMPL_DIR))

# ── DB 초기화 ─────────────────────────────────
from backend.database import init_db
init_db()
logger.info("DB 초기화 완료")

# ── 헬스체크 (Render/Railway 가동 확인용) ──────
@app.get("/health")
async def health():
    return {"status": "ok", "service": "장기재고 관리 시스템"}

# ── 글로벌 예외 핸들러 ────────────────────────
@app.exception_handler(Exception)
async def global_exception_handler(request: Request, exc: Exception):
    import traceback
    logger.error(f"Unhandled error: {exc}\n{traceback.format_exc()}")
    return JSONResponse(status_code=500, content={"detail": f"서버 오류: {str(exc)}"})

# ── SPA 라우팅 (선택적 - 프론트 통합 시) ────────
if templates:
    @app.get("/", response_class=HTMLResponse)
    async def root(request: Request):
        return templates.TemplateResponse("index.html", {"request": request})

    @app.get("/{full_path:path}", response_class=HTMLResponse)
    async def spa(request: Request, full_path: str):
        if full_path.startswith(("api/", "static/", "health")):
            return JSONResponse(status_code=404, content={"detail": "Not Found"})
        return templates.TemplateResponse("index.html", {"request": request})
