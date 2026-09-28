from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from starlette.middleware.base import BaseHTTPMiddleware
from starlette.requests import Request
from fastapi.staticfiles import StaticFiles
from pathlib import Path
from app.core.config import get_settings
from app.core.database import Base, SessionLocal, engine
from app.core.security import hash_password
from app.api.routes import router
from app.api.connectors import router as connector_router
from app.models.entities import User
from app.models import canonical
from app.services.engine import uid
from sqlalchemy import func, select

s=get_settings()
Base.metadata.create_all(engine)
if s.bootstrap_admin_username and s.bootstrap_admin_password:
    if len(s.bootstrap_admin_password) < 12:
        raise RuntimeError("BOOTSTRAP_ADMIN_PASSWORD must contain at least 12 characters")
    with SessionLocal() as db:
        if db.scalar(select(func.count()).select_from(User)) == 0:
            db.add(User(
                id=uid("USR"),
                username=s.bootstrap_admin_username,
                password_hash=hash_password(s.bootstrap_admin_password),
                role="ADMIN",
            ))
            db.commit()
app=FastAPI(title=s.app_name,version="2.3.0")
app.add_middleware(CORSMiddleware,allow_origins=s.origins,allow_credentials=True,allow_methods=["*"],allow_headers=["*"])

class SecurityHeaders(BaseHTTPMiddleware):
    async def dispatch(self, request: Request, call_next):
        response=await call_next(request)
        response.headers["X-Content-Type-Options"]="nosniff"
        response.headers["X-Frame-Options"]="DENY"
        response.headers["Referrer-Policy"]="no-referrer"
        if request.url.path.startswith("/docs") or request.url.path.startswith("/redoc"):
            response.headers["Content-Security-Policy"]=("default-src 'self' https: data:; script-src 'self' https://cdn.jsdelivr.net 'unsafe-inline'; style-src 'self' https://cdn.jsdelivr.net 'unsafe-inline'; img-src 'self' data: https:; font-src 'self' data: https:; frame-ancestors 'none';")
        else:
            response.headers["Content-Security-Policy"]="default-src 'self'; frame-ancestors 'none'"
        return response
app.add_middleware(SecurityHeaders)
from fastapi.responses import JSONResponse
from fastapi.exceptions import RequestValidationError

@app.exception_handler(RequestValidationError)
async def validation_exception_handler(request: Request, exc: RequestValidationError):
    return JSONResponse(
        status_code=422,
        content={"detail": "Request validation failed", "errors": exc.errors(), "remediation": "Verify input types and required fields."}
    )

@app.exception_handler(ValueError)
async def value_error_handler(request: Request, exc: ValueError):
    return JSONResponse(
        status_code=400,
        content={"detail": str(exc), "remediation": "Check project configuration, medallion plan, or source connection parameters."}
    )

@app.exception_handler(RuntimeError)
async def runtime_error_handler(request: Request, exc: RuntimeError):
    msg = str(exc)
    remediation = "Ensure local connector or database services are running and accessible."
    if "CONNECTOR_OFFLINE" in msg:
        remediation = "Start the local connector script in a terminal, verify the source ID, and ensure heartbeat is active."
    elif "Databricks" in msg:
        remediation = "Check your Databricks workspace URL, SQL warehouse HTTP path, and access token."
    return JSONResponse(status_code=400, content={"detail": msg, "remediation": remediation})

app.include_router(router)
app.include_router(connector_router)
@app.get("/health", tags=["System"])
def root_health(): return {"status":"ok","service":"migration-factory"}

frontend_dir = Path("/app/frontend_dist")
if not frontend_dir.exists():
    frontend_dir = Path(__file__).resolve().parents[2] / "frontend" / "dist"
if not frontend_dir.exists():
    frontend_dir = Path(__file__).resolve().parents[2] / "frontend_dist"
if frontend_dir.exists():
    app.mount("/", StaticFiles(directory=frontend_dir, html=True), name="frontend")
