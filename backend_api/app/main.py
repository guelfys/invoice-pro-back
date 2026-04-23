from fastapi import FastAPI, HTTPException, Request
from fastapi.responses import FileResponse, Response
from pathlib import Path
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from typing import Any, Dict

from .settings import settings
from .jobs import create_job, get_job, tail_log
from .excel_config import list_cuits, get_config_by_cuit, upsert_config_by_cuit
from .config_cuits import router as config_cuits_router
from .uploads_facturacion import router as uploads_facturacion_router

from pathlib import Path
from fastapi.responses import FileResponse, Response
from fastapi.staticfiles import StaticFiles
from starlette.requests import Request

app = FastAPI(title="InvoicerPRO API", version="0.1.0")

@app.middleware("http")
async def fix_double_api_prefix(request: Request, call_next):
    path = request.scope.get("path", "")
    if path.startswith("/api/api/"):
        request.scope["path"] = path.replace("/api/api/", "/api/", 1)
    return await call_next(request)

DIST_DIR = Path(__file__).resolve().parent / "web" / "dist"
FAVICON_PATH = DIST_DIR / "favicon.ico"

@app.get("/favicon.ico", include_in_schema=False)
def favicon():
    if FAVICON_PATH.exists():
        return FileResponse(FAVICON_PATH)
    return Response(status_code=204)

app.include_router(config_cuits_router)
app.include_router(uploads_facturacion_router)

app.add_middleware(
    CORSMiddleware,
    allow_origins=[settings.cors_origin],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

class ConfigUpdateRequest(BaseModel):
    data: Dict[str, Any] = {}

@app.get("/api/config/excel/cuits")
def api_config_excel_cuits():
    return {"cuits": list_cuits(settings.config_excel_path)}

@app.get("/api/config/excel/{tipo}/by-cuit/{cuit}")
def api_get_config_excel(tipo: str, cuit: int):
    try:
        return {"tipo": tipo.upper(), "config": get_config_by_cuit(settings.config_excel_path, tipo, cuit)}
    except KeyError:
        raise HTTPException(status_code=404, detail="CUIT no encontrado en Config.xlsx")
    except FileNotFoundError as e:
        raise HTTPException(status_code=500, detail=str(e))
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))

@app.put("/api/config/excel/{tipo}/by-cuit/{cuit}")
def api_put_config_excel(tipo: str, cuit: int, req: ConfigUpdateRequest):
    try:
        updated = upsert_config_by_cuit(settings.config_excel_path, tipo, cuit, req.data)
        return {"tipo": tipo.upper(), "config": updated}
    except PermissionError:
        raise HTTPException(status_code=409, detail="No se pudo guardar Config.xlsx (¿está abierto en Excel?). Cerralo y reintentá.")
    except FileNotFoundError as e:
        raise HTTPException(status_code=500, detail=str(e))
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))

class JobCreateRequest(BaseModel):
    mode: str = "local"
    cuit: int
    environment: str = "prod"  # demo | produccion


def _normalize_job_environment(value: str | None) -> str:
    env = (value or "produccion").lower().strip()
    if env in ("demo", "homo", "homologacion", "homologación", "test", "testing"):
        return "demo"
    if env in ("prod", "produccion", "producción", "production"):
        return "produccion"
    return "produccion"

@app.get("/health")
def health():
    return {"ok": True, "service": "invoicerpro-api"}

@app.post("/api/jobs/generar")
def jobs_generar(req: JobCreateRequest):
    env = _normalize_job_environment(req.environment)
    script = settings.generator_script_demo if env == "demo" else settings.generator_script_prod

    job = create_job(
        project_root=settings.project_root,
        script_path=script,
        jobs_dir=settings.jobs_dir,
        python_exe=settings.python_exe,
        script_args=["--cuit", str(req.cuit)],
        job_metadata={"environment": env},
    )
    return {"id": job.id, "status": job.status, "environment": env}

@app.get("/api/jobs/{job_id}")
def jobs_get(job_id: str):
    try:
        job = get_job(job_id)
        return {"id": job.id, "status": job.status, "exit_code": job.exit_code, "error": job.error}
    except KeyError:
        raise HTTPException(status_code=404, detail="Job not found")

@app.get("/api/jobs/{job_id}/log")
def jobs_log(job_id: str):
    try:
        return {"id": job_id, "log": tail_log(job_id)}
    except KeyError:
        raise HTTPException(status_code=404, detail="Job not found")


_dist_dir = Path(__file__).resolve().parent / "web" / "dist"
if _dist_dir.exists():

    app.mount("/", StaticFiles(directory=str(_dist_dir), html=False), name="frontend")

    @app.exception_handler(404)
    async def _spa_fallback(request: Request, exc): 
        path = request.url.path or "/"

        if path.startswith("/api") or path.startswith("/docs") or path.startswith("/openapi") or path.startswith("/redoc"):
            return Response("Not Found", status_code=404)

        last = path.rsplit("/", 1)[-1]
        if "." in last:
            return Response("Not Found", status_code=404)

        index = _dist_dir / "index.html"
        if index.exists():
            return FileResponse(str(index))
        return Response("Not Found", status_code=404)
