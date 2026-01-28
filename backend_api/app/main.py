from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from typing import Any, Dict

from .settings import settings
from .jobs import create_job, get_job, tail_log
from .excel_config import list_cuits, get_config_by_cuit, upsert_config_by_cuit
from .config_cuits import router as config_cuits_router

app = FastAPI(title="InvoicerPRO API", version="0.1.0")

app.include_router(config_cuits_router)

app.add_middleware(
    CORSMiddleware,
    allow_origins=[settings.cors_origin],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# =========================
# CONFIG (EXCEL)
# =========================

class ConfigUpdateRequest(BaseModel):
    data: Dict[str, Any] = {}

@app.get("/api/config/excel/cuits")
def api_config_cuits_excel():
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
    except FileNotFoundError as e:
        raise HTTPException(status_code=500, detail=str(e))
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))

# =========================
# JOBS
# =========================

class JobCreateRequest(BaseModel):
    mode: str = "local"
    cuit: int
    environment: str = "prod"  # prod por defecto

@app.get("/health")
def health():
    return {"ok": True, "service": "invoicerpro-api"}

@app.post("/jobs/generar")
def jobs_generar(req: JobCreateRequest):
    env = (req.environment or "prod").lower().strip()
    if env in ("produccion", "producción"):
        env = "prod"

    script = settings.generator_script_demo if env == "demo" else settings.generator_script_prod

    job = create_job(
        project_root=settings.project_root,
        script_path=script,
        jobs_dir=settings.jobs_dir,
        python_exe=settings.python_exe,
        script_args=["--cuit", str(req.cuit)],
    )
    return {"id": job.id, "status": job.status}

@app.get("/jobs/{job_id}")
def jobs_get(job_id: str):
    try:
        job = get_job(job_id)
        return {"id": job.id, "status": job.status, "exit_code": job.exit_code, "error": job.error}
    except KeyError:
        raise HTTPException(status_code=404, detail="Job not found")

@app.get("/jobs/{job_id}/log")
def jobs_log(job_id: str):
    try:
        return {"id": job_id, "log": tail_log(job_id)}
    except KeyError:
        raise HTTPException(status_code=404, detail="Job not found")
