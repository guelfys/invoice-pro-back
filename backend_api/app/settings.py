from pydantic import BaseModel
import os
import sys
from pathlib import Path

def runtime_base_dir() -> Path:
    env = os.getenv("INVOICERPRO_BASE_DIR")
    if env:
        return Path(env).resolve()

    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent

    return Path(__file__).resolve().parents[2]

BASE_DIR = runtime_base_dir()
PROJECT_ROOT = Path(os.getenv("INVOICER_PROJECT_ROOT", str(BASE_DIR))).resolve()


class Settings(BaseModel):
    cors_origin: str = os.getenv("CORS_ORIGIN", "http://localhost:5173")

    project_root: str = str(PROJECT_ROOT)

    generator_script_demo: str = os.getenv("INVOICER_GENERATOR_SCRIPT_DEMO", "main_invoicerPRO.py")
    generator_script_prod: str = os.getenv("INVOICER_GENERATOR_SCRIPT_PROD", "main_Produccion_invoicerPRO.py")
    generator_script: str = os.getenv("INVOICER_GENERATOR_SCRIPT", "main_invoicerPRO.py")

    jobs_dir: str = os.getenv("JOBS_DIR", str(PROJECT_ROOT / "jobs"))

    python_exe: str = os.getenv("INVOICER_PYTHON_EXE", sys.executable)

    config_excel_path: str = os.getenv(
        "INVOICER_CONFIG_XLSX",
        str(PROJECT_ROOT / "configuraciones" / "Config.xlsx"),
    )

settings = Settings()