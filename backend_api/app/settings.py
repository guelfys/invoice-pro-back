from pydantic import BaseModel
import os
import sys

class Settings(BaseModel):
    cors_origin: str = os.getenv("CORS_ORIGIN", "http://localhost:5173")

    project_root: str = os.getenv(
        "INVOICER_PROJECT_ROOT",
        os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..")),
    )

    generator_script_demo: str = os.getenv("INVOICER_GENERATOR_SCRIPT_DEMO", "main_invoicerPRO.py")
    generator_script_prod: str = os.getenv("INVOICER_GENERATOR_SCRIPT_PROD", "main_Produccion_invoicerPRO.py")

    generator_script: str = os.getenv("INVOICER_GENERATOR_SCRIPT", "main_invoicerPRO.py")

    jobs_dir: str = os.getenv(
        "JOBS_DIR",
        os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "jobs")),
    )

    python_exe: str = os.getenv("INVOICER_PYTHON_EXE", sys.executable)

    config_excel_path: str = os.getenv(
        "INVOICER_CONFIG_XLSX",
        os.path.join(project_root, "Configuraciones", "Config.xlsx"),
    )

settings = Settings()
