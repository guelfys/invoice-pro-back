from __future__ import annotations

from fastapi import APIRouter, UploadFile, File, Form, HTTPException
from typing import Dict, List, Optional
from pathlib import Path
import os
import shutil

from .settings import settings

router = APIRouter()

VALID_TIPOS = {"A", "B", "C"}

def _input_dir_for_tipo(tipo: str) -> Path:
    return Path(settings.project_root) / "input" / f"Factura {tipo}"

def _facturacion_path(tipo: str) -> Path:
    return _input_dir_for_tipo(tipo) / "Facturacion.xlsx"

def _validate_xlsx(upload: UploadFile, tipo: str) -> None:
    name = (upload.filename or "").lower().strip()
    if not name.endswith(".xlsx"):
        raise HTTPException(
            status_code=400,
            detail=f"El archivo para Factura {tipo} debe ser .xlsx (recibido: {upload.filename})",
        )

@router.post("/api/generar/facturacion/upload")
async def upload_facturacion(
    tipo: Optional[str] = Form(default=None),
    file: Optional[UploadFile] = File(default=None),

    fileA: Optional[UploadFile] = File(default=None),
    fileB: Optional[UploadFile] = File(default=None),
    fileC: Optional[UploadFile] = File(default=None),

    clear_others: bool = Form(default=True),
):
    files_map: Dict[str, UploadFile] = {}

    if tipo and file:
        t = tipo.upper().strip()
        if t not in VALID_TIPOS:
            raise HTTPException(status_code=400, detail="tipo inválido. Use A, B o C.")
        files_map[t] = file

    if fileA: files_map["A"] = fileA
    if fileB: files_map["B"] = fileB
    if fileC: files_map["C"] = fileC

    if not files_map:
        raise HTTPException(status_code=400, detail="No se recibió ningún archivo (.xlsx).")

    saved: List[dict] = []
    cleared: List[dict] = []

    if clear_others:
        to_clear = VALID_TIPOS - set(files_map.keys())
        for t in to_clear:
            p = _facturacion_path(t)
            if p.exists():
                try:
                    p.unlink()
                    cleared.append({"tipo": t, "path": str(p)})
                except Exception as e:
                    raise HTTPException(status_code=500, detail=f"No se pudo borrar {p}: {e}")

    for t, upload in files_map.items():
        _validate_xlsx(upload, t)
        out_dir = _input_dir_for_tipo(t)
        out_dir.mkdir(parents=True, exist_ok=True)
        out_path = _facturacion_path(t)

        try:
            with out_path.open("wb") as f:
                await upload.seek(0)
                shutil.copyfileobj(upload.file, f)
        except Exception as e:
            raise HTTPException(status_code=500, detail=f"No se pudo guardar Facturacion.xlsx para {t}: {e}")
        finally:
            try:
                await upload.close()
            except Exception:
                pass

        saved.append({"tipo": t, "path": str(out_path)})

    return {"ok": True, "saved": saved, "cleared": cleared}
