from fastapi import APIRouter, UploadFile, File, Form, HTTPException
from pathlib import Path
import re
import shutil
import json
from datetime import datetime
from pydantic import BaseModel
from .excel_config import sync_row2_all_tipos
from zeep.helpers import serialize_object

router = APIRouter(prefix="/api/config", tags=["config"])

import subprocess
from xml.etree import ElementTree as ET
from zeep import Client
from zeep.transports import Transport
import requests

PADRON_A5_WSDL = "https://aws.afip.gov.ar/sr-padron/webservices/personaServiceA5?wsdl"  # :contentReference[oaicite:4]{index=4}

def get_cert_key_paths_for_env(cuit: str, env: str) -> tuple[Path, Path]:
    fev1_source = USUARIOS_DIR / cuit / SERVICES["FEV1"]["auth_dir"] / "source"

    cert_name = ENV_FILES[env]["cert_name"]
    key_name = ENV_FILES[env]["key_names"][0]

    cert_path = fev1_source / cert_name
    key_path = fev1_source / key_name

    if not cert_path.exists():
        raise HTTPException(status_code=400, detail=f"No se encontró certificado en {cert_path}")
    if not key_path.exists():
        raise HTTPException(status_code=400, detail=f"No se encontró key en {key_path}")

    return cert_path, key_path


def obtener_token_sign_con_wsaa_cliente(env: str, servicio: str, cert_path: Path, key_path: Path) -> tuple[str, str]:
    env = norm_env(env)

    svc = (servicio or "").strip().lower()
    if svc in ("padron", "padron_a5", "a5"):
        service_id = "ws_sr_padron_a5"
    elif svc in ("constancia", "constancia_inscripcion", "constancia-inscripcion"):
        service_id = "ws_sr_constancia_inscripcion"
    else:
        service_id = (servicio or "").strip()

    wsaa_wsdl = "https://wsaahomo.afip.gov.ar/ws/services/LoginCms?WSDL" if env == "demo" else "https://wsaa.afip.gov.ar/ws/services/LoginCms?WSDL"

    ps1_candidates = [
        cert_path.parent / "wsaa-cliente.ps1",
        cert_path.parent / "wsaa_cliente.ps1",
        BASE_DIR / "wsaa-cliente.ps1",
        BASE_DIR / "wsaa_cliente.ps1",
    ]
    ps1_path = next((p for p in ps1_candidates if p.exists()), None)
    if not ps1_path:
        raise HTTPException(
            status_code=500,
            detail=(
                "No se encontró wsaa-cliente.ps1. "
                f"Busqué en: {', '.join(str(p) for p in ps1_candidates)}"
            ),
        )

    run_dir = cert_path.parent / "_wsaa_tmp"
    run_dir.mkdir(parents=True, exist_ok=True)

    cmd = [
        "powershell",
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        str(ps1_path),
        "-Certificado",
        str(cert_path),
        "-ClavePrivada",
        str(key_path),
        "-ServicioId",
        service_id,
        "-WsaaWsdl",
        wsaa_wsdl,
    ]

    try:
        completed = subprocess.run(
            cmd,
            cwd=str(run_dir),
            capture_output=True,
            text=True,
            timeout=90,
            check=False,
        )
    except FileNotFoundError:
        raise HTTPException(status_code=500, detail="No se pudo ejecutar PowerShell (no encontrado).")
    except subprocess.TimeoutExpired:
        raise HTTPException(status_code=504, detail="Timeout ejecutando wsaa-cliente.ps1.")
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error ejecutando wsaa-cliente.ps1: {e}")

    ok_files = sorted(run_dir.glob("*-loginTicketResponse.xml"), key=lambda p: p.stat().st_mtime, reverse=True)
    err_files = sorted(run_dir.glob("*-loginTicketResponse-ERROR.xml"), key=lambda p: p.stat().st_mtime, reverse=True)

    if not ok_files:
        if err_files:
            err_msg = err_files[0].read_text(encoding="utf-8", errors="ignore").strip()
        else:
            err_msg = ((completed.stdout or "") + "" + (completed.stderr or "")).strip()
            err_msg = err_msg or "No se generó loginTicketResponse.xml (sin detalle)."

        raise HTTPException(status_code=502, detail=f"WSAA ERROR: {err_msg}")

    ltr_path = ok_files[0]

    try:
        xml_text = ltr_path.read_text(encoding="utf-8", errors="ignore")
        root = ET.fromstring(xml_text)
    except Exception as e:
        raise HTTPException(status_code=502, detail=f"WSAA devolvió XML inválido: {e}")

    def _find_tag_text(tag_name: str):
        for el in root.iter():
            t = el.tag.split("}")[-1] if isinstance(el.tag, str) else el.tag
            if t == tag_name:
                return (el.text or "").strip()
        return None

    token = _find_tag_text("token")
    sign = _find_tag_text("sign")

    if not token or not sign:
        snippet = xml_text[:500].replace("", " ")
        raise HTTPException(status_code=502, detail=f"No se encontraron token/sign en loginTicketResponse.xml. Snippet: {snippet}")

    return token, sign

def call_padron_a5_get_persona(token: str, sign: str, cuit: str) -> dict:
    session = requests.Session()
    transport = Transport(session=session, timeout=30)
    client = Client(PADRON_A5_WSDL, transport=transport)

    resp = client.service.getPersona_v2(token, sign, int(cuit), int(cuit))

    return serialize_object(resp)

from typing import Any, Dict, List, Tuple, Optional

def _get(o: Any, key: str, default=None):
    if o is None:
        return default
    if isinstance(o, dict):
        return o.get(key, default)
    return getattr(o, key, default)

def _s(x: Any) -> str:
    if x is None:
        return ""
    return str(x).strip()

def _build_domicilio(domicilio_fiscal: Any) -> str:
    if not domicilio_fiscal:
        return ""
    direccion = _s(_get(domicilio_fiscal, "direccion"))
    localidad = _s(_get(domicilio_fiscal, "localidad"))
    prov = _s(_get(domicilio_fiscal, "descripcionProvincia"))
    cp = _s(_get(domicilio_fiscal, "codPostal"))

    parts = [p for p in [direccion, localidad, prov] if p]
    if cp:
        parts.append(f"CP {cp}")
    return " - ".join(parts)

def _detect_condicion_iva(padron: dict) -> str:

    dm = _get(padron, "datosMonotributo", {}) or {}
    dr = _get(padron, "datosRegimenGeneral", {}) or {}

    impuestos = []
    imp_dm = _get(dm, "impuesto", []) or []
    imp_dr = _get(dr, "impuesto", []) or []
    if isinstance(imp_dm, list):
        impuestos += imp_dm
    if isinstance(imp_dr, list):
        impuestos += imp_dr

    def imp_desc(i): return _s(_get(i, "descripcionImpuesto")).upper()
    def imp_estado(i): return _s(_get(i, "estadoImpuesto")).upper()

    if any(("MONOTRIBUTO" in imp_desc(i)) and (imp_estado(i).startswith("AC")) for i in impuestos):
        return "Monotributista"

    if any(("IVA" in imp_desc(i)) and (imp_estado(i).startswith("AC")) for i in impuestos):
        return "Responsable Inscripto"

    return ""  

def _extract_actividades(padron: dict) -> Tuple[List[dict], List[int]]:

    acts: List[dict] = []

    for origen_key in ("datosRegimenGeneral", "datosMonotributo"):
        sec = _get(padron, origen_key, {}) or {}
        arr = _get(sec, "actividad", []) or []
        if not isinstance(arr, list):
            continue

        for a in arr:
            id_act = _get(a, "idActividad", None)
            if id_act is None:
                continue
            try:
                id_int = int(id_act)
            except:
                continue

            acts.append({
                "id": id_int,
                "descripcion": _s(_get(a, "descripcionActividad")),
                "origen": origen_key,
                "orden": int(_get(a, "orden", 9999) or 9999),
                "periodo": _get(a, "periodo", None),
                "nomenclador": _get(a, "nomenclador", None),
            })

    by_id: Dict[int, dict] = {}
    for a in sorted(acts, key=lambda x: (x["orden"], x["id"])):
        if a["id"] not in by_id:
            by_id[a["id"]] = a
        else:
            if not by_id[a["id"]].get("descripcion") and a.get("descripcion"):
                by_id[a["id"]] = a

    out = list(by_id.values())
    out.sort(key=lambda x: (x["orden"], x["id"]))

    ids = [x["id"] for x in out]

    if 0 not in ids:
        out.insert(0, {
            "id": 0,
            "descripcion": "(Sin informar) 0",
            "origen": "system",
            "orden": 0,
            "periodo": None,
            "nomenclador": None,
        })
        ids.insert(0, 0)

    if not ids:
        out = [{
            "id": 0,
            "descripcion": "(Sin informar) 0",
            "origen": "system",
            "orden": 0,
            "periodo": None,
            "nomenclador": None,
        }]
        ids = [0]

    return out, ids

def map_padron_to_cache_fields(padron: dict) -> Tuple[dict, List[str]]:
    warnings: List[str] = []

    dg = _get(padron, "datosGenerales", {}) or {}

    tipo_persona = _s(_get(dg, "tipoPersona")).upper()
    razon = _s(_get(dg, "razonSocial"))
    apellido = _s(_get(dg, "apellido"))
    nombre = _s(_get(dg, "nombre"))

    razon_social = razon
    if not razon_social:
        if tipo_persona == "FISICA":
            razon_social = " ".join([p for p in [apellido, nombre] if p]).strip()
            if razon_social:
                warnings.append("ARCA no informó 'razonSocial'; se armó con apellido + nombre.")
        else:
            warnings.append("ARCA devolvió 'razonSocial' vacío/None para persona no física. Completar manualmente si aplica.")

    domicilio_fiscal = _get(dg, "domicilioFiscal", {}) or {}
    domicilio = _build_domicilio(domicilio_fiscal)

    if not domicilio:
        warnings.append("ARCA no devolvió domicilioFiscal (o vino incompleto).")

    condicion_iva = _detect_condicion_iva(padron)
    if not condicion_iva:
        warnings.append("No se pudo inferir condición IVA desde impuestos. Puede requerir otro endpoint o reglas adicionales.")

    actividades_detalle, actividades_ids = _extract_actividades(padron)

    mapped = {
        "razon_social": razon_social,
        "domicilio_comercial": domicilio,
        "condicion_iva": condicion_iva,
        "actividades_detalle": actividades_detalle,
        "actividades": actividades_ids,
    }
    return mapped, warnings


@router.post("/cuits/{cuit}/arca/refresh")
def refresh_arca(cuit: str):
    cuit = norm_cuit(cuit)
    db = load_cuits_db()

    cuit_item = None
    for it in db.get("cuits", []):
        if str(it.get("cuit")) == cuit:
            cuit_item = it
            break
    if not cuit_item:
        raise HTTPException(status_code=404, detail="CUIT no encontrado.")

    env = norm_env(cuit_item.get("environment", "demo"))

    cert_path, key_path = get_cert_key_paths_for_env(cuit, env)

    token, sign = obtener_token_sign_con_wsaa_cliente(
        env,
        "ws_sr_constancia_inscripcion",
        cert_path,
        key_path
    )

    try:
        padron = call_padron_a5_get_persona(token, sign, cuit)
    except Exception as e:
        raise HTTPException(status_code=502, detail=f"Error consultando Padrón/Constancia: {e}")

    cuit_item["padron_raw"] = str(padron)
    try:
        cuit_item["padron_json"] = json.dumps(padron, ensure_ascii=False, indent=2, default=str)
    except Exception:

        cuit_item["padron_json"] = None

    mapped, warnings = map_padron_to_cache_fields(padron)
    for k, v in mapped.items():
        cuit_item[k] = v

    if warnings:
        cuit_item.setdefault("arca_warnings", [])

        for w in warnings:
            if w not in cuit_item["arca_warnings"]:
                cuit_item["arca_warnings"].append(w)

    try:
        sel_act = int(cuit_item.get("selected_numero_actividad") or 0)
    except:
        sel_act = 0
    if sel_act not in (cuit_item.get("actividades") or [0]):
        cuit_item["selected_numero_actividad"] = 0

    cuit_item["last_refresh"] = datetime.now().isoformat(timespec="seconds")
    save_cuits_db(db)

    return {
        "ok": True,
        "cuit": int(cuit),
        "last_refresh": cuit_item["last_refresh"],
        "mapped": mapped,
        "warnings": cuit_item.get("arca_warnings", []),
    }

def get_base_dir() -> Path:
    p = Path(__file__).resolve()

    for parent in [p.parent, *p.parents]:
        if (parent / "Usuarios").exists() or (parent / "sources").exists() or (parent / "Configuraciones").exists():
            return parent

    return p.parents[2]

BASE_DIR = get_base_dir()
USUARIOS_DIR = BASE_DIR / "Usuarios"
SOURCES_DIR = BASE_DIR / "sources"
CONFIG_DIR = BASE_DIR / "Configuraciones"
CUITS_JSON = CONFIG_DIR / "cuits.json"
CONFIG_XLSX = CONFIG_DIR / "Config.xlsx"

AUTH_ABC_NEW = "Autorizacion_ABC_sin_item"
AUTH_AB_NEW  = "Autorizacion_AB_con_item"

SERVICES = {
    "FEV1": {"auth_dir": AUTH_ABC_NEW, "sources_subdir": "fev1"},
    "MTXCA": {"auth_dir": AUTH_AB_NEW,  "sources_subdir": "mtxca"},
}

ENV_FILES = {
    "demo": {
        "cert_name": "MiCertificado.cert",
        "key_names": ["MiClavePrivada.key"],
    },
    "produccion": {
        "cert_name": "MiCertificado.pem",
        "key_names": ["MiClavePrivada.key", "MiclavePrivada.key"],
    }
}

def norm_env(env: str) -> str:
    env = (env or "").strip().lower()
    if env in ("dev", "desarrollo", "demo"):
        return "demo"
    if env in ("prod", "produccion", "producción"):
        return "produccion"
    raise HTTPException(status_code=400, detail="environment inválido. Usá demo o produccion.")

def norm_cuit(cuit: str) -> str:
    digits = re.sub(r"\D", "", str(cuit or ""))
    if len(digits) != 11:
        raise HTTPException(status_code=400, detail="CUIT/CUIL inválido (debe tener 11 dígitos).")
    return digits

def validate_pem_cert(text: str):
    if "-----BEGIN CERTIFICATE-----" not in text:
        raise HTTPException(status_code=400, detail="Certificado inválido: falta '-----BEGIN CERTIFICATE-----'.")

def validate_private_key(text: str):
    if ("-----BEGIN PRIVATE KEY-----" not in text) and ("-----BEGIN RSA PRIVATE KEY-----" not in text):
        raise HTTPException(status_code=400, detail="Key inválida: falta '-----BEGIN PRIVATE KEY-----'.")

def ensure_dirs():
    USUARIOS_DIR.mkdir(parents=True, exist_ok=True)
    SOURCES_DIR.mkdir(parents=True, exist_ok=True)
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)

def load_cuits_db() -> dict:
    ensure_dirs()
    if CUITS_JSON.exists():
        try:
            return json.loads(CUITS_JSON.read_text(encoding="utf-8"))
        except Exception:
            return {"cuits": []}
    return {"cuits": []}

def save_cuits_db(db: dict):
    ensure_dirs()
    CUITS_JSON.write_text(json.dumps(db, ensure_ascii=False, indent=2), encoding="utf-8")

def upsert_cuit(db: dict, item: dict):
    cuits = db.get("cuits", [])
    found = False
    for i, it in enumerate(cuits):
        if str(it.get("cuit")) == str(item.get("cuit")):
            cuits[i] = {**it, **item}
            found = True
            break
    if not found:
        cuits.append(item)
    db["cuits"] = cuits

def copy_sources_to_user(cuit: str, env: str):
    for svc, meta in SERVICES.items():
        src = SOURCES_DIR / env / meta["sources_subdir"]
        if not src.exists():
            raise HTTPException(
                status_code=500,
                detail=f"Falta sources: {src}. (Revisá instalación/instalador global)"
            )
        dst = USUARIOS_DIR / cuit / meta["auth_dir"] / "source"
        dst.mkdir(parents=True, exist_ok=True)
        shutil.copytree(src, dst, dirs_exist_ok=True)

def write_cert_and_key_to_sources(cuit: str, env: str, cert_bytes: bytes, key_bytes: bytes):
    cert_name = ENV_FILES[env]["cert_name"]
    key_names = ENV_FILES[env]["key_names"]

    for svc, meta in SERVICES.items():
        dst_source = USUARIOS_DIR / cuit / meta["auth_dir"] / "source"
        dst_source.mkdir(parents=True, exist_ok=True)

        (dst_source / cert_name).write_bytes(cert_bytes)
        for kn in key_names:
            (dst_source / kn).write_bytes(key_bytes)

def ensure_token_txts(cuit: str):

    for meta in SERVICES.values():
        dst_source = USUARIOS_DIR / cuit / meta["auth_dir"] / "source"
        dst_source.mkdir(parents=True, exist_ok=True)
        (dst_source / "UltimoTokenWSFEV1.txt").write_text("", encoding="utf-8")
        (dst_source / "UltimoTokenWSMTXCA.txt").write_text("", encoding="utf-8")


@router.get("/cuits")
def list_cuits():
    db = load_cuits_db()
    cuits_out = []
    for it in db.get("cuits", []):
        cuits_out.append({
            "cuit": int(it["cuit"]),
            "razon_social": it.get("razon_social", "") or ""
        })
    return {"cuits": cuits_out}


@router.get("/cuits/{cuit}")
def get_cuit_detail(cuit: str):
    cuit = norm_cuit(cuit)
    db = load_cuits_db()
    for it in db.get("cuits", []):
        if str(it.get("cuit")) == cuit:
            out = dict(it)
            out.setdefault("selected_punto_venta", 0)
            out.setdefault("selected_numero_actividad", 0)
            return out
    raise HTTPException(status_code=404, detail="CUIT no encontrado.")


@router.post("/cuits")
async def add_cuit(
    cuit: str = Form(...),
    environment: str = Form(...),
    cert: UploadFile = File(...),
    key: UploadFile = File(...),
):
    cuit = norm_cuit(cuit)
    env = norm_env(environment)

    cert_bytes = await cert.read()
    key_bytes = await key.read()

    cert_text = cert_bytes.decode("utf-8", errors="ignore")
    key_text = key_bytes.decode("utf-8", errors="ignore")
    validate_pem_cert(cert_text)
    validate_private_key(key_text)

    ensure_dirs()

    copy_sources_to_user(cuit, env)

    write_cert_and_key_to_sources(cuit, env, cert_bytes, key_bytes)

    ensure_token_txts(cuit)

    db = load_cuits_db()
    upsert_cuit(db, {
        "cuit": cuit,
        "environment": env,
        "razon_social": "",
        "domicilio_comercial": "",
        "condicion_iva": "",
        "puntos_venta": [],
        "actividades": [],
        "selected_punto_venta": 0,
        "selected_numero_actividad": 0,
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "last_refresh": None,
    })
    save_cuits_db(db)

    return {
        "ok": True,
        "cuit": int(cuit),
        "environment": env,
        "files_written": {
            "cert": ENV_FILES[env]["cert_name"],
            "keys": ENV_FILES[env]["key_names"],
        }
    }


class SyncExcelRequest(BaseModel):
    punto_venta: int
    numero_actividad: int = 0
@router.post("/cuits/{cuit}/sync-excel")
def sync_excel_from_cuit_cache(cuit: str, req: SyncExcelRequest):

    cuit_norm = norm_cuit(cuit)

    db = load_cuits_db()
    cuit_item = None
    for it in db.get("cuits", []):
        if str(it.get("cuit")) == cuit_norm:
            cuit_item = it
            break
    if not cuit_item:
        raise HTTPException(status_code=404, detail="CUIT no encontrado en cuits.json")

    if not CONFIG_XLSX.exists():
        raise HTTPException(status_code=500, detail=f"No existe Config.xlsx en: {CONFIG_XLSX}")

    pv = int(req.punto_venta) if req.punto_venta is not None else 0
    na = int(req.numero_actividad) if req.numero_actividad is not None else 0

    cuit_item["selected_punto_venta"] = pv
    cuit_item["selected_numero_actividad"] = na
    cuit_item["last_sync_excel"] = datetime.now().isoformat(timespec="seconds")
    save_cuits_db(db)

    try:
        result = sync_row2_all_tipos(
            str(CONFIG_XLSX),
            cuit=int(cuit_norm),
            razon_social=str(cuit_item.get("razon_social") or ""),
            domicilio_comercial=str(cuit_item.get("domicilio_comercial") or ""),
            condicion_iva=str(cuit_item.get("condicion_iva") or ""),
            punto_venta=pv,
            numero_actividad=na,
        )
    except PermissionError:
        raise HTTPException(status_code=409, detail="No se pudo guardar Config.xlsx (¿está abierto en Excel?). Cerralo y reintentá.")
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error actualizando Config.xlsx: {e}")

    return {"ok": True, "cuit": int(cuit_norm), "result": result}