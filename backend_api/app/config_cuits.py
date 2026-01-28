from fastapi import APIRouter, UploadFile, File, Form, HTTPException
from pathlib import Path
import re
import shutil
import json
from datetime import datetime

router = APIRouter(prefix="/api/config", tags=["config"])

import subprocess
from xml.etree import ElementTree as ET
from zeep import Client
from zeep.transports import Transport
import requests

# --- WSDL de Padrón A5 (personaServiceA5) ---
PADRON_A5_WSDL = "https://aws.afip.gov.ar/sr-padron/webservices/personaServiceA5?wsdl"  # :contentReference[oaicite:4]{index=4}

def get_cert_key_paths_for_env(cuit: str, env: str) -> tuple[Path, Path]:
    """
    Devuelve (cert_path, key_path) según env, apuntando a la carpeta source del usuario.
    Tomamos FEV1 source como "canónica" (igual escribimos en ambos).
    """
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
    """
    ESTA FUNCIÓN LA CONECTAMOS A TU wsaa-cliente.

    Como vos ya tenés wsaa-cliente dentro de sources/<env>/<servicio>/,
    lo ideal es invocar el script que genere el loginTicketResponse y parsearlo.

    - servicio para padrón/constancia: normalmente 'ws_sr_padron_a5' o 'ws_sr_ws_constancia_inscripcion'
      (según cuál uses/habilites).
    - Devuelve (token, sign).
    """

    raise HTTPException(status_code=501, detail="Falta conectar wsaa-cliente para obtener token/sign (WSAA).")


def call_padron_a5_get_persona(token: str, sign: str, cuit: str) -> dict:
    """
    Llama a getPersona_v2 (Padrón A5 / Constancia).
    Estructura de request token/sign/cuitRepresentada/idPersona está documentada. :contentReference[oaicite:5]{index=5}
    """
    session = requests.Session()
    transport = Transport(session=session, timeout=30)
    client = Client(PADRON_A5_WSDL, transport=transport)

    resp = client.service.getPersona_v2(token, sign, int(cuit), int(cuit))

    return resp


@router.post("/cuits/{cuit}/arca/refresh")
def refresh_arca(cuit: str):
    cuit = norm_cuit(cuit)
    db = load_cuits_db()

    # buscar env en db
    cuit_item = None
    for it in db.get("cuits", []):
        if str(it.get("cuit")) == cuit:
            cuit_item = it
            break
    if not cuit_item:
        raise HTTPException(status_code=404, detail="CUIT no encontrado.")

    env = norm_env(cuit_item.get("environment", "demo"))

    # 1) ubicar cert/key
    cert_path, key_path = get_cert_key_paths_for_env(cuit, env)

    # 2) obtener token/sign WSAA para Padrón/Constancia
    #    (acá lo conectamos a tu wsaa-cliente)
    token, sign = obtener_token_sign_con_wsaa_cliente(env, "padron", cert_path, key_path)

    # 3) llamar padrón
    try:
        padron = call_padron_a5_get_persona(token, sign, cuit)
    except Exception as e:
        raise HTTPException(status_code=502, detail=f"Error consultando Padrón/Constancia: {e}")

    cuit_item["padron_raw"] = str(padron)
    cuit_item["last_refresh"] = datetime.now().isoformat(timespec="seconds")

    save_cuits_db(db)
    return {"ok": True, "cuit": int(cuit), "last_refresh": cuit_item["last_refresh"]}

# =========================
# Paths base (portable)
# =========================
def get_base_dir() -> Path:
    """
    Devuelve la carpeta base donde viven:
    - Usuarios/
    - sources/
    - Configuraciones/
    base_dir = parents[2]
    """
    p = Path(__file__).resolve()

    # Búsqueda robusta hacia arriba
    for parent in [p.parent, *p.parents]:
        if (parent / "Usuarios").exists() or (parent / "sources").exists() or (parent / "Configuraciones").exists():
            return parent

    # Fallback para tu estructura actual (app -> backend_api -> carpeta_con_todo)
    return p.parents[2]

BASE_DIR = get_base_dir()
USUARIOS_DIR = BASE_DIR / "Usuarios"
SOURCES_DIR = BASE_DIR / "sources"
CONFIG_DIR = BASE_DIR / "Configuraciones"
CUITS_JSON = CONFIG_DIR / "cuits.json"

# Carpetas nuevas (sin acentos)
AUTH_ABC_NEW = "Autorizacion_ABC_sin_item"
AUTH_AB_NEW  = "Autorizacion_AB_con_item"

# Servicios y rutas sources esperadas
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
    """
    Copia sources/<env>/<fev1|mtxca> a Usuarios/<cuit>/<auth_dir>/source
    """
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
    """
    Guarda cert + key en AMBOS source (FEV1 y MTXCA) con nombres obligatorios según env.
    """
    cert_name = ENV_FILES[env]["cert_name"]
    key_names = ENV_FILES[env]["key_names"]

    for svc, meta in SERVICES.items():
        dst_source = USUARIOS_DIR / cuit / meta["auth_dir"] / "source"
        dst_source.mkdir(parents=True, exist_ok=True)

        # cert
        (dst_source / cert_name).write_bytes(cert_bytes)

        # key (en prod: guardamos 2 nombres por compat de casing)
        for kn in key_names:
            (dst_source / kn).write_bytes(key_bytes)

def ensure_token_txts(cuit: str):
    """
    Crea los TXT de 'ultimo token' para evitar faltantes.
    (Ajustalo si tus scripts los guardan en otro lugar.)
    """
    # Los creo en ambos sources por simplicidad/robustez
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
            return it
    raise HTTPException(status_code=404, detail="CUIT no encontrado.")


@router.post("/cuits")
async def add_cuit(
    cuit: str = Form(...),
    environment: str = Form("prod"), 
    cert: UploadFile = File(...),
    key: UploadFile = File(...),
):
    cuit = norm_cuit(cuit)
    env = norm_env(environment) 
    cert_bytes = await cert.read()
    key_bytes = await key.read()

    # Validación "funcional mínima" (lo que pediste)
    cert_text = cert_bytes.decode("utf-8", errors="ignore")
    key_text = key_bytes.decode("utf-8", errors="ignore")
    validate_pem_cert(cert_text)
    validate_private_key(key_text)

    ensure_dirs()

    # 1) Copiar WSAA / sources a la carpeta del usuario
    copy_sources_to_user(cuit, env)

    # 2) Escribir cert y key con nombres obligatorios según env
    write_cert_and_key_to_sources(cuit, env, cert_bytes, key_bytes)

    # 3) Asegurar txts token (robusto)
    ensure_token_txts(cuit)

    # 4) Persistir en "DB" (json)
    db = load_cuits_db()
    upsert_cuit(db, {
        "cuit": cuit,
        "environment": env,
        "razon_social": "",
        "domicilio_comercial": "",
        "condicion_iva": "",
        "puntos_venta": [],
        "actividades": [],
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
