from fastapi import APIRouter, UploadFile, File, Form, HTTPException
from pathlib import Path
import re
import shutil
import json
import os
import sys
from datetime import datetime
from pydantic import BaseModel
from .excel_config import sync_row2_all_tipos, get_config_by_cuit
from zeep.helpers import serialize_object

router = APIRouter(prefix="/api/config", tags=["config"])

import subprocess
from xml.etree import ElementTree as ET
from zeep import Client
from zeep.transports import Transport
import requests

import ssl
from requests.adapters import HTTPAdapter
from urllib3.poolmanager import PoolManager
import glob
from datetime import datetime, timedelta
from pathlib import Path

CONSTANCIA_WSDL = "https://aws.afip.gov.ar/sr-padron/webservices/personaServiceA5?wsdl"
CONSTANCIA_SERVICE_ID = "ws_sr_constancia_inscripcion"

WSFE_WSDL = {
    "demo": "https://wswhomo.afip.gov.ar/wsfev1/service.asmx?WSDL",
    "produccion": "https://servicios1.afip.gov.ar/wsfev1/service.asmx?WSDL",
}

WSAA_WSDL_PROD = "https://wsaa.afip.gov.ar/ws/services/LoginCms?WSDL"
WSAA_WSDL_HOMO = "https://wsaahomo.afip.gov.ar/ws/services/LoginCms?WSDL"

class AFIPLegacyTLSAdapter(HTTPAdapter):
    def __init__(self, ssl_context: ssl.SSLContext, **kwargs):
        self._ssl_context = ssl_context
        super().__init__(**kwargs)

    def init_poolmanager(self, connections, maxsize, block=False, **pool_kwargs):
        pool_kwargs["ssl_context"] = self._ssl_context
        self.poolmanager = PoolManager(
            num_pools=connections,
            maxsize=maxsize,
            block=block,
            **pool_kwargs
        )

    def proxy_manager_for(self, proxy, **proxy_kwargs):
        proxy_kwargs["ssl_context"] = self._ssl_context
        return super().proxy_manager_for(proxy, **proxy_kwargs)

def make_afip_legacy_session() -> requests.Session:

    ctx = ssl.create_default_context()

    ctx.set_ciphers("DEFAULT:@SECLEVEL=1")

    s = requests.Session()
    s.mount("https://servicios1.afip.gov.ar/", AFIPLegacyTLSAdapter(ctx))
    s.mount("https://wswhomo.afip.gov.ar/", AFIPLegacyTLSAdapter(ctx))
    return s


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


def parse_login_ticket_response(xml_path: str) -> tuple[str, str]:

    tree = ET.parse(xml_path)
    root = tree.getroot()

    token_el = root.find(".//token")
    sign_el = root.find(".//sign")

    token = token_el.text.strip() if token_el is not None and token_el.text else ""
    sign = sign_el.text.strip() if sign_el is not None and sign_el.text else ""

    if not token or not sign:
        raise RuntimeError(f"No se encontró token/sign en: {xml_path}")

    return token, sign


def _wsaa_wsdl_for_env(env: str) -> str:
    e = (env or "").strip().lower()
    # ajustá si tus env se llaman distinto
    if e in ("demo", "testing", "test", "homo", "homologacion", "homologación"):
        return WSAA_WSDL_HOMO
    return WSAA_WSDL_PROD

def _wsaa_stamp_filename(service_id) -> str:
    # a prueba de errores por si llega Path
    sid = service_id
    if isinstance(sid, Path):
        sid = sid.name
    sid = str(sid or "").strip().lower()

    if sid == "wsfe":
        return "UltimoTokenWSFEV1.txt"
    if sid == "wsmtxca":
        return "UltimoTokenWSMTXCA.txt"
    return "UltimoTokenPADRON.txt"

def _read_stamp_if_fresh(stamp_path: Path, timeout_min: int) -> bool:
    if not stamp_path.exists():
        return False
    try:
        raw = stamp_path.read_text(encoding="utf-8").strip()
        try:
            ts = datetime.fromisoformat(raw.replace("Z", ""))
        except Exception:
            ts = datetime.strptime(raw, "%Y-%m-%d %H:%M:%S")
        return datetime.now() - ts <= timedelta(minutes=timeout_min)
    except Exception:
        return False

def _write_stamp(stamp_path: Path) -> None:
    stamp_path.write_text(datetime.now().strftime("%Y-%m-%d %H:%M:%S"), encoding="utf-8")

def _find_latest_login_ticket_response(run_dir: Path) -> Path | None:
    candidates = []
    candidates += list(run_dir.glob("*-loginTicketResponse.xml"))
    candidates += list(run_dir.glob("loginTicketResponse.xml"))
    if not candidates:
        return None
    return max(candidates, key=lambda p: p.stat().st_mtime)

def _parse_login_ticket_response(xml_path: Path) -> tuple[str, str]:
    txt = xml_path.read_text(encoding="utf-8", errors="ignore")
    root = ET.fromstring(txt)
    token = (root.findtext(".//credentials/token") or "").strip()
    sign = (root.findtext(".//credentials/sign") or "").strip()
    if not token or not sign:
        raise RuntimeError(f"No se pudo extraer token/sign desde {xml_path.name}")
    return token, sign

def _cleanup_wsaa_artifacts(run_dir: Path) -> None:
    patterns = [
        "*-LoginTicketRequest.xml",
        "*-LoginTicketRequest.xml.cms*",
        "*-loginTicketResponse.xml",
        "*-loginTicketResponse-ERROR.xml",
        "*-LoginTicketRequest.xml.cms-DER*",
    ]
    for pat in patterns:
        for f in run_dir.glob(pat):
            try:
                f.unlink()
            except Exception:
                pass

def obtener_token_sign_con_wsaa_cliente(
    env: str,
    service_id: str,
    cert_path: Path,
    key_path: Path,
    timeout_min: int = 10
) -> tuple[str, str]:

    wsaa_wsdl = _wsaa_wsdl_for_env(env)

    cert_path = Path(cert_path)
    key_path = Path(key_path)
    run_dir = cert_path.parent

    stamp_path = run_dir / _wsaa_stamp_filename(service_id)

    # Reusar token si está fresco
    if _read_stamp_if_fresh(stamp_path, timeout_min):
        latest = _find_latest_login_ticket_response(run_dir)
        if latest:
            return _parse_login_ticket_response(latest)

    # Renovar
    _cleanup_wsaa_artifacts(run_dir)

    ps1_path = run_dir / "wsaa-cliente.ps1"
    if not ps1_path.exists():
        raise FileNotFoundError(f"No existe wsaa-cliente.ps1 en: {run_dir}")

    cmd = [
        "powershell",
        "-ExecutionPolicy", "Bypass",
        "-File", str(ps1_path),
        "-Certificado", str(cert_path),
        "-ClavePrivada", str(key_path),
        "-ServicioId", str(service_id),
        "-WsaaWsdl", str(wsaa_wsdl),
    ]

    result = subprocess.run(
        cmd,
        cwd=str(run_dir),
        capture_output=True,
        text=True
    )

    if result.returncode != 0:
        raise RuntimeError(
            "WSAA falló.\n"
            f"STDOUT:\n{result.stdout}\n\n"
            f"STDERR:\n{result.stderr}"
        )

    latest = _find_latest_login_ticket_response(run_dir)
    if not latest:
        raise RuntimeError(
            f"WSAA OK pero no apareció loginTicketResponse.xml en {run_dir}\n"
            f"STDOUT:\n{result.stdout}\n\n"
            f"STDERR:\n{result.stderr}"
        )

    token, sign = _parse_login_ticket_response(latest)
    _write_stamp(stamp_path)
    return token, sign

def call_constancia_get_persona(token: str, sign: str, cuit: str) -> dict:
    session = requests.Session()
    transport = Transport(session=session, timeout=30)
    client = Client(CONSTANCIA_WSDL, transport=transport)

    resp = client.service.getPersona_v2(token, sign, int(cuit), int(cuit))
    return serialize_object(resp)

def call_wsfe_get_ptos_venta(env: str, token: str, sign: str, cuit: str) -> tuple[list[int], dict]:

    env = norm_env(env)
    wsdl = WSFE_WSDL.get(env)
    if not wsdl:
        raise ValueError(f"No hay WSDL configurado para env={env}")

    session = make_afip_legacy_session() if env in ("produccion", "demo") else requests.Session()
    transport = Transport(session=session, timeout=30)
    client = Client(wsdl, transport=transport)

    auth = {"Token": token, "Sign": sign, "Cuit": int(cuit)}

    try:
        resp = client.service.FEParamGetPtosVenta(Auth=auth)
    except TypeError:
        resp = client.service.FEParamGetPtosVenta(auth)

    raw = serialize_object(resp)

    node = raw
    if isinstance(node, dict) and "FEParamGetPtosVentaResult" in node:
        node = node.get("FEParamGetPtosVentaResult")
    if isinstance(node, dict) and "ResultGet" in node:
        node = node.get("ResultGet")

    def _deep_find(obj, keys_lower):
        if obj is None:
            return None
        if isinstance(obj, dict):
            for k, v in obj.items():
                if str(k).lower() in keys_lower:
                    return v
                got = _deep_find(v, keys_lower)
                if got is not None:
                    return got
        elif isinstance(obj, list):
            for it in obj:
                got = _deep_find(it, keys_lower)
                if got is not None:
                    return got
        return None

    pto_entries = None
    if isinstance(node, dict):
        pto_entries = node.get("PtoVenta") or node.get("PtoVta") or node.get("PuntosVenta")
        if isinstance(pto_entries, dict):
            pto_entries = pto_entries.get("PtoVenta") or pto_entries.get("item") or pto_entries

    if pto_entries is None:
        pto_entries = _deep_find(node, {"ptoventa", "ptovta", "puntosventa", "puntos_venta"})

    ptos: list[int] = []

    def _add_nro(d: dict):
        nro = d.get("Nro") or d.get("nro") or d.get("PtoVta") or d.get("ptoVta") or d.get("numero") or d.get("Numero")
        try:
            n = int(nro)
        except Exception:
            return
        ptos.append(n)

    if isinstance(pto_entries, list):
        for it in pto_entries:
            if isinstance(it, dict):
                _add_nro(it)
            else:
                try:
                    _add_nro(serialize_object(it) or {})
                except Exception:
                    continue
    elif isinstance(pto_entries, dict):
        _add_nro(pto_entries)

    ptos_sorted = sorted(set(ptos))
    return ptos_sorted, raw


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

    # --- 1) WSFE: puntos de venta (ya lo tenías con try/except) ---
    try:
        token_fe, sign_fe = obtener_token_sign_con_wsaa_cliente(env, "wsfe", cert_path, key_path)
        pv_list, pv_raw = call_wsfe_get_ptos_venta(env, token_fe, sign_fe, cuit)

        cuit_item["puntos_venta"] = pv_list
        cuit_item["wsfe_ptos_venta_raw"] = pv_raw
        cuit_item.pop("wsfe_ptos_venta_error", None)
    except Exception as e:
        cuit_item["wsfe_ptos_venta_error"] = str(e)

    # --- 2) CONSTANCIA: token/sign + consulta ---
    constancia_err = None
    try:
        token_c, sign_c = obtener_token_sign_con_wsaa_cliente(
            env,
            CONSTANCIA_SERVICE_ID,   # <- estandarizado
            cert_path,
            key_path
        )
    except Exception as e:
        constancia_err = f"WSAA/Constancia: {e}"
        cuit_item["constancia_error"] = constancia_err

        cuit_item["last_refresh"] = datetime.now().isoformat(timespec="seconds")
        save_cuits_db(db)

        # devolvemos 200 pero ok=False (para que el front muestre toast y siga)
        return {
            "ok": False,
            "cuit": int(cuit),
            "last_refresh": cuit_item["last_refresh"],
            "errors": {
                "constancia": constancia_err,
                "wsfe": cuit_item.get("wsfe_ptos_venta_error"),
            },
            "warnings": cuit_item.get("arca_warnings", []),
        }

    # Si llegamos acá, hay token/sign de constancia
    try:
        padron = call_constancia_get_persona(token_c, sign_c, cuit)
        cuit_item.pop("constancia_error", None)
    except Exception as e:
        constancia_err = f"Error consultando Constancia: {e}"
        cuit_item["constancia_error"] = constancia_err

        cuit_item["last_refresh"] = datetime.now().isoformat(timespec="seconds")
        save_cuits_db(db)

        return {
            "ok": False,
            "cuit": int(cuit),
            "last_refresh": cuit_item["last_refresh"],
            "errors": {
                "constancia": constancia_err,
                "wsfe": cuit_item.get("wsfe_ptos_venta_error"),
            },
            "warnings": cuit_item.get("arca_warnings", []),
        }

    # --- 3) Guardar raw/json + mapear a cache fields (tu lógica original) ---
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
        "errors": {
            "wsfe": cuit_item.get("wsfe_ptos_venta_error"),
            "constancia": cuit_item.get("constancia_error"),
        }
    }

def get_base_dir() -> Path:

    env_root = os.getenv("INVOICER_PROJECT_ROOT") or os.getenv("INVOICER_BASE_DIR")
    if env_root:
        cand = Path(env_root).resolve()
        if cand.exists():
            return cand


    if getattr(sys, "frozen", False):
        cand = Path(sys.executable).resolve().parent
        if (cand / "Usuarios").exists() or (cand / "sources").exists() or (cand / "Configuraciones").exists() or (cand / "configuraciones").exists():
            return cand

    p = Path(__file__).resolve()
    for parent in [p.parent, *p.parents]:
        if (parent / "Usuarios").exists() or (parent / "sources").exists() or (parent / "Configuraciones").exists() or (parent / "configuraciones").exists():
            if parent.name.lower() == "_internal":
                continue
            return parent

    return p.parents[2]

BASE_DIR = get_base_dir()
USUARIOS_DIR = BASE_DIR / "Usuarios"
SOURCES_DIR = BASE_DIR / "sources"
CONFIG_DIR = BASE_DIR / ("Configuraciones" if (BASE_DIR / "Configuraciones").exists() else "configuraciones")
CUITS_JSON = CONFIG_DIR / "cuits.json"
CONFIG_XLSX = CONFIG_DIR / "Config.xlsx"

print(f"[config_cuits] BASE_DIR={BASE_DIR} | CONFIG_DIR={CONFIG_DIR} | CUITS_JSON={CUITS_JSON} | exists={CUITS_JSON.exists()} | CWD={Path.cwd()}")

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

def _extract_puntos_venta_from_raw(raw: dict) -> list[int]:
    try:
        pv = raw.get("ResultGet", {}).get("PtoVenta", [])
        out: list[int] = []
        if isinstance(pv, list):
            for it in pv:
                if isinstance(it, dict) and "Nro" in it:
                    try:
                        out.append(int(it["Nro"]))
                    except Exception:
                        pass

        return sorted(set(out))
    except Exception:
        return []

def _ensure_puntos_venta_fields(it: dict) -> None:

    pv = it.get("puntos_venta")
    if not (isinstance(pv, list) and len(pv) > 0):
        raw = it.get("wsfe_ptos_venta_raw")
        if isinstance(raw, dict):
            pv2 = _extract_puntos_venta_from_raw(raw)
            if pv2:
                it["puntos_venta"] = pv2

    if it.get("selected_punto_venta") is None:
        pv = it.get("puntos_venta")
        if isinstance(pv, list) and pv:
            it["selected_punto_venta"] = pv[0]

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

    try:
        raw = CUITS_JSON.read_text(encoding="utf-8-sig")
        raw = raw.strip()
        if not raw:
            return {"cuits": []}

        db = json.loads(raw)

        # Normalizar estructura
        if not isinstance(db, dict):
            return {"cuits": []}
        if "cuits" not in db or not isinstance(db.get("cuits"), list):
            db["cuits"] = []
        return db

    except Exception as e:
        print(f"[config_cuits] ERROR leyendo {CUITS_JSON}: {e}")
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
            _ensure_puntos_venta_fields(it)
            result = dict(it)
            if CONFIG_XLSX.exists():
                for tipo in ("A", "B", "C"):
                    try:
                        excel_data = get_config_by_cuit(str(CONFIG_XLSX), tipo, int(cuit))
                        result["ingresos_brutos"] = excel_data.get("ingresos_brutos", "")
                        result["fecha_inicio_actividades"] = excel_data.get("fecha_inicio_actividades", "")
                        break
                    except KeyError:
                        continue
                    except Exception:
                        break
            return result

    raise HTTPException(status_code=404, detail="CUIT no encontrado")


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
        "ingresos_brutos": "",
        "fecha_inicio_actividades": "",
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
    ingresos_brutos: Optional[str] = None
    inicio_actividades: Optional[str] = None
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
            ingresos_brutos=req.ingresos_brutos,
            fecha_inicio_actividades=req.inicio_actividades,
            punto_venta=pv,
            numero_actividad=na,
        )
    except PermissionError:
        raise HTTPException(status_code=409, detail="No se pudo guardar Config.xlsx (¿está abierto en Excel?). Cerralo y reintentá.")
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error actualizando Config.xlsx: {e}")

    return {"ok": True, "cuit": int(cuit_norm), "result": result}