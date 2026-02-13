import os
import sys
from pathlib import Path

def _is_frozen() -> bool:
    return bool(getattr(sys, "frozen", False))

RUNTIME_DIR = Path(sys.executable).resolve().parent if _is_frozen() else Path(__file__).resolve().parent

RESOURCE_DIR = Path(getattr(sys, "_MEIPASS", str(RUNTIME_DIR))) if _is_frozen() else RUNTIME_DIR

BASE_DIR = str(RUNTIME_DIR)

def resource_path(*parts: str) -> str:

    p1 = RESOURCE_DIR.joinpath(*parts)
    if p1.exists():
        return str(p1)
    return str(RUNTIME_DIR.joinpath(*parts))

#* VARIABLE GENERAL DE CONFIGURACIÓN
CONFIGURACIONES_DIR = os.path.join(BASE_DIR, 'Configuraciones')

#* REVISIÓN PARA CORRER ARCHIVOS .PS1 (RUTA DONDE ESTAN LOS TX PARA REVISAR ESO)
ULTIMO_TOKEN_WSFEV1_PATH = os.path.join(CONFIGURACIONES_DIR, 'UltimoTokenWSFEV1.txt')
ULTIMO_TOKEN_WSMTXCA_PATH = os.path.join(CONFIGURACIONES_DIR, 'UltimoTokenWSMTXCA.txt')

import os

USUARIOS_DIR = os.path.join(BASE_DIR, "Usuarios")

_AUTORIZA_DIR_POR_SERVICIO = {
    "wsfev1": "Autorizacion_ABC_sin_item",
    "wsmtxca": "Autorizacion_AB_con_item",
}

def autorizacion_source_dir(cuit: int | str, servicio: str) -> str:
    cuit_str = str(cuit).strip()
    key = (servicio or "").lower().strip()

    if key not in _AUTORIZA_DIR_POR_SERVICIO:
        raise ValueError(f"Servicio inválido: {servicio}. Use 'wsfev1' o 'wsmtxca'.")

    return os.path.join(USUARIOS_DIR, cuit_str, _AUTORIZA_DIR_POR_SERVICIO[key], "source")

def login_ticket_response_wsfev1_path(cuit: int | str) -> str:
    return autorizacion_source_dir(cuit, "wsfev1")

def login_ticket_response_wsmtxca_path(cuit: int | str) -> str:
    return autorizacion_source_dir(cuit, "wsmtxca")

#* CARPETAS DEL PROCESO GENERALES
CONFIG_PATH = os.path.join(CONFIGURACIONES_DIR, 'Config.txt')
INPUT_DIR = os.path.join(BASE_DIR, 'input')
FACTURAS_DIR = os.path.join(BASE_DIR, 'facturas')
FACTURA_A_DIR = os.path.join(INPUT_DIR, 'Factura A')
FACTURA_B_DIR = os.path.join(INPUT_DIR, 'Factura B')
FACTURA_C_DIR = os.path.join(INPUT_DIR, 'Factura C')
OUTPUT_DIR = os.path.join(BASE_DIR, 'output')
LOG_DIR = os.path.join(BASE_DIR, 'log')

# Variables para almacenar datos
# Estas variables son inicializadas al principio para utilizar a lo largo del proceso
ListaValidacionCAE = []