import os
import sys

#* ----------------------------------------- [DEFINIR RUTA BASE DEL PROYECTO] ----------------------------------------- *# 
# tiene un IF para que pueda ser corrido tanto por visual studio (script de python)
# como para que pueda ser ejecutado por un .exe en cualquier lugar de la computadora.
BASE_DIR = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.dirname(os.path.abspath(__file__))

#* VARIABLE GENERAL DE CONFIGURACIÓN
CONFIGURACIONES_DIR = os.path.join(BASE_DIR, 'Configuraciones')

#* REVISIÓN PARA CORRER ARCHIVOS .PS1 (RUTA DONDE ESTAN LOS TX PARA REVISAR ESO)
ULTIMO_TOKEN_WSFEV1_PATH = os.path.join(CONFIGURACIONES_DIR, 'UltimoTokenWSFEV1.txt')
ULTIMO_TOKEN_WSMTXCA_PATH = os.path.join(CONFIGURACIONES_DIR, 'UltimoTokenWSMTXCA.txt')

#* DONDE ESTAN LOS .PS1 PARA SER CORRIDOS
LOGIN_TICKET_RESPONSE_WSFEV1_PATH = os.path.join(BASE_DIR, 'Autorización A, B y C sin item', 'source')
LOGIN_TICKET_RESPONSE_WSMTXCA_PATH = os.path.join(BASE_DIR, 'Autorización A y B con item', 'source')

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