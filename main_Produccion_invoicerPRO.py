#
#                        !    INVOICER PRO    !    
#* [               PROCESO PARA FACTURACIÓN ELECTRONICA                  ]
#* Realizado por FysTech Group
#* Desarrollador: Matias Nicolas Franco
#* Versión: 4.0
#* Codigo Modularizado
#* Ultima Update: 28/01/2026 En Proceso
#* ----------------------------------------- [IMPORTS] ----------------------------------------- *# 

# * IMPORTS MODULOS
#! < BACKEND >
from backend.log import escribir_log, obtener_timestamp, cerrar_log, abrir_log
from backend.config import leer_datos_excel
#! < TOKEN >
from autorizar.obtener_datos import verificar_token, ejecutar_powershell, leer_ticket_response, limpiar_output_dir
#! < SERVICIO: FEV1 - Facturas A B o C sin detalle de items >
from fev1.procesar_datos import procesar_datos_fev1
from fev1.armado_cuerpo import armar_cuerpo_solicitud_fev1
from fev1.nota_credito import procesar_datos_nota_credito_fev1, armar_cuerpo_nota_credito_fev1
from fev1.nota_debito import procesar_datos_nota_debito_fev1, armar_cuerpo_nota_debito_fev1
#! < SERVICIO: MTXCA - Facturas A y B con detalle de items >
from mtxca.procesar_datos import procesar_datos_mtxca
from mtxca.armado_cuerpo import armar_cuerpo_solicitud_mtxca_A, armar_cuerpo_solicitud_mtxca_B
from mtxca.nota_credito import procesar_datos_nota_credito_mtxca, armar_cuerpo_nota_credito_mtxca
from mtxca.nota_debito import procesar_datos_nota_debito_mtxca, armar_cuerpo_nota_debito_mtxca
#! < SOLICITUDES AFIP >
from solicitudesAFIP.solicitar_cae import enviar_solicitud_cae
#! < FACTURAS >
from facturasOuput.armado_pdf import completar_plantilla, convertir_xlsx_a_pdf
#! < VARIABLES INICIALES >
from variables import FACTURA_A_DIR, FACTURA_B_DIR, FACTURA_C_DIR, FACTURAS_DIR
from variables import CONFIGURACIONES_DIR, ListaValidacionCAE
from variables import OUTPUT_DIR, BASE_DIR
from variables import os
#! < ARMADO DE FACTURAS >
from zeep import Client
import shutil
import json
import pandas as pd
from datetime import datetime
import argparse

import ssl
import requests
from requests.adapters import HTTPAdapter
from urllib3.poolmanager import PoolManager
from zeep.transports import Transport


# * EXCEL Y DISPLAY DE LAS COLUMNAS:
from backend.excel_armador_input import (
    columns_factura_a, columns_factura_a_credito, columns_factura_a_debito,
    columns_factura_b, columns_factura_b_credito, columns_factura_b_debito,
    columns_factura_c, columns_factura_c_credito, columns_factura_c_debito
)
from backend.excel_armador_input import save_to_excel, select_existing_file

# Definición de las plantillas
templates = {
    "Factura A": "vacio_factura_a.xlsx",
    "Nota Credito A": "vacio_factura_a.xlsx",
    "Nota Debito A": "vacio_factura_a.xlsx",
    "Factura B": "vacio_factura_b.xlsx",
    "Nota Credito B": "vacio_factura_b.xlsx",
    "Nota Debito B": "vacio_factura_b.xlsx",
    "Factura C": "vacio_factura_c.xlsx",
    "Nota Credito C": "vacio_factura_c.xlsx",
    "Nota Debito C": "vacio_factura_c.xlsx",
}

#! ------------------------------------------------------------------------------------------------------------------------------------------------------
#!
#! CONFIGURACIONES INICIALES PLANILLA EXCEL
#!
#! ------------------------------------------------------------------------------------------------------------------------------------------------------

def leer_configuraciones_desde_excel(ruta_excel):
    try:
        hoja_a = pd.read_excel(ruta_excel, sheet_name='FacturaA').fillna("")
        hoja_b = pd.read_excel(ruta_excel, sheet_name='FacturaB').fillna("")
        hoja_c = pd.read_excel(ruta_excel, sheet_name='FacturaC').fillna("")

        # Devolvemos LISTA de configs (una por fila)
        return {
            "Factura A": hoja_a.to_dict(orient="records"),
            "Factura B": hoja_b.to_dict(orient="records"),
            "Factura C": hoja_c.to_dict(orient="records"),
        }
    except Exception as e:
        escribir_log(f"{obtener_timestamp()} - Error al leer el archivo de configuración desde Excel: {e}")
        cerrar_log()
        return None

def _to_int(val):
    # Normaliza CUIT/PtoVta aunque venga como float (ej: 3071...641.0)
    try:
        if val is None:
            return None
        if isinstance(val, str):
            val = val.strip()
            if val == "":
                return None
        return int(float(val))
    except Exception:
        return None

def obtener_configuracion_por_cuit(configuraciones, tipo_factura, cuit_seleccionado, punto_venta=None):
    key = {"A": "Factura A", "B": "Factura B", "C": "Factura C"}.get(tipo_factura)
    if not key:
        return None

    cuit_sel = _to_int(cuit_seleccionado)
    if cuit_sel is None:
        return None

    filas = configuraciones.get(key, [])
    candidatos = []

    for fila in filas:
        if _to_int(fila.get("Cuit")) == cuit_sel:
            candidatos.append(fila)

    if not candidatos:
        return None

    # Si viene PV desde el front (a futuro), lo usamos para desempatar
    if punto_venta is not None:
        pv_sel = _to_int(punto_venta)
        for fila in candidatos:
            if _to_int(fila.get("Punto Venta")) == pv_sel:
                return fila

    return candidatos[0]

def get_usuario_auth_source_dir(base_dir: str, cuit: int | str, servicio: str) -> str:
    cuit_str = str(int(float(cuit)))

    # nombres nuevos (sin acentos)
    carpeta_servicio_new = {
        "FEV1": "Autorizacion_ABC_sin_item",
        "MTXCA": "Autorizacion_AB_con_item",
    }[servicio]

    # nombres viejos (por compat)
    carpeta_servicio_old = {
        "FEV1": "Autorización A, B y C sin item",
        "MTXCA": "Autorización A y B con item",
    }[servicio]

    path_new = os.path.join(base_dir, "Usuarios", cuit_str, carpeta_servicio_new, "source")
    if os.path.isdir(os.path.join(base_dir, "Usuarios", cuit_str, carpeta_servicio_new)):
        return path_new

    return os.path.join(base_dir, "Usuarios", cuit_str, carpeta_servicio_old, "source")


def get_ultimo_token_path(base_dir: str, cuit: int | str, servicio: str) -> str:
    """
    Devuelve la ruta del TXT dentro del source del CUIT.
    """
    filename = {
        "FEV1": "UltimoTokenWSFEV1.txt",
        "MTXCA": "UltimoTokenWSMTXCA.txt",
    }[servicio]

    return os.path.join(get_usuario_auth_source_dir(base_dir, cuit, servicio), filename)

def get_cuit_seleccionado() -> int:
    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument("--cuit", type=int, required=False)
    args, _ = parser.parse_known_args()

    if args.cuit:
        return int(args.cuit)

    raise ValueError("No se recibió CUIT. Ejecutá con --cuit <CUIT> (lo manda la API).")

#* FLUJO PRINCIPAL -----------------------------------------------------------------------------

#! ---------------------------------------------------------------------------------------------
#! OPTIMIZAR FLUJO MAIN, CADA MARCA ROJA, DEBERÍA CONVERTIRSE EN UNA FUNCIÓN
#! ---------------------------------------------------------------------------------------------

# Variable global para almacenar el libro de trabajo actual
current_workbook = None

class TLSAdapter(HTTPAdapter):
    def __init__(self, ssl_context: ssl.SSLContext, **kwargs):
        self.ssl_context = ssl_context
        super().__init__(**kwargs)

    def init_poolmanager(self, connections, maxsize, block=False, **pool_kwargs):
        self.poolmanager = PoolManager(
            num_pools=connections,
            maxsize=maxsize,
            block=block,
            ssl_context=self.ssl_context,
            **pool_kwargs
        )

def crear_client_afip(wsdl_url: str) -> Client:
    ctx = ssl.create_default_context()
    ctx.set_ciphers("DEFAULT:@SECLEVEL=1")

    session = requests.Session()
    session.mount("https://servicios1.afip.gov.ar", TLSAdapter(ctx))

    transport = Transport(session=session, timeout=60)
    return Client(wsdl_url, transport=transport)



def main():

    abrir_log()
    limpiar_output_dir()

    escribir_log("--------------------------------------------------")
    escribir_log("Iniciando el proceso de Facturación Electrónica...")
    escribir_log(f"{obtener_timestamp()} - Horario de inicio")
    escribir_log("--------------------------------------------------")
    escribir_log("")

    print("--------------------------------------------------")
    print("Inicia el proceso de Facturación Electronica")
    print(f"{obtener_timestamp()} - Horario de inicio")
    print("Desarrollado por Matias Nicolas Franco")
    print("--------------------------------------------------")
    print("")

    limpiar_output_dir()

    # Revisión de archivos Facturacion.xlsx y configuración del servicio
    carpetas_input = [
        (FACTURA_A_DIR, 'A'),
        (FACTURA_B_DIR, 'B'),
        (FACTURA_C_DIR, 'C')
    ]

    for carpeta, tipo_factura in carpetas_input:
        ruta_archivo_excel = os.path.join(carpeta, 'Facturacion.xlsx')

        # Limpiar variables utilizadas
        ListaValidacionCAE.clear()

        if os.path.exists(ruta_archivo_excel):
            # Obtener configuración específica para el tipo de factura

            config_path = (f'{CONFIGURACIONES_DIR}/Config.xlsx')
            configuraciones = leer_configuraciones_desde_excel(config_path)

            # Selección de la configuración según el tipo de factura
            def obtener_configuracion(tipo_factura):

                if tipo_factura == 'A':
                    return configuraciones.get('Factura A')
                elif tipo_factura == 'B':
                    return configuraciones.get('Factura B')
                elif tipo_factura == 'C':
                    return configuraciones.get('Factura C')
                else:
                    escribir_log(f"{obtener_timestamp()} - Tipo de factura {tipo_factura} no reconocido.")
                    return None

            if not configuraciones:

                escribir_log(f"{obtener_timestamp()} - No se pudieron cargar las configuraciones desde el archivo Excel.")
                escribir_log("")
                cerrar_log()
                return

            cuit_seleccionado = get_cuit_seleccionado()

            config = obtener_configuracion_por_cuit(configuraciones, tipo_factura, cuit_seleccionado)

            if not config:
                escribir_log(f"{obtener_timestamp()} - No se encontró configuración para tipo {tipo_factura} y CUIT {cuit_seleccionado}.")
                continue

            numero_actividad = config["Numero Actividad"]

            if not config:
                escribir_log(f"{obtener_timestamp()} - No se pudo obtener la configuración para Factura {tipo_factura}.")
                continue
            print("")
            print("Contenido de config:", config)
            print("--------------------------------------------------")
            print("")

            #! ---------------------------------------------------------------------------------------------
            #! FUNCIÓN DE CLASIFICACIÓN POR TIPO DE FACTURA
            #! ---------------------------------------------------------------------------------------------

            if tipo_factura == 'A':
                if config["¿Con detalle de Items?"].lower() == "si":
                    wsdl_url = "https://servicios1.afip.gov.ar/wsmtxca/services/MTXCAService?wsdl"
                    servicio = "MTXCA"
                    print(f"{obtener_timestamp()} - Encontró archivo en la carpeta {tipo_factura}, se utilizará el servicio {servicio}")
                else:
                    wsdl_url = "https://servicios1.afip.gov.ar/wsfev1/service.asmx?WSDL"
                    servicio = "FEV1"
                    print(f"{obtener_timestamp()} - Encontró archivo en la carpeta {tipo_factura}, se utilizará el servicio {servicio}")

                tipo_comprobante = "1"
                punto_venta = config["Punto Venta"]
                plantilla_path = os.path.join(CONFIGURACIONES_DIR, "Plantilla_factura_A.xlsx")

            elif tipo_factura == 'B':
                if config["¿Con detalle de Items?"].lower() == "si":
                    wsdl_url = "https://servicios1.afip.gov.ar/wsmtxca/services/MTXCAService?wsdl"
                    servicio = "MTXCA"
                    print(f"{obtener_timestamp()} - Encontró archivo en la carpeta {tipo_factura}, se utilizará el servicio {servicio}")
                else:
                    wsdl_url = "https://servicios1.afip.gov.ar/wsfev1/service.asmx?WSDL"
                    servicio = "FEV1"
                    print(f"{obtener_timestamp()} - Encontró archivo en la carpeta {tipo_factura}, se utilizará el servicio {servicio}")

                tipo_comprobante = "6"
                punto_venta = config["Punto Venta"]
                plantilla_path = os.path.join(CONFIGURACIONES_DIR, "Plantilla_factura_B.xlsx")

            else:  # Factura C
                wsdl_url = "https://servicios1.afip.gov.ar/wsfev1/service.asmx?WSDL"
                servicio = "FEV1"
                tipo_comprobante = "11"
                punto_venta = config["Punto Venta"]
                plantilla_path = os.path.join(CONFIGURACIONES_DIR, "Plantilla_factura_C.xlsx")
                print(f"{obtener_timestamp()} - Encontró archivo en la carpeta {tipo_factura}, se utilizará el servicio {servicio}")

            token_path = get_usuario_auth_source_dir(BASE_DIR, cuit_seleccionado, servicio)
            ultimo_token_path = get_ultimo_token_path(BASE_DIR, cuit_seleccionado, servicio)

            # Validación para no explotar si falta la carpeta del CUIT
            if not os.path.isdir(token_path):
                escribir_log(f"{obtener_timestamp()} - No existe la carpeta de autorización para CUIT {cuit_seleccionado}: {token_path}")
                print(f"{obtener_timestamp()} - No existe la carpeta de autorización para CUIT {cuit_seleccionado}: {token_path}")
                continue

            print("")

            # Verificar y actualizar token y sign
            if verificar_token(ultimo_token_path):
                ejecutar_powershell(token_path, ultimo_token_path)

            token, sign = leer_ticket_response(token_path)
            if token is None or sign is None:
                escribir_log("--------------------------------------------------")
                escribir_log(f"{obtener_timestamp()} - Token o sign no válidos para {tipo_factura}, el proceso no puede continuar con este tipo de factura.")
                escribir_log("--------------------------------------------------")
                escribir_log("")

                print("--------------------------------------------------")
                print(f"{obtener_timestamp()} - Token o sign no válidos para {tipo_factura}, el proceso no puede continuar con este tipo de factura.")
                print("--------------------------------------------------")
                print("")

                continue

            # Crear auth y auth_cae después de verificar y actualizar el token
            auth = {
                "Token": str(token),
                "Sign": str(sign),
                "Cuit": int(config["Cuit"]),
                "PtoVta": int(punto_venta),
                "CbteTipo": int(tipo_comprobante)
            }

            try:
                client = crear_client_afip(wsdl_url)

                #! ---------------------------------------------------------------------------------------------
                #! FUNCIÓN PARA CORRER SEGÚN DATOS CARGADOS EN EXCEL
                #! ---------------------------------------------------------------------------------------------

                # Leer datos del archivo Excel
                try:
                    print("Inicia proceso de lectura de datos")
                    datos_excel, datos_nota_credito, datos_nota_debito = leer_datos_excel(ruta_archivo_excel, tipo_factura)
                    
                    #* SEGÚN FACTURAS
                    if datos_excel is not None:

                        realizando_facturas = True
                        realizando_nota_credito = False
                        realizando_nota_debito = False
                        
                        if servicio == "MTXCA":
                            print(f"{obtener_timestamp()} - Inicia proceso para armar solicitudes {servicio}")
                            solicitudes_mtxca = procesar_datos_mtxca(datos_excel, tipo_comprobante, punto_venta)
                            
                            auth_cae = {
                                "token": str(token),
                                "sign": str(sign),
                                "cuitRepresentada": int(config["Cuit"])
                            }

                            print(f"{obtener_timestamp()} - Inicia proceso para armar los cuerpos con las solicitudes {servicio}")
                            print("")
                            

                            if tipo_factura == "A":
                                cuerpos_solicitud = armar_cuerpo_solicitud_mtxca_A(solicitudes_mtxca, client, auth, servicio, numero_actividad)
                            else:
                                cuerpos_solicitud = armar_cuerpo_solicitud_mtxca_B(solicitudes_mtxca, client, auth, servicio, numero_actividad)


                        elif servicio == "FEV1":
                            print(f"{obtener_timestamp()} - Inicia proceso para armar solicitudes {servicio}")
                            print("")
                            solicitudes_fev1 = procesar_datos_fev1(datos_excel, tipo_comprobante, punto_venta)

                            auth_cae = {
                                "Token": str(token),
                                "Sign": str(sign),
                                "Cuit": int(config["Cuit"])
                            }   
            
                            print(f"{obtener_timestamp()} - Inicia proceso para armar los cuerpos con las solicitudes {servicio}")
                            print("")
                            cuerpos_solicitud = armar_cuerpo_solicitud_fev1(solicitudes_fev1, client, auth, servicio, numero_actividad)

                    #* SEGÚN NOTAS DE CREDITO
                    if datos_nota_credito is not None:

                        realizando_facturas = False
                        realizando_nota_credito = True
                        realizando_nota_debito = False
                        
                        if tipo_factura == "A":
                            tipo_comprobante = "3"
                            tipo_comprobante_asociado = "1" 
                            plantilla_path = os.path.join(CONFIGURACIONES_DIR, "Plantilla_credito_A.xlsx")

                        if tipo_factura == "B":
                            tipo_comprobante = "8"
                            tipo_comprobante_asociado = "6" 
                            plantilla_path = os.path.join(CONFIGURACIONES_DIR, "Plantilla_credito_B.xlsx")

                        if tipo_factura == "C":
                            tipo_comprobante = "13"
                            tipo_comprobante_asociado = "11" 
                            plantilla_path = os.path.join(CONFIGURACIONES_DIR, "Plantilla_credito_C.xlsx")

                        if servicio == "MTXCA":
                            print(f"{obtener_timestamp()} - Inicia proceso para armar solicitudes {servicio}, [NOTA CRÉDITO MTXCA]")
                            solicitudes_mtxca = procesar_datos_nota_credito_mtxca(datos_nota_credito, tipo_comprobante, tipo_comprobante_asociado, punto_venta)

                            auth_cae = {
                                "token": str(token),
                                "sign": str(sign),
                                "cuitRepresentada": int(config["Cuit"])
                            }
                            
                            auth = {
                                "Token": str(token),
                                "Sign": str(sign),
                                "Cuit": int(config["Cuit"]),
                                "PtoVta": int(punto_venta),
                                "CbteTipo": int(tipo_comprobante)
                            }

                            print(f"{obtener_timestamp()} - Inicia proceso para armar los cuerpos con las solicitudes {servicio}")
                            print("")
                            
                            cuerpos_solicitud = armar_cuerpo_nota_credito_mtxca(solicitudes_mtxca, client, auth, servicio)

                        elif servicio == "FEV1":
                            print(f"{obtener_timestamp()} - Inicia proceso para armar solicitudes {servicio}, [NOTA CRÉDITO FEV1]")
                            print("")
                            solicitudes_fev1 = procesar_datos_nota_credito_fev1(datos_nota_credito, tipo_comprobante, punto_venta)
                            
                            auth_cae = {
                                "Token": str(token),
                                "Sign": str(sign),
                                "Cuit": int(config["Cuit"])
                            }   

                            auth = {
                                "Token": str(token),
                                "Sign": str(sign),
                                "Cuit": int(config["Cuit"]),
                                "PtoVta": int(punto_venta),
                                "CbteTipo": int(tipo_comprobante)
                            }     

                            print(f"{obtener_timestamp()} - Inicia proceso para armar los cuerpos con las solicitudes {servicio}")
                            print("")
                            cuerpos_solicitud = armar_cuerpo_nota_credito_fev1(solicitudes_fev1, client, auth, servicio)


                    #* SEGÚN NOTAS DE DEBITO
                    if datos_nota_debito is not None:

                        realizando_facturas = False
                        realizando_nota_credito = False
                        realizando_nota_debito = True

                        if tipo_factura == "A":
                            tipo_comprobante = "2"
                            tipo_comprobante_asociado = "1" 
                            plantilla_path = os.path.join(CONFIGURACIONES_DIR, "Plantilla_debito_A.xlsx")

                        if tipo_factura == "B":
                            tipo_comprobante = "7"
                            tipo_comprobante_asociado = "6" 
                            plantilla_path = os.path.join(CONFIGURACIONES_DIR, "Plantilla_debito_B.xlsx")

                        if tipo_factura == "C":
                            tipo_comprobante = "12"
                            tipo_comprobante_asociado = "11" 
                            plantilla_path = os.path.join(CONFIGURACIONES_DIR, "Plantilla_debito_C.xlsx")

                        if servicio == "MTXCA":
                            print(f"{obtener_timestamp()} - Inicia proceso para armar solicitudes {servicio}, [NOTA DEBITO MTXCA]")
                            solicitudes_mtxca = procesar_datos_nota_debito_mtxca(datos_nota_debito, tipo_comprobante, tipo_comprobante_asociado, punto_venta)

                            auth_cae = {
                                "token": str(token),
                                "sign": str(sign),
                                "cuitRepresentada": int(config["Cuit"])
                            }

                            auth = {
                                "Token": str(token),
                                "Sign": str(sign),
                                "Cuit": int(config["Cuit"]),
                                "PtoVta": int(punto_venta),
                                "CbteTipo": int(tipo_comprobante)
                            }

                            print(f"{obtener_timestamp()} - Inicia proceso para armar los cuerpos con las solicitudes {servicio}")
                            print("")
                            
                            cuerpos_solicitud = armar_cuerpo_nota_debito_mtxca(solicitudes_mtxca, client, auth, servicio)

                        elif servicio == "FEV1":
                            print(f"{obtener_timestamp()} - Inicia proceso para armar solicitudes {servicio}, [NOTA DEBITO FEV1]")
                            print("")
                            solicitudes_fev1 = procesar_datos_nota_debito_fev1(datos_nota_debito, tipo_comprobante, punto_venta)

                            auth_cae = {
                                "Token": str(token),
                                "Sign": str(sign),
                                "Cuit": int(config["Cuit"])
                            }   
            
                            auth = {
                                "Token": str(token),
                                "Sign": str(sign),
                                "Cuit": int(config["Cuit"]),
                                "PtoVta": int(punto_venta),
                                "CbteTipo": int(tipo_comprobante)
                            }

                            print(f"{obtener_timestamp()} - Inicia proceso para armar los cuerpos con las solicitudes {servicio}")
                            print("")
                            cuerpos_solicitud = armar_cuerpo_nota_debito_fev1(solicitudes_fev1, client, auth, servicio)

                    print("")
                    for cuerpo in cuerpos_solicitud:
                        print("")
                        print("---------------------- CUERPO DE SOLICITUD SEGÚN JSON")
                        print(json.dumps(cuerpo, indent=4))
                        print("-----------------------------------------------------")
                        print("")

                    print("")

                    # Enviar solicitudes de CAE y manejar respuestas
                    print("-----------------------------------------------------")
                    print(f"{obtener_timestamp()} - Inicia proceso para enviar las solicitudes CAE")
                    print("")
                    enviar_solicitud_cae(servicio, client, auth_cae, cuerpos_solicitud)

                    #!----------------------------------------------------------------------------------------------------------
                    #! HACER COPIA DE PLANTILLA EXCEL PARA USAR EN COMPLETAR PLANTILLA
                    # Verificar si la plantilla existe
                    if not os.path.exists(plantilla_path):
                        escribir_log(f"{obtener_timestamp()} - Error: No se encontró la plantilla en {plantilla_path}")
                    else:
                        # Copiar la plantilla a OUTPUT_DIR antes de abrirla
                        plantilla_output_path = os.path.join(OUTPUT_DIR, os.path.basename(plantilla_path))
                        try:
                            shutil.copy(plantilla_path, plantilla_output_path)
                            escribir_log(f"{obtener_timestamp()} - Plantilla copiada a {plantilla_output_path}")
                            
                        except Exception as e:
                            escribir_log(f"{obtener_timestamp()} - Error al copiar la plantilla: {e}")


                    #!----------------------------------------------------------------------------------------------------------
                    #! HACER COPIA DE FACTURACIÓN, PARA COMPLETAR LOS DATOS FALTANTES
                    # Obtener la fecha y hora actual para generar el nombre único
                    timestamp = datetime.now().strftime("Corrida_%d-%m-%Y-%H_%M_%S")

                    # Verificar si la plantilla existe
                    if not os.path.exists(ruta_archivo_excel):
                        escribir_log(f"{obtener_timestamp()} - Error: No se encontró la plantilla en {ruta_archivo_excel}")
                    else:
                        # Construir la nueva ruta de salida en la carpeta FACTURAS_DIR con el nombre dinámico
                        input_facturacion_path = os.path.join(FACTURAS_DIR, f"{timestamp}.xlsx")
                        
                        try:
                            # Copiar y renombrar la plantilla
                            shutil.copy(ruta_archivo_excel, input_facturacion_path)
                            escribir_log(f"{obtener_timestamp()} - Plantilla copiada y renombrada a {input_facturacion_path}")
                            
                            # Guardar la ruta completa en una variable
                            ruta_completa_archivo = input_facturacion_path
                            escribir_log(f"{obtener_timestamp()} - Ruta completa del archivo: {ruta_completa_archivo}")
                            
                        except Exception as e:
                            escribir_log(f"{obtener_timestamp()} - Error al copiar la plantilla: {e}")

                    print("")
                    print("-----------------------------------------------------")
                    for elemento in ListaValidacionCAE:
                        print(elemento)
                    print("-----------------------------------------------------")
                    print("")

                    print("-----------------------------------------------------")

                    # FACTURAS NORMALES - A, B o C
                    if realizando_facturas:
                        tipo_nota = "Factura"
                        completar_plantilla(ruta_completa_archivo, plantilla_output_path, datos_excel, cuerpos_solicitud, config, ListaValidacionCAE, tipo_factura, tipo_nota)

                    # NOTAS DE CREDITO - A, B o C
                    if realizando_nota_credito:
                        tipo_nota = "Credito"
                        completar_plantilla(ruta_completa_archivo, plantilla_output_path, datos_nota_credito, cuerpos_solicitud, config, ListaValidacionCAE, tipo_factura, tipo_nota)

                    # NOTAS DE DEBITO - A, B o C
                    if realizando_nota_debito:
                        tipo_nota = "Debito"
                        completar_plantilla(ruta_completa_archivo, plantilla_output_path, datos_nota_debito, cuerpos_solicitud, config, ListaValidacionCAE, tipo_factura, tipo_nota)
                    print("")
                    print("-----------------------------------------------------")

                    cuerpos_solicitud.clear()
                    config.clear()
                    token = ""
                    sign = ""


                except Exception as e:
                    print(f"{obtener_timestamp()} - Falló en: {e}")
                    escribir_log(f"{obtener_timestamp()} - Falló en: {e}")

            except Exception as e:
                escribir_log(f"{obtener_timestamp()} - Error al conectar al web service para {tipo_factura}: {e}")

                continue
        else:
            escribir_log("--------------------------------------------------")
            escribir_log("No se encontro archivo Facturacion.xlsx, en ninguno de las 3 carpetas listadas")
            escribir_log("> input/Factura A")
            escribir_log("> input/Factura B")
            escribir_log("> input/Factura C")
            escribir_log(f"{obtener_timestamp()} - Hora de finalización")         
            escribir_log("--------------------------------------------------")
            escribir_log("")

            print("--------------------------------------------------")
            print("No se encontro archivo Facturacion.xlsx, en ninguno de las 3 carpetas listadas")
            print(f"{obtener_timestamp()} - Horario de Finalización")
            print("--------------------------------------------------")
            print("")
  
    try:
        print("")
        print("-----------------------------------------------------")
        convertir_xlsx_a_pdf(OUTPUT_DIR, FACTURAS_DIR)
        print("-----------------------------------------------------")
        print("")


        cerrar_log()
        
    except Exception as e:
        print(f"{obtener_timestamp()} - Falló en: {e}")
        escribir_log(f"{obtener_timestamp()} - Falló en: {e}")

    cerrar_log()

if __name__ == "__main__":
    main()