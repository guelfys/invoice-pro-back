import os
from variables import LOG_DIR
from datetime import datetime
import random

#? Función para generar el timestamp
def obtener_timestamp():
    return datetime.now().strftime("[%H:%M:%S]")

#? Función para generar un nombre único para el archivo de log
def generar_nombre_log():
    id_corrida = f"ID{random.randint(10000, 99999)}"
    fecha = datetime.now().strftime("%d-%m-%Y")
    return os.path.join(LOG_DIR, f"FacturacionElectronica_{id_corrida}_{fecha}.txt")

#? Función para abrir el archivo de log
def abrir_log():
    global log_file
    if not os.path.exists(LOG_DIR):
        os.makedirs(LOG_DIR)
    nombre_log = generar_nombre_log()
    log_file = open(nombre_log, 'a')

#? Función para escribir mensajes en el archivo de log
def escribir_log(mensaje):
    if log_file:
        log_file.write(mensaje + "\n")

#? Función para cerrar el archivo de log
def cerrar_log():
    if log_file:
        log_file.close()