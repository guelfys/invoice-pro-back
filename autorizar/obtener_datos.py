import os
import chardet
import subprocess
import time
import glob
import shutil
import xml.etree.ElementTree as ET
from backend.log import escribir_log, obtener_timestamp, datetime
from variables import BASE_DIR

#? Función para leer y verificar el token
def verificar_token(Ultimo_Token_Path):
    mensaje = "Verificando token...\n"
    try:
        with open(Ultimo_Token_Path, 'r') as file:
            ultimo_token = file.read().strip()
            ultima_fecha = datetime.strptime(ultimo_token, "%d/%m/%Y %H:%M")
            ahora = datetime.now()
            diferencia = (ahora - ultima_fecha).total_seconds() / 60
            mensaje += f"Último token generado el: {ultima_fecha}\nTiempo actual: {ahora}\nDiferencia en minutos: {diferencia}\n"
            if diferencia < 10:
                mensaje += "El token se generó hace menos de 10 minutos, no es necesario ejecutar el script PowerShell.\n"
                escribir_log(f"{obtener_timestamp} - {mensaje}")
                return False
    except Exception as e:
        mensaje += f"{obtener_timestamp()} - Error al leer {Ultimo_Token_Path}: {e}\n"
    escribir_log(f"{obtener_timestamp()} - {mensaje}")
    return True

#? Función para ejecutar script PowerShell si es necesario
def ejecutar_powershell(Login_Ticket_Path, Ultimo_Token_Path):
    script_path = os.path.join(Login_Ticket_Path, 'wsaa-cliente.ps1')

    escribir_log(f"{obtener_timestamp()} - Ejecutando PS1: {script_path}")
    escribir_log(f"{obtener_timestamp()} - CWD: {Login_Ticket_Path}")

    # borrar respuestas anteriores
    eliminar_archivos_patron(Login_Ticket_Path, "*loginTicketResponse*.xml")
    eliminar_archivos_patron(Login_Ticket_Path, "*loginTicketResponse*.txt")
    time.sleep(1)

    result = subprocess.run(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", script_path],
        cwd=Login_Ticket_Path,
        capture_output=True,
        text=True
    )

    escribir_log(f"{obtener_timestamp()} - PS1 returncode: {result.returncode}")
    if result.stdout:
        escribir_log(f"{obtener_timestamp()} - PS1 stdout:\n{result.stdout}")
    if result.stderr:
        escribir_log(f"{obtener_timestamp()} - PS1 stderr:\n{result.stderr}")

    # ¿se generó TicketResponse válido?
    hay_ok = any(
        ("TicketResponse" in f and f.endswith(".xml") and "ERROR" not in f)
        for f in os.listdir(Login_Ticket_Path)
    )

    if result.returncode != 0 or not hay_ok:
        escribir_log(f"{obtener_timestamp()} - PS1 falló o no generó TicketResponse OK. NO se actualiza UltimoToken.")
        limpiar_archivos_xml(Login_Ticket_Path)
        raise RuntimeError("WSAA no generó TicketResponse válido. Revisar log PS1 stdout/stderr y archivo -ERROR.")

    # recién acá: actualizar “último token”
    with open(Ultimo_Token_Path, 'w') as file:
        file.write(datetime.now().strftime("%d/%m/%Y %H:%M"))

    limpiar_archivos_xml(Login_Ticket_Path)
    escribir_log(f"{obtener_timestamp()} - TicketResponse generado OK y UltimoToken actualizado.")

#? Función para limpiar archivos XML no deseados
def limpiar_archivos_xml(Login_Ticket_Path):
    mensaje = "Limpiando archivos XML no deseados...\n"
    for file in os.listdir(Login_Ticket_Path):
        if file.endswith('.xml') and 'TicketResponse' not in file:
            mensaje += f"Eliminando archivo: {file}\n"
            os.remove(os.path.join(Login_Ticket_Path, file))
        if file.endswith('.xml.cms-DER') or file.endswith('.xml.cms-DER-b64'):
            mensaje += f"Eliminando archivo: {file}\n"
            os.remove(os.path.join(Login_Ticket_Path, file))
    escribir_log(f"{obtener_timestamp()} - {mensaje}")

#? Eliminar archivos según patrón
def eliminar_archivos_patron(directorio, patron):
    archivos = glob.glob(os.path.join(directorio, patron))
    for archivo in archivos:
        escribir_log(f"Eliminando archivo: {archivo}")
        os.remove(archivo)

#? Eliminar todo lo que haya en output_dir
def limpiar_output_dir():
    # Definir la ruta de la carpeta 'output'
    OUTPUT_DIR = os.path.join(BASE_DIR, 'output')

    # Verificar si la carpeta existe
    if os.path.exists(OUTPUT_DIR):
        # Recorrer todos los archivos y subcarpetas dentro de 'output'
        for filename in os.listdir(OUTPUT_DIR):
            file_path = os.path.join(OUTPUT_DIR, filename)

            # Verificar si es un archivo o una carpeta
            try:
                if os.path.isfile(file_path) or os.path.islink(file_path):
                    os.unlink(file_path)  # Eliminar el archivo o enlace simbólico
                elif os.path.isdir(file_path):
                    shutil.rmtree(file_path)  # Eliminar la carpeta y su contenido
            except Exception as e:
                print(f"{obtener_timestamp()} - Error al eliminar {file_path}. Razón: {e}")
    else:
        print(f"{obtener_timestamp()} - La carpeta {OUTPUT_DIR} no existe.")

#? Función para detectar la codificación del archivo
def detectar_codificacion(file_path):
    with open(file_path, 'rb') as file:
        raw_data = file.read()
        result = chardet.detect(raw_data)
    return result['encoding']

#? Función para leer el archivo loginTicketResponse.xml
def leer_ticket_response(Login_Ticket_Path):
    candidatos = []
    for file in os.listdir(Login_Ticket_Path):
        if 'TicketResponse' in file and file.endswith('.xml') and 'ERROR' not in file:
            full = os.path.join(Login_Ticket_Path, file)
            candidatos.append(full)

    if not candidatos:
        escribir_log(f"{obtener_timestamp()} - AFIP no logró generar un Token y Sign validos, volver a probar en 10 minutos")
        print(f"{obtener_timestamp()} - AFIP no logró generar un Token y Sign validos, volver a probar en 10 minutos")
        raise FileNotFoundError(f"{obtener_timestamp()} - No se encontró un archivo de TicketResponse válido.")

    # elegir el más reciente
    ticket_response_file = max(candidatos, key=os.path.getmtime)

    try:
        encoding = detectar_codificacion(ticket_response_file)
        with open(ticket_response_file, 'r', encoding=encoding) as file:
            tree = ET.parse(file)
            root = tree.getroot()
            token = root.findtext('.//token')
            sign = root.findtext('.//sign')
            return token, sign
    except Exception as e:
        escribir_log(f"{obtener_timestamp()} - Error al leer TicketResponse: {e}")
        return None, None