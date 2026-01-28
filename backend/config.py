import datetime
import pandas as pd
import openpyxl
from backend.log import obtener_timestamp

"""
#? Detecta el formato de la fecha de manera flexible y devuelve un objeto datetime.
#? Intenta con múltiples formatos de fecha comunes.
"""
def detectar_formato_fecha(fecha):
    formatos_fecha = ['%Y-%m-%d', '%d/%m/%Y', '%Y/%m/%d', '%d-%m-%Y']
    for formato in formatos_fecha:
        try:
            return datetime.datetime.strptime(fecha, formato)
        except ValueError:
            continue
    raise ValueError(f"{obtener_timestamp()} - Formato de fecha no reconocido.")

#? Función para validar y transformar la fecha según el servicio
def validar_y_transformar_fecha(fecha, servicio):
    try:
        # Detectar y convertir la fecha si es cadena
        if isinstance(fecha, str):
            fecha_dt = detectar_formato_fecha(fecha)
        else:
            # Si ya es un objeto datetime
            fecha_dt = pd.to_datetime(fecha, errors='coerce')
        
        if pd.isnull(fecha_dt):
            raise ValueError("Fecha inválida")

        # Devolver en el formato correspondiente según el servicio
        if servicio == 'MTXCA':
            return fecha_dt.strftime('%Y-%m-%d')  # Formato YYYY-MM-DD
        elif servicio == 'FEV1':
            return fecha_dt.strftime('%Y%m%d')    # Formato YYYYMMDD
        else:
            raise ValueError("Servicio inválido. Utiliza 'MTXCA' o 'FEV1'.")
    
    except Exception as e:
        print(f"{obtener_timestamp()} - Error al procesar la fecha: {e}")
        return fecha

#? Eliminar ".0" de números enteros
def remove_dot_zero(value):
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)

#? Función para obtener los datos de cada hoja de Excel por separado
def leer_datos_excel(ruta_archivo, tipo_factura):
    wb = openpyxl.load_workbook(ruta_archivo)
    
    datos_factura = None
    datos_nota_credito = None
    datos_nota_debito = None

    hoja_leer1 = 'Factura {}'.format(tipo_factura)
    hoja_leer2 = 'Nota Credito {}'.format(tipo_factura) 
    hoja_leer3 = 'Nota Debito {}'.format(tipo_factura)

    print(hoja_leer1)
    print(hoja_leer2)
    print(hoja_leer3)

    # Leer Hoja1 (Facturas)
    if hoja_leer1 in wb.sheetnames:
        try:
            hoja_factura = wb[hoja_leer1]
        except:
            print("No se encontró la facturación general de Factura {}, para realizar la lectura de datos".format(tipo_factura))
        # Leer todas las filas de Hoja1 (ignorando el encabezado)
        datos_factura = [
            [celda if celda is not None else "" for celda in fila]
            for fila in hoja_factura.iter_rows(min_row=2, values_only=True)  # Ignorar encabezado
        ]
        # Convertir en DataFrame solo si hay datos
        if len(datos_factura) > 0:
            columnas_factura = [celda.value for celda in hoja_factura[1]]  # Tomar los encabezados de la primera fila
            datos_factura = pd.DataFrame(datos_factura, columns=columnas_factura)
        else:
            datos_factura = None

    # Leer Hoja2 (Nota de Crédito)
    if hoja_leer2 in wb.sheetnames:
        try:
            hoja_nota_credito = wb[hoja_leer2]
        except:
            print("No se encontró la Nota de Credito {}, para realizar la lectura de datos".format(tipo_factura))
        # Leer todas las filas de Hoja2 (ignorando el encabezado)
        datos_nota_credito = [
            [celda if celda is not None else "" for celda in fila]
            for fila in hoja_nota_credito.iter_rows(min_row=2, values_only=True)  # Ignorar encabezado
        ]
        # Convertir en DataFrame solo si hay datos
        if len(datos_nota_credito) > 0:
            columnas_nota_credito = [celda.value for celda in hoja_nota_credito[1]]  # Tomar los encabezados de la primera fila
            datos_nota_credito = pd.DataFrame(datos_nota_credito, columns=columnas_nota_credito)
        else:
            datos_nota_credito = None

    # Leer Hoja3 (Nota de Debito)
    if hoja_leer3 in wb.sheetnames:
        try:
            hoja_nota_debito = wb[hoja_leer3]
        except:
            print("No se encontró la Nota de Debito {}, para realizar la lectura de datos".format(tipo_factura))
        # Leer todas las filas de Hoja3 (ignorando el encabezado)
        datos_nota_debito = [
            [celda if celda is not None else "" for celda in fila]
            for fila in hoja_nota_debito.iter_rows(min_row=2, values_only=True)  # Ignorar encabezado
        ]
        # Convertir en DataFrame solo si hay datos
        if len(datos_nota_debito) > 0:
            columnas_nota_debito = [celda.value for celda in hoja_nota_debito[1]]  # Tomar los encabezados de la primera fila
            datos_nota_debito = pd.DataFrame(datos_nota_debito, columns=columnas_nota_debito)
        else:
            datos_nota_debito = None

    print(datos_factura)

    return datos_factura, datos_nota_credito, datos_nota_debito
