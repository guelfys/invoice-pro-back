
from backend.log import escribir_log, obtener_timestamp
from backend.config import validar_y_transformar_fecha
import pandas as pd

#! ------------------------------------------------------------------------------------------------------------------------------------------------------
#!
#! SISTEMAS PARA MTXCA
#! 
#! ------------------------------------------------------------------------------------------------------------------------------------------------------

def procesar_datos_mtxca(datos, tipo_comprobante, punto_venta):
    escribir_log(f"{obtener_timestamp()} - Iniciando el procesamiento de datos para MTXCA.")
    
    if not isinstance(datos, pd.DataFrame):
        escribir_log(f"{obtener_timestamp()} - Error: Se esperaba un DataFrame, pero se recibió un objeto de tipo {type(datos)}")
        return []

    solicitudes = []
    comprobante_actual = None
    items_actuales = []
    contador_producto = 1

    doc_tipo = {
        "CUIT": 80,
        "CUIL": 86,
        "CDI": 87,
        "CI Extranjera": 91,
        "Pasaporte": 94,
        "DNI": 96,
        "D.N.I": 96,
        "Otro": 99,
        "IVA Responsable inscripto": 1,
        "IVA Responsable no inscripto": 2,
        "IVA No responsable": 3,
        "IVA Sujeto exento": 4,
        "Consumidor final": 5,
        "Responsable monotributo": 6,
        "Sujeto no categorizado": 7,
        "Proveedor del exterior": 8,
        "Cliente del exterior": 9,
        "IVA Liberado - Ley 19.640": 10,
        "IVA Responsable inscripto - Agente de percepción": 11,
        "Pequeño contribuyente eventual": 12,
        "Monotributista social": 13,
        "Pequeño contribuyente eventual social": 14,
        "IVA No alcanzado": 15
    }

    iva_receptor = {
        "responsable inscripto": 1,
        "responsable no inscripto": 2,
        "iva responsable inscripto": 1,
        "iva responsable no inscripto": 2,
        "iva no responsable": 3,
        "iva sujeto exento": 4,
        "consumidor final": 5,
        "responsable monotributo": 6,
        "monotributista": 6,
        "sujeto no categorizado": 7,
        "proveedor del exterior": 8,
        "cliente del exterior": 9,
        "iva liberado - ley 19.640": 10,
        "iva responsable inscripto - agente de percepción": 11,
        "pequeño contribuyente eventual": 12,
        "monotributista social": 13,
        "pequeño contribuytente eventual social": 14,
        "iva no alcanzado": 15,
        "1": 1,
        "2": 2,
        "3": 3,
        "4": 4,
        "5": 5,
        "6": 6,
        "7": 7,
        "8": 8,
        "9": 9,
        "10": 10,
        "11": 11,
        "12": 12,
        "13": 13,
        "14": 14,
        "15": 15
    }

    for index, fila in datos.iterrows():
        tipo_fila = fila.iloc[0].strip().lower()

        if tipo_fila == "comprobante":
            contador_producto = 1
            if comprobante_actual:
                comprobante_actual["items"] = items_actuales
                solicitudes.append(comprobante_actual)
            
            comprobante_actual = {
                "codigoConcepto": fila.iloc[21],
                "tipo_comprobante": tipo_comprobante,
                "punto_venta": punto_venta,
                "fecha_emision": validar_y_transformar_fecha(fila.iloc[1], "MTXCA"),
                "fecha_servicio_desde": validar_y_transformar_fecha(fila.iloc[2], "MTXCA"),
                "fecha_servicio_hasta": validar_y_transformar_fecha(fila.iloc[3], "MTXCA"),
                "fecha_vencimiento_pago": validar_y_transformar_fecha(fila.iloc[1], "MTXCA"),
                "tipo_documento": int(doc_tipo.get(fila.iloc[5], 99)),
                "numero_documento": int(0 if pd.isna(fila.iloc[7]) else fila.iloc[7]),
                "importe_gravado": 0.0,
                "importe_no_gravado": 0.0,
                "importe_exento": float(0 if pd.isna(fila.iloc[17]) else fila.iloc[17]),
                "importe_otros_tributos": float(0 if pd.isna(fila.iloc[19]) else fila.iloc[19]),
                "importe_total": float(fila.iloc[12] if pd.notna(fila.iloc[12]) else 0.0),
                "importe_bonificacion": fila.iloc[16],
                "codigo_OTributos": fila.iloc[22],
                "descripcion_OTributos": fila.iloc[23],
                "base_imponible_OTributos": fila.iloc[24],
                "iva_receptor": int(iva_receptor.get(str(fila.iloc[4]).lower(), 5)),
                "items": []
            }
            items_actuales = []

            if float(fila.iloc[19]) > 0:
                comprobante_actual["importe_otros_tributos"] = float(0 if pd.isna(fila.iloc[19]) else fila.iloc[19])

            print(f"{obtener_timestamp()} - contador del producto: {contador_producto}")
            items_actuales.append({
                "unidadMtx": int(fila.iloc[25]),
                "codigoMtx": str(fila.iloc[26]),
                "codigo": int(contador_producto),
                "descripcion": str(fila.iloc[10] if pd.notna(fila.iloc[10]) else ''),
                "cantidad": int(fila.iloc[9] if pd.notna(fila.iloc[9]) else 0),
                "codigoUnidadMedida": int(6),
                "precioUnitario": float(fila.iloc[11] if pd.notna(fila.iloc[11]) else 0),
                "importeItem": float(fila.iloc[12] if pd.notna(fila.iloc[12]) else 0),
                "codigoCondicionIVA": str(fila.iloc[27]),
                "importeIVA": float(fila.iloc[18] if pd.notna(fila.iloc[18]) else 0)
            })
            contador_producto += 1

        elif tipo_fila == "item":
            print(f"{obtener_timestamp()} - contador del producto: {contador_producto}")
            items_actuales.append({
                "unidadMtx": int(fila.iloc[25]),
                "codigoMtx": str(fila.iloc[26]),
                "codigo": int(contador_producto),
                "descripcion": str(fila.iloc[10] if pd.notna(fila.iloc[10]) else ''),
                "cantidad": int(fila.iloc[9] if pd.notna(fila.iloc[9]) else 0),
                "codigoUnidadMedida": int(6),
                "precioUnitario": float(fila.iloc[11] if pd.notna(fila.iloc[11]) else 0),
                "importeItem": float(fila.iloc[12] if pd.notna(fila.iloc[12]) else 0),
                "codigoCondicionIVA": str(fila.iloc[27]),
                "importeIVA": float(fila.iloc[18] if pd.notna(fila.iloc[18]) else 0)
            })
            contador_producto += 1

    if comprobante_actual:
        comprobante_actual["items"] = items_actuales
        solicitudes.append(comprobante_actual)
    
    escribir_log(f"{obtener_timestamp()} - Procesamiento de datos MTXCA finalizado. Total de solicitudes generadas: {len(solicitudes)}.")
    return solicitudes
