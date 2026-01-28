
from backend.log import escribir_log, obtener_timestamp
from backend.config import validar_y_transformar_fecha
import pandas as pd


#! ------------------------------------------------------------------------------------------------------------------------------------------------------
#!
#! SISTEMAS PARA FEV1
#! 
#! ------------------------------------------------------------------------------------------------------------------------------------------------------

def to_float(v):
    if v is None:
        return 0.0
    if isinstance(v, str) and v.strip() == "":
        return 0.0
    return float(v)

def procesar_datos_fev1(datos, tipo_comprobante, punto_venta):
    escribir_log(f"{obtener_timestamp()} - Iniciando el procesamiento de datos para FEV1.")

    if not isinstance(datos, pd.DataFrame):
        escribir_log(f"{obtener_timestamp()} - Error: Se esperaba un DataFrame, pero se recibió un objeto de tipo {type(datos)}")
        return []

    solicitudes = []
    comprobante_actual = None
    total_importe = 0.0
    total_iva = 0.0
    total_neto = 0.0

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

    for index, fila in datos.iterrows():
        tipo_fila = fila.iloc[0].strip().lower()

        if tipo_fila == "comprobante":
            escribir_log(f"{obtener_timestamp()} - Procesando comprobante en la fila {index}.")

            if comprobante_actual:
                # Finaliza el procesamiento del comprobante actual
                comprobante_actual["importe_total"] = total_importe
                comprobante_actual["ImpIVA"] = total_iva
                comprobante_actual["ImpNeto"] = total_neto
                solicitudes.append(comprobante_actual)

            # Reinicia los acumuladores para el nuevo comprobante
            total_importe = to_float(fila.iloc[12])
            total_iva = to_float(fila.iloc[18])
            total_neto = to_float(fila.iloc[12])

            comprobante_actual = {
                "Concepto": int(fila.iloc[21]),
                "tipo_comprobante": tipo_comprobante,
                "punto_venta": punto_venta,
                "fecha_emision": str(validar_y_transformar_fecha(fila.iloc[1], "FEV1")),
                "fecha_vencimiento_pago": str(validar_y_transformar_fecha(fila.iloc[1], "FEV1")),
                "fecha_servicio_desde": str(validar_y_transformar_fecha(fila.iloc[2], "FEV1")),
                "fecha_servicio_hasta": str(validar_y_transformar_fecha(fila.iloc[3], "FEV1")),
                "tipo_documento": int(doc_tipo.get(fila.iloc[5], 99)),
                "numero_documento": int(0 if pd.isna(fila.iloc[7]) else fila.iloc[7]),
                "ID_Moneda": "PES",
                "moneda_cotizacion": 1,
                "importe_total": 0.0,
                "ImpIVA": 0.0,
                "ImpTotConc": float(0 if to_float(pd.isna(fila.iloc[16])) else to_float(fila.iloc[16])),
                "ImpNeto": 0.0,
                "ImpOpEx": float(0 if to_float(pd.isna(fila.iloc[17])) else to_float(fila.iloc[17])),
                "ImpTrib": float(0 if to_float(pd.isna(fila.iloc[19])) else to_float(fila.iloc[19])),
                "importe_bonificacion": fila.iloc[16],
                "codigo_OTributos": fila.iloc[22],
                "descripcion_OTributos": fila.iloc[23],
                "base_imponible_OTributos": fila.iloc[24],
                "Iva": []
            }

        elif tipo_fila == "item":
            escribir_log(f"{obtener_timestamp()} - Procesando ítem en la fila {index}.")
            importeItem = float(0 if pd.isna(fila.iloc[12]) else fila.iloc[12])
            iva_item = float(0 if pd.isna(fila.iloc[18]) else fila.iloc[18])

            total_importe += importeItem
            total_iva += iva_item
            total_neto += importeItem

            if tipo_comprobante == 1 or tipo_comprobante == 6:
                comprobante_actual["ImpNeto"] = total_neto
                comprobante_actual["ImpIVA"] = total_iva
                comprobante_actual["Iva"].append({
                    "AlicIva": {
                        "Id": 5,
                        "BaseImp": total_neto,
                        "Importe": total_iva
                    }
                })

    if comprobante_actual:
        # Finaliza el último comprobante
        comprobante_actual["importe_total"] = total_importe
        comprobante_actual["ImpIVA"] = total_iva
        comprobante_actual["ImpNeto"] = total_neto
        solicitudes.append(comprobante_actual)

    escribir_log(f"{obtener_timestamp()} - Procesamiento de datos FEV1 finalizado. Total de solicitudes generadas: {len(solicitudes)}.")
    return solicitudes