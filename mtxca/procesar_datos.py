from backend.log import escribir_log, obtener_timestamp
from backend.config import validar_y_transformar_fecha
import pandas as pd

#! ------------------------------------------------------------------------------------------------------------------------------------------------------
#!
#! SISTEMAS PARA MTXCA
#! 
#! ------------------------------------------------------------------------------------------------------------------------------------------------------

def _pick_value(fila, *columnas):
    for col in columnas:
        if col in fila.index:
            value = fila.get(col)
            if pd.notna(value) and str(value).strip() != "":
                return value
    return None


def _safe_date(value):
    if value is None or pd.isna(value) or str(value).strip() == "":
        return None
    try:
        return validar_y_transformar_fecha(value, "MTXCA")
    except Exception:
        return None


def _safe_int(value, default=0):
    if value is None or pd.isna(value) or str(value).strip() == "":
        return default
    try:
        return int(float(value))
    except Exception:
        return default


def _safe_float(value, default=0.0):
    if value is None or pd.isna(value) or str(value).strip() == "":
        return default
    try:
        return float(value)
    except Exception:
        return default


def _safe_str(value, default=""):
    if value is None or pd.isna(value):
        return default
    return str(value).strip()


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
        tipo_fila = _safe_str(_pick_value(fila, "Tipo Dato", "Tipo")).lower()

        if tipo_fila == "comprobante":
            contador_producto = 1
            if comprobante_actual:
                comprobante_actual["items"] = items_actuales
                solicitudes.append(comprobante_actual)

            codigo_concepto = _safe_int(_pick_value(fila, "Concepto"), 1)

            fecha_emision = _safe_date(_pick_value(fila, "Fecha"))
            fecha_servicio_desde = _safe_date(_pick_value(fila, "Periodo Desde", "Fecha servicio desde"))
            fecha_servicio_hasta = _safe_date(_pick_value(fila, "Periodo Hasta", "Fecha servicio hasta"))
            fecha_vencimiento_pago_excel = _safe_date(_pick_value(fila, "Fecha vencimiento pago"))

            # Normalización por concepto
            if codigo_concepto == 1:
                # Productos: no informar fechas de servicio ni vencimiento
                fecha_servicio_desde = None
                fecha_servicio_hasta = None
                fecha_vencimiento_pago = None

            elif codigo_concepto in (2, 3):
                # Servicios / Productos y Servicios:
                # si faltan fechas, usar fecha de emisión para evitar rechazo
                if fecha_servicio_desde is None:
                    fecha_servicio_desde = fecha_emision

                if fecha_servicio_hasta is None:
                    fecha_servicio_hasta = fecha_emision

                fecha_vencimiento_pago = (
                    fecha_vencimiento_pago_excel
                    if fecha_vencimiento_pago_excel is not None
                    else fecha_emision
                )

            else:
                # fallback defensivo
                fecha_servicio_desde = None
                fecha_servicio_hasta = None
                fecha_vencimiento_pago = None

            comprobante_actual = {
                "codigoConcepto": codigo_concepto,
                "tipo_comprobante": tipo_comprobante,
                "punto_venta": punto_venta,
                "fecha_emision": fecha_emision,
                "fecha_servicio_desde": fecha_servicio_desde,
                "fecha_servicio_hasta": fecha_servicio_hasta,
                "fecha_vencimiento_pago": fecha_vencimiento_pago,
                "tipo_documento": int(doc_tipo.get(_safe_str(_pick_value(fila, "Tipo Doc")), 99)),
                "numero_documento": _safe_int(_pick_value(fila, "Documento"), 0),
                "importe_gravado": 0.0,
                "importe_no_gravado": 0.0,
                "importe_exento": _safe_float(_pick_value(fila, "Importe Op Ex"), 0.0),
                "importe_otros_tributos": _safe_float(_pick_value(fila, "Importe Tributos"), 0.0),
                "importe_total": _safe_float(_pick_value(fila, "Total"), 0.0),
                "importe_bonificacion": _safe_float(_pick_value(fila, "Importe Bonificación"), 0.0),
                "codigo_OTributos": _pick_value(fila, "Codigo Otros Tributos"),
                "descripcion_OTributos": _pick_value(fila, "Descripcion Otros Tributos"),
                "base_imponible_OTributos": _pick_value(fila, "Base Imponible otros Tributos"),
                "iva_receptor": int(iva_receptor.get(_safe_str(_pick_value(fila, "Condicion frente al IVA")).lower(), 5)),
                "items": []
            }
            items_actuales = []

            print(f"{obtener_timestamp()} - contador del producto: {contador_producto}")
            items_actuales.append({
                "unidadMtx": _safe_int(_pick_value(fila, "UnidadMtx"), 0),
                "codigoMtx": _safe_str(_pick_value(fila, "CodigoMtx")),
                "codigo": int(contador_producto),
                "descripcion": _safe_str(_pick_value(fila, "Descripcion")),
                "cantidad": _safe_int(_pick_value(fila, "Cant.", "Cantidad"), 0),
                "codigoUnidadMedida": int(6),
                "precioUnitario": _safe_float(_pick_value(fila, "$ Unit.", "Precio Unitario"), 0),
                "importeItem": _safe_float(_pick_value(fila, "Total", "Importe Total"), 0),
                "codigoCondicionIVA": _safe_str(_pick_value(fila, "Codigo Condición IVA", "Codigo Condicion IVA")),
                "importeIVA": _safe_float(_pick_value(fila, "Importe IVA"), 0)
            })
            contador_producto += 1

        elif tipo_fila == "item":
            print(f"{obtener_timestamp()} - contador del producto: {contador_producto}")
            items_actuales.append({
                "unidadMtx": _safe_int(_pick_value(fila, "UnidadMtx"), 0),
                "codigoMtx": _safe_str(_pick_value(fila, "CodigoMtx")),
                "codigo": int(contador_producto),
                "descripcion": _safe_str(_pick_value(fila, "Descripcion")),
                "cantidad": _safe_int(_pick_value(fila, "Cant.", "Cantidad"), 0),
                "codigoUnidadMedida": int(6),
                "precioUnitario": _safe_float(_pick_value(fila, "$ Unit.", "Precio Unitario"), 0),
                "importeItem": _safe_float(_pick_value(fila, "Total", "Importe Total"), 0),
                "codigoCondicionIVA": _safe_str(_pick_value(fila, "Codigo Condición IVA", "Codigo Condicion IVA")),
                "importeIVA": _safe_float(_pick_value(fila, "Importe IVA"), 0)
            })
            contador_producto += 1

    if comprobante_actual:
        comprobante_actual["items"] = items_actuales
        solicitudes.append(comprobante_actual)

    escribir_log(f"{obtener_timestamp()} - Procesamiento de datos MTXCA finalizado. Total de solicitudes generadas: {len(solicitudes)}.")
    return solicitudes