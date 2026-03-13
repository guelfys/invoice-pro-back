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
    if pd.isna(v):
        return 0.0
    try:
        return float(v)
    except Exception:
        return 0.0


def _normalize_key(value):
    if value is None:
        return ""
    return (
        str(value)
        .strip()
        .lower()
        .replace("\n", " ")
        .replace("\r", " ")
    )


def _build_column_lookup(df: pd.DataFrame):
    return {_normalize_key(col): col for col in df.columns if col is not None}


def _get_cell_value(row, column_lookup, *possible_names):
    for name in possible_names:
        real_col = column_lookup.get(_normalize_key(name))
        if real_col is not None:
            return row.get(real_col)
    return None


def _get_text(row, column_lookup, *possible_names):
    value = _get_cell_value(row, column_lookup, *possible_names)
    if value is None or pd.isna(value):
        return ""
    return str(value).strip()


def _get_int(row, column_lookup, *possible_names, default=0):
    value = _get_cell_value(row, column_lookup, *possible_names)
    if value is None or pd.isna(value) or str(value).strip() == "":
        return default
    try:
        return int(float(value))
    except Exception:
        return default


def _get_float(row, column_lookup, *possible_names, default=0.0):
    value = _get_cell_value(row, column_lookup, *possible_names)
    return to_float(value) if value is not None else default


def _safe_date(value):
    if value is None or pd.isna(value):
        return None
    if isinstance(value, str) and value.strip() == "":
        return None
    try:
        return str(validar_y_transformar_fecha(value, "FEV1"))
    except Exception:
        return None


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

    column_lookup = _build_column_lookup(datos)

    doc_tipo = {
        "cuit": 80,
        "cuil": 86,
        "cdi": 87,
        "ci extranjera": 91,
        "pasaporte": 94,
        "dni": 96,
        "d.n.i": 96,
        "otro": 99,
        "iva responsable inscripto": 1,
        "iva responsable no inscripto": 2,
        "iva no responsable": 3,
        "iva sujeto exento": 4,
        "consumidor final": 5,
        "responsable monotributo": 6,
        "sujeto no categorizado": 7,
        "proveedor del exterior": 8,
        "cliente del exterior": 9,
        "iva liberado - ley 19.640": 10,
        "iva responsable inscripto - agente de percepción": 11,
        "pequeño contribuyente eventual": 12,
        "monotributista social": 13,
        "pequeño contribuyente eventual social": 14,
        "iva no alcanzado": 15
    }

    for index, fila in datos.iterrows():
        tipo_fila = _get_text(fila, column_lookup, "Tipo Dato", "Tipo").lower()

        if not tipo_fila:
            continue

        if tipo_fila == "comprobante":
            escribir_log(f"{obtener_timestamp()} - Procesando comprobante en la fila {index}.")

            if comprobante_actual:
                comprobante_actual["importe_total"] = total_importe
                comprobante_actual["ImpIVA"] = total_iva
                comprobante_actual["ImpNeto"] = total_neto
                solicitudes.append(comprobante_actual)

            total_importe = _get_float(fila, column_lookup, "Total")
            total_iva = _get_float(fila, column_lookup, "Importe IVA")
            total_neto = _get_float(fila, column_lookup, "Total")

            concepto = _get_int(fila, column_lookup, "Concepto", default=1)

            fecha_emision = _safe_date(_get_cell_value(fila, column_lookup, "Fecha"))

            # Acepta ambas variantes de columnas
            fecha_servicio_desde = _safe_date(
                _get_cell_value(
                    fila,
                    column_lookup,
                    "Periodo Desde",
                    "Fecha servicio desde"
                )
            )

            fecha_servicio_hasta = _safe_date(
                _get_cell_value(
                    fila,
                    column_lookup,
                    "Periodo Hasta",
                    "Fecha servicio hasta"
                )
            )

            fecha_vencimiento_pago_excel = _safe_date(
                _get_cell_value(
                    fila,
                    column_lookup,
                    "Fecha vencimiento pago"
                )
            )

            # Normalización por concepto
            if concepto == 1:
                # Productos: no se informan fechas de servicio ni vencimiento
                fecha_servicio_desde = None
                fecha_servicio_hasta = None
                fecha_vencimiento_pago = None

            elif concepto in (2, 3):
                # Servicios / Productos y Servicios:
                # si faltan las fechas, usar fecha de emisión para evitar rechazo
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

            tipo_doc_texto = _get_text(fila, column_lookup, "Tipo Doc")
            tipo_documento = int(doc_tipo.get(tipo_doc_texto.lower(), 99))

            comprobante_actual = {
                "Concepto": concepto,
                "tipo_comprobante": tipo_comprobante,
                "punto_venta": punto_venta,
                "fecha_emision": fecha_emision,
                "fecha_vencimiento_pago": fecha_vencimiento_pago,
                "fecha_servicio_desde": fecha_servicio_desde,
                "fecha_servicio_hasta": fecha_servicio_hasta,
                "tipo_documento": tipo_documento,
                "numero_documento": _get_int(fila, column_lookup, "Documento", default=0),
                "ID_Moneda": "PES",
                "moneda_cotizacion": 1,
                "importe_total": 0.0,
                "ImpIVA": 0.0,
                "ImpTotConc": _get_float(fila, column_lookup, "Importe Bonificación"),
                "ImpNeto": 0.0,
                "ImpOpEx": _get_float(fila, column_lookup, "Importe Op Ex"),
                "ImpTrib": _get_float(fila, column_lookup, "Importe Tributos"),
                "importe_bonificacion": _get_float(fila, column_lookup, "Importe Bonificación"),
                "codigo_OTributos": _get_cell_value(fila, column_lookup, "Codigo Otros Tributos"),
                "descripcion_OTributos": _get_cell_value(
                    fila,
                    column_lookup,
                    "Descripión Otros Tributos",
                    "Descripción Otros Tributos"
                ),
                "base_imponible_OTributos": _get_float(fila, column_lookup, "Base Imponible Otros Tributos"),
                "Iva": []
            }

        elif tipo_fila == "item":
            escribir_log(f"{obtener_timestamp()} - Procesando ítem en la fila {index}.")

            if comprobante_actual is None:
                continue

            importe_item = _get_float(fila, column_lookup, "Total")
            iva_item = _get_float(fila, column_lookup, "Importe IVA")

            total_importe += importe_item
            total_iva += iva_item
            total_neto += importe_item

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
        comprobante_actual["importe_total"] = total_importe
        comprobante_actual["ImpIVA"] = total_iva
        comprobante_actual["ImpNeto"] = total_neto
        solicitudes.append(comprobante_actual)

    escribir_log(f"{obtener_timestamp()} - Procesamiento de datos FEV1 finalizado. Total de solicitudes generadas: {len(solicitudes)}.")
    return solicitudes