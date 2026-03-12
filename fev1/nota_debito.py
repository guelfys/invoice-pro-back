import pandas as pd
from backend.config import validar_y_transformar_fecha
from backend.log import obtener_timestamp
from solicitudesAFIP.solicitar_ultimo_comprobante import solicitar_ultimo_comprobante


def _pick_value(fila, *columnas):
    for col in columnas:
        if col in fila.index:
            value = fila.get(col)
            if pd.notna(value) and str(value).strip() != "":
                return value
    return None


def _safe_date(value):
    if value is None:
        return None
    return validar_y_transformar_fecha(value, "FEV1")


def procesar_datos_nota_debito_fev1(datos, tipo_comprobante, punto_venta):
    print(f"{obtener_timestamp()} - Inicia con el procesamiento de datos para Nota de Debito {tipo_comprobante}")
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

    solicitudes = []
    for index, fila in datos.iterrows():
        tipo_fila = str(fila["Tipo"]).strip().lower()
        print(f"{obtener_timestamp()} - Valida si la celda tiene datos: {tipo_fila}")

        if tipo_fila == "comprobante":
            fecha_emision_raw = _pick_value(fila, "Fecha")
            periodo_desde_raw = _pick_value(fila, "Periodo Desde", "Fecha servicio desde")
            periodo_hasta_raw = _pick_value(fila, "Periodo Hasta", "Fecha servicio hasta")
            fecha_vto_raw = _pick_value(
                fila,
                "Fecha vencimiento pago",
                "Periodo Hasta",
                "Fecha servicio hasta",
                "Fecha"
            )

            solicitud = {
                "tipo_comprobante": tipo_comprobante,
                "punto_venta": punto_venta,
                "fecha_emision": _safe_date(fecha_emision_raw),
                "fecha_servicio_desde": _safe_date(periodo_desde_raw),
                "fecha_servicio_hasta": _safe_date(periodo_hasta_raw),
                "fecha_vencimiento_pago": _safe_date(fecha_vto_raw),
                "tipo_documento": int(doc_tipo.get(fila["Tipo Doc"], 99)),
                "numero_documento": fila["Documento"],
                "importe_total": fila["Importe Total"],
                "codigo_moneda": "PES",
                "cotizacion_moneda": 1,
                "concepto": fila["Concepto"],
                "motivo_nota_debito": fila.get("Motivo Nota"),
                "comprobante_original_tipo": tipo_comprobante,
                "comprobante_original_punto_venta": punto_venta,
                "comprobante_original_numero": fila.get("Número Comprobante Original")
            }
            solicitudes.append(solicitud)

    return solicitudes


def armar_cuerpo_nota_debito_fev1(solicitudes, client, auth, servicio, numero_actividad=None):
    """
    Arma el cuerpo de la solicitud para las notas de débito a enviar al servicio FEV1.
    """
    cuerpos_solicitud = []

    ultimo_comprobante = solicitar_ultimo_comprobante(client, auth, servicio)
    numero_comprobante = ultimo_comprobante + 1

    for solicitud in solicitudes:
        imp_neto = solicitud["importe_total"]
        imp_iva = solicitud.get("importe_iva", 0)
        imp_trib = solicitud.get("importe_tributos", 0)

        cuerpo = {
            "FeCabReq": {
                "CantReg": 1,
                "PtoVta": solicitud["punto_venta"],
                "CbteTipo": int(solicitud["tipo_comprobante"])
            },
            "FeDetReq": {
                "FECAEDetRequest": [
                    {
                        "Concepto": solicitud["concepto"],
                        "DocTipo": solicitud["tipo_documento"],
                        "DocNro": int(solicitud["numero_documento"]),
                        "CbteFch": solicitud["fecha_emision"],
                        "FchServDesde": solicitud["fecha_servicio_desde"],
                        "FchServHasta": solicitud["fecha_servicio_hasta"],
                        "FchVtoPago": solicitud["fecha_vencimiento_pago"],
                        "ImpTotal": solicitud["importe_total"],
                        "ImpNeto": imp_neto,
                        "ImpOpEx": solicitud.get("importe_exento", 0),
                        "ImpTrib": imp_trib,
                        "ImpIVA": imp_iva,
                        "MonId": solicitud["codigo_moneda"],
                        "MonCotiz": solicitud["cotizacion_moneda"],
                        "CbteDesde": numero_comprobante,
                        "CbteHasta": numero_comprobante,
                        "ImpTotConc": solicitud.get("importe_no_gravado", 0),
                        "CbtesAsoc": {
                            "CbteAsoc": [
                                {
                                    "Tipo": solicitud["comprobante_original_tipo"],
                                    "PtoVta": solicitud["comprobante_original_punto_venta"],
                                    "Nro": solicitud["comprobante_original_numero"]
                                }
                            ]
                        }
                    }
                ]
            }
        }

        det = cuerpo["FeDetReq"]["FECAEDetRequest"][0]

        if int(solicitud.get("concepto", 0)) == 1:
            det.pop("FchServDesde", None)
            det.pop("FchServHasta", None)
            det.pop("FchVtoPago", None)

        if imp_iva > 0:
            det["Iva"] = {
                "AlicIva": [{
                    "Id": 5,
                    "BaseImp": imp_neto,
                    "Importe": imp_iva
                }]
            }

        if imp_trib > 0:
            det["Tributos"] = {
                "Tributo": [{
                    "Id": solicitud.get("codigo_OTributos", 99),
                    "Desc": solicitud.get("descripcion_OTributos", "Otros Tributos"),
                    "BaseImp": solicitud.get("base_imponible_OTributos", imp_neto),
                    "Importe": imp_trib
                }]
            }

        if numero_actividad and numero_actividad != 0:
            det["Actividades"] = {
                "Actividad": [
                    {
                        "Id": int(numero_actividad)
                    }
                ]
            }

        cuerpos_solicitud.append(cuerpo)
        numero_comprobante += 1

    return cuerpos_solicitud