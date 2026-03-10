from backend.log import escribir_log, obtener_timestamp
from solicitudesAFIP.solicitar_ultimo_comprobante import solicitar_ultimo_comprobante

def _safe_float(v):
    if v is None:
        return 0.0
    if isinstance(v, str) and v.strip() == "":
        return 0.0
    try:
        return float(v)
    except Exception:
        return 0.0

def armar_cuerpo_solicitud_fev1(solicitudes, client, auth, servicio, numero_actividad=None):
    cuerpos_solicitud = []

    # Solicitar el último comprobante autorizado
    ultimo_comprobante = solicitar_ultimo_comprobante(client, auth, servicio)
    numero_comprobante = ultimo_comprobante + 1

    for solicitud in solicitudes:

        imp_totconc = _safe_float(solicitud.get("ImpTotConc", 0.0))
        imp_neto    = _safe_float(solicitud.get("ImpNeto", 0.0))
        imp_opex    = _safe_float(solicitud.get("ImpOpEx", 0.0))
        imp_iva     = _safe_float(solicitud.get("ImpIVA", 0.0))
        imp_trib    = _safe_float(solicitud.get("ImpTrib", 0.0))

        imp_total_calc = round(imp_totconc + imp_neto + imp_opex + imp_iva + imp_trib, 2)

        importe_total_excel = _safe_float(solicitud.get("importe_total", 0.0))
        if imp_total_calc == 0.0 and importe_total_excel > 0.0:
            imp_neto = importe_total_excel
            imp_total_calc = round(imp_totconc + imp_neto + imp_opex + imp_iva + imp_trib, 2)

        cuerpo = {
            "FeCabReq": {
                "CantReg": 1,
                "PtoVta": solicitud["punto_venta"],
                "CbteTipo": solicitud["tipo_comprobante"]
            },
            "FeDetReq": {
                "FECAEDetRequest": [{
                    "Concepto": solicitud["Concepto"],
                    "CbteDesde": int(numero_comprobante),
                    "CbteHasta": int(numero_comprobante),
                    "CbteFch": solicitud["fecha_emision"],
                    "DocTipo": solicitud["tipo_documento"],
                    "DocNro": solicitud["numero_documento"],
                    "ImpTotal": imp_total_calc,
                    "ImpTotConc": imp_totconc,
                    "ImpNeto": imp_neto,
                    "ImpOpEx": imp_opex,
                    "ImpIVA": imp_iva,
                    "ImpTrib": imp_trib,
                    "FchServDesde": solicitud.get("fecha_servicio_desde", None),
                    "FchServHasta": solicitud.get("fecha_servicio_hasta", None),
                    "FchVtoPago": solicitud.get("fecha_vencimiento_pago", None),
                    "MonId": "PES",
                    "MonCotiz": 1,
                    "Iva": solicitud.get("Iva", [])
                }]
            }
        }

        det = cuerpo["FeDetReq"]["FECAEDetRequest"][0]

        if int(solicitud.get("Concepto", 0)) == 1:
            det.pop("FchServDesde", None)
            det.pop("FchServHasta", None)
            det.pop("FchVtoPago", None)

        if imp_trib > 0:
            cuerpo["Tributos"] = {
                "Tributo": [{
                    "Id": solicitud["codigo_OTributos"],
                    "Desc": solicitud["descripcion_OTributos"],
                    "BaseImp": solicitud["base_imponible_OTributos"],
                    "Importe": imp_trib
                }]
            }


        if imp_iva > 0:
            cuerpo["FeDetReq"]["FECAEDetRequest"][0]["Iva"] = {
                "AlicIva": [{
                    "Id": 5,
                    "BaseImp": imp_neto,
                    "Importe": imp_iva
                }]
            }
        else:
            # Asegura que no existe el campo "Iva" si no hay IVA
            if "Iva" in cuerpo["FeDetReq"]["FECAEDetRequest"][0]:
                del cuerpo["FeDetReq"]["FECAEDetRequest"][0]["Iva"]

        # Agregar el número de actividad si está presente
        if numero_actividad and numero_actividad != 0:
            cuerpo["FeDetReq"]["FECAEDetRequest"][0]["Actividades"] = {
                "Actividad": [
                    {
                        "Id": int(numero_actividad)
                    }
                ]
            }

        cuerpos_solicitud.append(cuerpo)
        numero_comprobante += 1

    escribir_log(f"{obtener_timestamp()} - Cuerpos de solicitud armados para FEV1.")
    return cuerpos_solicitud
