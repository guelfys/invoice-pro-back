from backend.log import escribir_log, obtener_timestamp
import pandas as pd
from backend.config import validar_y_transformar_fecha
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
    return validar_y_transformar_fecha(value, "MTXCA")


def procesar_datos_nota_debito_mtxca(datos, tipo_comprobante, tipo_comprobante_asociado, punto_venta):
    escribir_log(f"{obtener_timestamp()} - Iniciando el procesamiento de datos para MTXCA.")

    if not isinstance(datos, pd.DataFrame):
        escribir_log(f"{obtener_timestamp()} - Error: Se esperaba un DataFrame, pero se recibió un objeto de tipo {type(datos)}")
        return []

    solicitudes = []
    comprobante_actual = None
    items = []
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
        tipo_fila = str(fila["Tipo"]).strip().lower()

        if tipo_fila == "comprobante":
            contador_producto = 1

            if comprobante_actual:
                comprobante_actual["items"] = items
                solicitudes.append(comprobante_actual)

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

            comprobante_actual = {
                "tipo_comprobante": tipo_comprobante,
                "punto_venta": punto_venta,
                "fecha_emision": _safe_date(fecha_emision_raw),
                "tipo_documento": int(doc_tipo.get(fila["Tipo Doc"], 99)),
                "numero_documento": fila["Documento"],
                "importe_total": float(fila.iloc[17] if pd.notna(fila.iloc[17]) else 0.0),
                "importe_gravado": 0.0,
                "importe_no_gravado": 0.0,
                "importe_exento": 0.0,
                "importe_otros_tributos": fila["Importe Otros Tributos"],
                "items": [],
                "codigo_OTributos": fila["Codigo Otros Tributos"],
                "descripcion_OTributos": fila["Descripcion Otros Tributos"],
                "base_imponible_OTributos": fila["Base Imponible Otros Tributos"],
                "concepto": fila["Concepto"],
                "motivo_nota": fila.get("Motivo Nota"),
                "comprobante_original_tipo": tipo_comprobante_asociado,
                "comprobante_original_punto_venta": fila.get("Punto Venta C. Original"),
                "comprobante_original_numero": fila.get("Número Comprobante Original"),
                "iva_receptor": int(iva_receptor.get(str(fila.iloc[3]).lower(), 5)),
                "fecha_servicio_desde": _safe_date(periodo_desde_raw),
                "fecha_servicio_hasta": _safe_date(periodo_hasta_raw),
                "fecha_vencimiento_pago": _safe_date(fecha_vto_raw)
            }

            items = []

            print(f"{obtener_timestamp()} - contador del producto: {contador_producto}")
            items.append({
                "unidadMtx": int(fila["unidadMtx"]),
                "codigoMtx": str(fila["codigoMtx"]),
                "codigo": int(contador_producto),
                "descripcion": str(fila["Descripcion"] if pd.notna(fila["Descripcion"]) else ''),
                "cantidad": int(fila["Cantidad"] if pd.notna(fila["Cantidad"]) else 0),
                "codigoUnidadMedida": int(6),
                "precioUnitario": float(fila["Precio Unitario"] if pd.notna(fila["Precio Unitario"]) else 0),
                "importeItem": float(fila["Importe Total"] if pd.notna(fila["Importe Total"]) else 0),
                "codigoCondicionIVA": str(fila["Codigo Condicion IVA"]),
                "importeIVA": float(fila["Importe IVA"] if pd.notna(fila["Importe IVA"]) else 0)
            })
            contador_producto += 1

        elif tipo_fila == "item":
            print(f"{obtener_timestamp()} - contador del producto: {contador_producto}")
            items.append({
                "unidadMtx": int(fila["unidadMtx"]),
                "codigoMtx": str(fila["codigoMtx"]),
                "codigo": int(contador_producto),
                "descripcion": str(fila["Descripcion"] if pd.notna(fila["Descripcion"]) else ''),
                "cantidad": int(fila["Cantidad"] if pd.notna(fila["Cantidad"]) else 0),
                "codigoUnidadMedida": int(6),
                "precioUnitario": float(fila["Precio Unitario"] if pd.notna(fila["Precio Unitario"]) else 0),
                "importeItem": float(fila["Importe Total"] if pd.notna(fila["Importe Total"]) else 0),
                "codigoCondicionIVA": str(fila["Codigo Condicion IVA"]),
                "importeIVA": float(fila["Importe IVA"] if pd.notna(fila["Importe IVA"]) else 0)
            })
            contador_producto += 1

    if comprobante_actual:
        comprobante_actual["items"] = items
        solicitudes.append(comprobante_actual)

    escribir_log(f"{obtener_timestamp()} - Procesamiento de datos MTXCA finalizado. Total de solicitudes generadas: {len(solicitudes)}.")
    return solicitudes


def armar_cuerpo_nota_debito_mtxca(solicitudes, client, auth, servicio, numero_actividad=None):
    cuerpos_solicitud = []

    ultimo_comprobante = solicitar_ultimo_comprobante(client, auth, servicio)
    numero_comprobante = ultimo_comprobante + 1

    for solicitud in solicitudes:
        importe_no_gravado = sum(item['importeItem'] for item in solicitud['items'] if item['codigoCondicionIVA'] == "1")
        importe_exento = sum(item['importeItem'] for item in solicitud['items'] if item['codigoCondicionIVA'] == "2")
        importe_gravado = sum(item['importeItem'] for item in solicitud['items'] if item['codigoCondicionIVA'] in ["3", "4", "5", "6", "8", "9"])
        importe_iva = sum(item['importeIVA'] for item in solicitud['items'] if item.get("codigoCondicionIVA") in ["4", "5", "6", "8", "9"])

        importe_subtotal = importe_gravado + importe_no_gravado + importe_exento
        importe_otros_tributos = solicitud.get("importe_otros_tributos", 0.0)

        if importe_subtotal > 0:
            importe_total = importe_subtotal + importe_iva + importe_otros_tributos
        else:
            importe_total = solicitud["importe_total"]

        if solicitud["concepto"] == 2 or solicitud["concepto"] == "2" or solicitud["concepto"] == 3 or solicitud["concepto"] == "3":
            cuerpo = {
                "codigoConcepto": solicitud["concepto"],
                "codigoTipoComprobante": solicitud["tipo_comprobante"],
                "numeroPuntoVenta": solicitud["punto_venta"],
                "numeroComprobante": numero_comprobante,
                "fechaEmision": solicitud["fecha_emision"],
                "fechaServicioDesde": solicitud["fecha_servicio_desde"],
                "fechaServicioHasta": solicitud["fecha_servicio_hasta"],
                "fechaVencimientoPago": solicitud["fecha_vencimiento_pago"],
                "codigoTipoDocumento": solicitud["tipo_documento"],
                "numeroDocumento": solicitud["numero_documento"],
                "importeSubtotal": importe_subtotal,
                "importeTotal": importe_total,
                "codigoMoneda": "PES",
                "cotizacionMoneda": 1,
                "condicionIVAReceptor": solicitud["iva_receptor"]
            }
        else:
            cuerpo = {
                "codigoConcepto": solicitud["concepto"],
                "codigoTipoComprobante": solicitud["tipo_comprobante"],
                "numeroPuntoVenta": solicitud["punto_venta"],
                "numeroComprobante": solicitud["comprobante_original_numero"],
                "fechaEmision": solicitud["fecha_emision"],
                "codigoTipoDocumento": solicitud["tipo_documento"],
                "numeroDocumento": solicitud["numero_documento"],
                "importeSubtotal": importe_subtotal,
                "importeTotal": importe_total,
                "codigoMoneda": "PES",
                "cotizacionMoneda": 1,
                "condicionIVAReceptor": solicitud["iva_receptor"]
            }

        items = []
        array_items = []
        for idx, item in enumerate(solicitud["items"]):
            if item.get("importeIVA") > 0:
                item_data = {
                    "unidadesMtx": item["unidadMtx"],
                    "codigoMtx": item["codigoMtx"],
                    "codigo": idx + 1,
                    "descripcion": item["descripcion"],
                    "cantidad": item["cantidad"],
                    "codigoUnidadMedida": 6,
                    "precioUnitario": item["precioUnitario"],
                    "importeItem": item["importeItem"] + item["importeIVA"],
                    "importeBonificacion": solicitud.get("importe_bonificacion", 0.0),
                    "codigoCondicionIVA": item["codigoCondicionIVA"],
                    "importeIVA": item.get("importeIVA", 0.0)
                }
            else:
                item_data = {
                    "unidadesMtx": item["unidadMtx"],
                    "codigoMtx": item["codigoMtx"],
                    "codigo": idx + 1,
                    "descripcion": item["descripcion"],
                    "cantidad": item["cantidad"],
                    "codigoUnidadMedida": 6,
                    "precioUnitario": item["precioUnitario"],
                    "importeItem": item["importeItem"],
                    "importeBonificacion": solicitud.get("importe_bonificacion", 0.0),
                    "codigoCondicionIVA": item["codigoCondicionIVA"]
                }
            items.append(item_data)

        array_items.append({"item": items})
        cuerpo["arrayItems"] = array_items

        items_iva = [item for item in solicitud['items'] if item.get("codigoCondicionIVA") in ['4', '5', '6', '8', '9']]

        if importe_gravado >= 0:
            cuerpo["importeGravado"] = importe_gravado
        elif importe_exento > 0:
            cuerpo["importeExento"] = importe_exento

        if importe_iva > 0:
            cuerpo["importeGravado"] = importe_gravado
            cuerpo["importeNoGravado"] = importe_no_gravado
            cuerpo["importeExento"] = importe_exento
            cuerpo["arraySubtotalesIVA"] = {
                "subtotalIVA": [{
                    "codigo": item["codigoCondicionIVA"],
                    "importe": item.get("importeIVA")
                } for item in items_iva]
            }

        if importe_otros_tributos > 0:
            cuerpo["importeOtrosTributos"] = importe_otros_tributos
            cuerpo["arrayOtrosTributos"] = [{"tributo": {
                "codigo": solicitud["codigo_OTributos"],
                "descripcion": solicitud["descripcion_OTributos"],
                "baseImponible": solicitud["base_imponible_OTributos"],
                "importe": importe_otros_tributos
            }}]

        cuerpo["arrayComprobantesAsociados"] = {
            "comprobanteAsociado": [
                {
                    "codigoTipoComprobante": solicitud["comprobante_original_tipo"],
                    "numeroPuntoVenta": solicitud["comprobante_original_punto_venta"],
                    "numeroComprobante": solicitud["comprobante_original_numero"]
                }
            ]
        }

        if numero_actividad and numero_actividad != 0:
            cuerpo["arrayActividades"] = {
                "actividad": [
                    {
                        "codigo": int(numero_actividad)
                    }
                ]
            }

        cuerpos_solicitud.append(cuerpo)
        numero_comprobante += 1

    return cuerpos_solicitud
