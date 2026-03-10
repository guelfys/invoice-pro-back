
import pandas as pd
from backend.config import validar_y_transformar_fecha
from backend.log import obtener_timestamp
from solicitudesAFIP.solicitar_ultimo_comprobante import solicitar_ultimo_comprobante


def procesar_datos_nota_debito_fev1(datos, tipo_comprobante, punto_venta):
    """
    Procesa los datos de la nota de débito desde un DataFrame y retorna una lista de solicitudes.
    """
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

        tipo_fila = fila["Tipo"].lower()
        print(f"{obtener_timestamp()} - Valida si la celda tiene datos: {tipo_fila}")

        if tipo_fila == "comprobante":

            solicitud = {
                "tipo_comprobante": tipo_comprobante,  # Código de nota de débito (ej: 2 para Nota de Débito A)
                "punto_venta": punto_venta,
                "fecha_emision": validar_y_transformar_fecha(fila["Fecha"], "FEV1"),
                "fecha_servicio_desde": validar_y_transformar_fecha(fila["Fecha servicio desde"], "FEV1"),
                "fecha_servicio_hasta": validar_y_transformar_fecha(fila["Fecha servicio hasta"], "FEV1"),
                "fecha_vencimiento_pago": validar_y_transformar_fecha(fila["Fecha vencimiento pago"], "FEV1"),
                "tipo_documento": int(doc_tipo.get(fila["Tipo Doc"], 99)),
                "numero_documento": fila["Documento"],
                "importe_total": fila["Importe Total"],
                "codigo_moneda": "PES",
                "cotizacion_moneda": 1,
                "concepto": fila["Concepto"],  # 1 para productos, 2 para servicios
                "motivo_nota_debito": fila.get("Motivo Nota"),
                # Agregar referencia al comprobante original
                "comprobante_original_tipo": tipo_comprobante,  # Tipo del comprobante original
                "comprobante_original_punto_venta": punto_venta,  # Punto de venta del comprobante original
                "comprobante_original_numero": fila.get("Número Comprobante Original")  # Número del comprobante original
            }
            solicitudes.append(solicitud)
    
    return solicitudes


def armar_cuerpo_nota_debito_fev1(solicitudes, client, auth, servicio, numero_actividad=None):
    """
    Arma el cuerpo de la solicitud para las notas de débito a enviar al servicio FEV1.
    """
    cuerpos_solicitud = []
    
    # Solicitar el último comprobante autorizado
    ultimo_comprobante = solicitar_ultimo_comprobante(client, auth, servicio)
    numero_comprobante = ultimo_comprobante + 1

    for solicitud in solicitudes:
        imp_neto = solicitud["importe_total"] # Ojo: Asegurate que para A/B esto no deba ser distinto al total
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

        # 1. Corrección del Concepto (Igual que en armado_cuerpo.py)
        if int(solicitud.get("concepto", 0)) == 1:
            det.pop("FchServDesde", None)
            det.pop("FchServHasta", None)
            det.pop("FchVtoPago", None)

        # 2. Agregar Array de IVA si corresponde
        if imp_iva > 0:
            det["Iva"] = {
                "AlicIva": [{
                    "Id": 5, # Cuidado: el 5 es 21%. Deberías mapearlo dinámicamente si hay varias alícuotas
                    "BaseImp": imp_neto,
                    "Importe": imp_iva
                }]
            }

        # 3. Agregar Array de Tributos si corresponde
        if imp_trib > 0:
            det["Tributos"] = {
                "Tributo": [{
                    "Id": solicitud.get("codigo_OTributos", 99), 
                    "Desc": solicitud.get("descripcion_OTributos", "Otros Tributos"),
                    "BaseImp": solicitud.get("base_imponible_OTributos", imp_neto),
                    "Importe": imp_trib
                }]
            }

        # 4. Agregar Actividad si fue enviada
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