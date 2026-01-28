
import pandas as pd
from backend.config import validar_y_transformar_fecha
from backend.log import obtener_timestamp
from solicitudesAFIP.solicitar_ultimo_comprobante import solicitar_ultimo_comprobante


def procesar_datos_nota_credito_fev1(datos, tipo_comprobante, punto_venta):
    
    print(f"{obtener_timestamp()} - Inicia con el procesamiento de datos para Nota de Credito {tipo_comprobante}")
    """
    Procesa los datos de la nota de crédito desde un DataFrame y retorna una lista de solicitudes.
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
            print(f"{obtener_timestamp()} - Encontró para armar una nota de credito")

            solicitud = {
                "tipo_comprobante": tipo_comprobante,  # Código de nota de crédito (ej: 3 para Nota de Crédito A)
                "punto_venta": punto_venta,
                "fecha_emision": validar_y_transformar_fecha(fila["Fecha"], "FEV1"),
                "fecha_servicio_desde": validar_y_transformar_fecha(fila["Fecha servicio desde"], "FEV1"),
                "fecha_servicio_hasta": validar_y_transformar_fecha(fila["Fecha servicio hasta"], "FEV1"),
                "fecha_vencimiento_pago": validar_y_transformar_fecha(fila["Fecha vencimiento pago"], "FEV1"),
                "tipo_documento": int(doc_tipo.get(fila["Tipo Doc"], 99)),
                "numero_documento": fila["Documento"],
                "importe_total": fila["Importe Total"],
                "codigo_moneda": "PES",  # Se podría hacer más dinámico según datos de la moneda
                "cotizacion_moneda": 1,
                "concepto": fila["Concepto"],  # 1 para productos, 2 para servicios
                "motivo_nota_credito": fila.get("Motivo Nota"),
                # Agregar referencia al comprobante original
                "comprobante_original_tipo": tipo_comprobante,  # Tipo del comprobante original
                "comprobante_original_punto_venta": punto_venta,  # Punto de venta del comprobante original
                "comprobante_original_numero": fila.get("Número Comprobante Original")  # Número del comprobante original
            }
            solicitudes.append(solicitud)
    
    return solicitudes


def armar_cuerpo_nota_credito_fev1(solicitudes, client, auth, servicio):
    """
    Arma el cuerpo de la solicitud para las notas de crédito a enviar al servicio FEV1.
    """
    cuerpos_solicitud = []
    
    # Solicitar el último comprobante autorizado
    ultimo_comprobante = solicitar_ultimo_comprobante(client, auth, servicio)
    numero_comprobante = ultimo_comprobante + 1

    for solicitud in solicitudes:
        cuerpo = {
            "FeCabReq": {
                    "CantReg": 1,
                    "PtoVta": solicitud["punto_venta"],
                    "CbteTipo": int(solicitud["tipo_comprobante"])  # Código del tipo de comprobante (Nota de Crédito)
                },
                "FeDetReq": {
                    "FECAEDetRequest": [
                        {
                            "Concepto": solicitud["concepto"],  # 1: Productos, 2: Servicios, 3: Productos y Servicios
                            "DocTipo": solicitud["tipo_documento"],  # Tipo de documento del receptor
                            "DocNro": int(solicitud["numero_documento"]),  # Número de documento del receptor
                            "CbteFch": solicitud["fecha_emision"],  # Fecha en formato YYYYMMDD
                            "FchServDesde": solicitud["fecha_servicio_desde"],
                            "FchServHasta": solicitud["fecha_servicio_hasta"],
                            "FchVtoPago": solicitud["fecha_vencimiento_pago"],
                            "ImpTotal": solicitud["importe_total"],  # Importe total de la nota de crédito
                            "ImpNeto": solicitud["importe_total"],  # Importe NETO de la nota de crédito
                            "ImpOpEx": solicitud.get("importe_exento", 0),
                            "ImpTrib": solicitud.get("importe_tributos", 0),
                            "ImpIVA": solicitud.get("importe_iva", 0),
                            "MonId": solicitud["codigo_moneda"],  # Código de moneda
                            "MonCotiz": solicitud["cotizacion_moneda"],  # Cotización de la moneda
                            # Asociar con el comprobante original
                            "CbteDesde": numero_comprobante,  # Número del comprobante desde
                            "CbteHasta": numero_comprobante,  # Número del comprobante hasta (usualmente igual a `CbteDesde`)
                            "ImpTotConc": solicitud.get("importe_no_gravado", 0),  # Importe no gravado, puede ser 0
                            "CbtesAsoc": {
                                "CbteAsoc": [
                                    {
                                        "Tipo": solicitud["comprobante_original_tipo"],  # Tipo del comprobante original (1: Factura A, 6: Factura B)
                                        "PtoVta": solicitud["comprobante_original_punto_venta"],  # Punto de venta original
                                        "Nro": solicitud["comprobante_original_numero"]  # Número del comprobante original
                                    }
                                ]
                            }
                        }
                    ]
                }
            }

        cuerpos_solicitud.append(cuerpo)
        numero_comprobante += 1

    return cuerpos_solicitud
