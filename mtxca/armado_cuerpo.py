
from solicitudesAFIP.solicitar_ultimo_comprobante import solicitar_ultimo_comprobante


def armar_cuerpo_solicitud_mtxca_A(solicitudes, client, auth, servicio, numero_actividad):
    cuerpos_solicitud = []

    ultimo_comprobante = solicitar_ultimo_comprobante(client, auth, servicio)
    numero_comprobante = ultimo_comprobante + 1

    for solicitud in solicitudes:
        # Calcular importes
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

        # Construir cuerpo
        cuerpo = {
            "codigoConcepto": solicitud["codigoConcepto"],
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
            "condicionIVAReceptor": solicitud["iva_receptor"],
                    }

        items = []
        array_items = []
        for idx, item in enumerate(solicitud["items"]):
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
            items.append(item_data)
        
        array_items.append({"item": items})

        cuerpo["arrayItems"] = array_items

        # Filtrar los ítems que tienen los códigos de IVA esperados
        items_iva = [item for item in solicitud['items'] if item.get("codigoCondicionIVA") in ['4', '5', '6', '8', '9']]

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

        elif importe_gravado > 0:
            cuerpo["importeGravado"] = importe_gravado          

        elif importe_exento > 0:
            cuerpo["importeExento"] = importe_exento

        if importe_otros_tributos > 0:
            cuerpo["importeOtrosTributos"] = importe_otros_tributos
            cuerpo["arrayOtrosTributos"] = [{"tributo": {
                "codigo": solicitud["codigo_OTributos"],
                "descripcion": solicitud["descripcion_OTributos"],
                "baseImponible": solicitud["base_imponible_OTributos"],
                "importe": importe_otros_tributos
            }}]

        if numero_actividad and numero_actividad != 0:
            cuerpo["codigoActividad"] = numero_actividad

        cuerpos_solicitud.append(cuerpo)
        numero_comprobante += 1

    return cuerpos_solicitud

def armar_cuerpo_solicitud_mtxca_B(solicitudes, client, auth, servicio, numero_actividad=None): 
    cuerpos_solicitud = []

    ultimo_comprobante = solicitar_ultimo_comprobante(client, auth, servicio)
    numero_comprobante = ultimo_comprobante + 1

    for solicitud in solicitudes:

        # Calcular importes
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

        print("")
        print("------------------- SUMATORIA IMPORTES ------------------")
        print(f"Importe Gravado: {importe_gravado}")
        print(f"Importe No Gravado: {importe_no_gravado}")
        print(f"Importe Exento: {importe_exento}")
        print(f"Importe Sub Total: {importe_subtotal}")
        print(f"Importe Otros Tributos: {importe_otros_tributos}")
        print(f"Importe IVA: {importe_iva}")
        print(f"Importe Total: {importe_total}")
        print("---------------------------------------------------------")
        print("")

        cuerpo = {
            "codigoConcepto": solicitud["codigoConcepto"],
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
            "condicionIVAReceptor": solicitud["iva_receptor"],
            "arrayItems": [{"item": {
                "unidadesMtx": item["unidadMtx"],
                "codigoMtx": item["codigoMtx"],
                "codigo": idx + 1,
                "descripcion": item["descripcion"],
                "cantidad": item["cantidad"],
                "codigoUnidadMedida": 6,  # Unidades, código AFIP
                "precioUnitario": item["precioUnitario"],
                "importeItem": item["importeItem"],        # SI ES FACTURA CON IVA = SE SUMA "+ item["importeIVA"]"
                "importeBonificacion": solicitud["importe_bonificacion"],
                "codigoCondicionIVA": item["codigoCondicionIVA"]                      
                # SI HAY IVA MAYOR A 0, SE AGREGA ESTE PARAMETRO <"importeIVA": item["importeIVA"]>
            }} for idx, item in enumerate(solicitud["items"])]
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

        # Filtrar los ítems que tienen los códigos de IVA esperados
        items_iva = [item for item in solicitud['items'] if item.get("codigoCondicionIVA") in ['4', '5', '6', '8', '9']]

        # Si el importe de IVA es mayor a 0 y hay al menos un ítem con los códigos de IVA esperados
        if importe_iva > 0 and items_iva:
            cuerpo["arraySubtotalesIVA"] = {
                "subtotalIVA": [{
                    "codigo": item["codigoCondicionIVA"],
                    "importe": item.get("importeIVA")
                } for item in items_iva]
            }

            cuerpo["importeGravado"] = importe_gravado
        
        if importe_gravado > 0:
            cuerpo["importeGravado"] = importe_gravado

        if importe_no_gravado > 0:
            cuerpo["importeNoGravado"] = importe_no_gravado
        
        if importe_exento > 0:
            cuerpo["importeExento"] = importe_exento

        if importe_otros_tributos > 0:
            cuerpo["importeOtrosTributos"] = importe_otros_tributos

            cuerpo["arrayOtrosTributos"] = [{"tributo": {
                "codigo": solicitud["codigo_OTributos"],  # Ejemplo de código de tributo, ajusta según sea necesario
                "descripcion": solicitud["descripcion_OTributos"],
                "baseImponible": solicitud["base_imponible_OTributos"],  # Ajusta según sea necesario
                "importe": importe_otros_tributos
            }}]

        # Corrección para MTXCA: Estructura de array de actividades
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
