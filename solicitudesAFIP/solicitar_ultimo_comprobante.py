from backend.log import escribir_log, obtener_timestamp

#? Solicitar último comprobante autorizado
def solicitar_ultimo_comprobante(client, auth, servicio):
    try:
        if servicio == "FEV1":
            # Construir el objeto Auth
            auth_request = {
                "Token": auth["Token"],
                "Sign": auth["Sign"],
                "Cuit": auth["Cuit"]
            }
            response = client.service.FECompUltimoAutorizado(

                Auth=auth_request,
                PtoVta=auth['PtoVta'],
                CbteTipo=auth['CbteTipo']
            )
            
            cbte_nro = response.CbteNro

        elif servicio =="MTXCA":
            # Construir el objeto Auth
            auth_request = {
                "token": auth["Token"],
                "sign": auth["Sign"],
                "cuitRepresentada": auth["Cuit"]
            }

            # Construir el objeto ConsultaUltimoComprobanteAutorizadoRequest
            consulta_ultimo_comprobante_request = {
                "numeroPuntoVenta": auth['PtoVta'],
                "codigoTipoComprobante": auth['CbteTipo']
            }

            # Llamada al servicio con los parámetros correctos
            response = client.service.consultarUltimoComprobanteAutorizado(
                authRequest=auth_request,
                consultaUltimoComprobanteAutorizadoRequest=consulta_ultimo_comprobante_request
            )
        
            cbte_nro = response.numeroComprobante

        escribir_log(f"{obtener_timestamp()} - Logró obtener el ultimo comprobante autorizado, {cbte_nro}")
        return cbte_nro if cbte_nro else 0
    except Exception as e:
        escribir_log(f"{obtener_timestamp()} - Error al solicitar el último comprobante autorizado: {e}")
        return 0
