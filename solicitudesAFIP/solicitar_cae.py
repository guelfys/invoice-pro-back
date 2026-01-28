from backend.log import escribir_log, obtener_timestamp
from zeep.helpers import serialize_object
from variables import ListaValidacionCAE
from zeep.exceptions import Fault

#! ------------------------------------------------------------------------------------------------------------------------------------------------------
#! 
#! ENVIAR SOLICITUDES CAES
#!
#! ------------------------------------------------------------------------------------------------------------------------------------------------------

# Llamada al web service para obtener el CAE
def enviar_solicitud_cae(servicio, client, auth, Cuerpos_Solicitud):
    for cuerpo in Cuerpos_Solicitud:
        try:
            # Serializar objeto para verificar la estructura
            comprobante_serializado = serialize_object(cuerpo)
            
            if servicio == "FEV1":
                # Envío de la solicitud
                response = client.service.FECAESolicitar(Auth=auth, FeCAEReq=comprobante_serializado)

                # Verificar la respuesta y acceder a los datos
                if hasattr(response, 'FeDetResp') and response.FeDetResp:
                    fe_det_resp = response.FeDetResp
                    if hasattr(fe_det_resp, 'FECAEDetResponse') and len(fe_det_resp.FECAEDetResponse) > 0:
                        primer_elemento = fe_det_resp.FECAEDetResponse[0]
                        
                        escribir_log("")
                        escribir_log("-------------------------------------------------")
                        escribir_log(f"{response}")
                        escribir_log("-------------------------------------------------")
                        escribir_log("")

                        # Obtener el CAE y la fecha de vencimiento del CAE
                        cae = getattr(primer_elemento, 'CAE', None)
                        cae_fch_vto = getattr(primer_elemento, 'CAEFchVto', None)
                        
                        # Obtener el CbteDesde desde el cuerpo
                        if 'FeDetReq' in cuerpo and 'FECAEDetRequest' in cuerpo['FeDetReq']:
                            cbte_desde = cuerpo['FeDetReq']['FECAEDetRequest'][0]['CbteDesde']
                        else:
                            cbte_desde = 'Desconocido'
                        
                        # Verificar y escribir los resultados en el log y en la consola
                        if cae and cae_fch_vto:
                            escribir_log(f"{obtener_timestamp()} - CAE obtenido: {cae} con fecha de vencimiento: {cae_fch_vto} para comprobante {cbte_desde}")
                            print(f"{obtener_timestamp()} - CAE obtenido: {cae} con fecha de vencimiento: {cae_fch_vto}")

                            # Crear la lista de validación para el comprobante actual
                            if not cae:
                                validacion_actual = [False, "-", "-", "-"]
                            else:
                                validacion_actual = [True, cae, cae_fch_vto, cbte_desde]

                            # Agregar la validación actual a la lista de validaciones
                            ListaValidacionCAE.append(validacion_actual)

                        else:
                            escribir_log(f"{obtener_timestamp()} - Error en la respuesta: El primer elemento en FECAEDetResponse no contiene CAE o fecha de vencimiento.")
                            print(f"{obtener_timestamp()} - Error en la respuesta: El primer elemento en FECAEDetResponse no contiene CAE o fecha de vencimiento.")
                            validacion_actual = [False, "-", "-", "-"]
                            ListaValidacionCAE.append(validacion_actual)
                    else:
                        escribir_log(f"{obtener_timestamp()} - Error en la solicitud de CAE, respuesta vacía en FECAEDetResponse.")
                        print(f"{obtener_timestamp()} - Error en la solicitud de CAE, respuesta vacía en FECAEDetResponse.")
                        validacion_actual = [False, "-", "-", "-"]
                        ListaValidacionCAE.append(validacion_actual)
                else:
                    escribir_log(f"{obtener_timestamp()} - Error en la solicitud de CAE, FeDetResp no está presente en la respuesta.")
                    print(f"{obtener_timestamp()} - Error en la solicitud de CAE, FeDetResp no está presente en la respuesta.")
                    validacion_actual = [False, "-", "-", "-"]
                    ListaValidacionCAE.append(validacion_actual)

            if servicio == "MTXCA":
                # Envío de la solicitud
                response = client.service.autorizarComprobante(authRequest=auth, comprobanteCAERequest=comprobante_serializado)

                # Verificar la respuesta y acceder a los datos
                if hasattr(response, 'comprobanteResponse') and response.comprobanteResponse:
                    comprobante_response = response.comprobanteResponse

                    escribir_log("")
                    escribir_log("-------------------------------------------------")
                    escribir_log(f"{response}")
                    escribir_log("-------------------------------------------------")
                    escribir_log("")
                    
                    # Obtener el CAE, la fecha de vencimiento del CAE, y el número de comprobante
                    cae = getattr(comprobante_response, 'CAE', None)
                    fecha_vencimiento_cae = getattr(comprobante_response, 'fechaVencimientoCAE', None)
                    numero_comprobante = getattr(comprobante_response, 'numeroComprobante', None)
                    
                    # Verificar y escribir los resultados en el log y en la consola
                    if cae and fecha_vencimiento_cae:
                        escribir_log(f"{obtener_timestamp()} - CAE obtenido: {cae} con fecha de vencimiento: {fecha_vencimiento_cae} para comprobante {numero_comprobante}")
                        print(f"{obtener_timestamp()} - CAE obtenido: {cae} con fecha de vencimiento: {fecha_vencimiento_cae}")

                        # Crear la lista de validación para el comprobante actual
                        validacion_actual = [True, cae, fecha_vencimiento_cae, numero_comprobante]
                        
                        # Agregar la validación actual a la lista de validaciones
                        ListaValidacionCAE.append(validacion_actual)
                    else:
                        escribir_log(f"{obtener_timestamp()} - Error en la respuesta: La respuesta no contiene CAE o fecha de vencimiento.")
                        print(f"{obtener_timestamp()} - Error en la respuesta: La respuesta no contiene CAE o fecha de vencimiento.")
                        validacion_actual = [False, "-", "-", "-"]
                        ListaValidacionCAE.append(validacion_actual)
                else:
                    escribir_log(f"{obtener_timestamp()} - Error en la solicitud de CAE, comprobanteResponse no está presente en la respuesta.")
                    print(f"{obtener_timestamp()} - Error en la solicitud de CAE, comprobanteResponse no está presente en la respuesta.")
                    validacion_actual = [False, "-", "-", "-"]
                    ListaValidacionCAE.append(validacion_actual)           

            # Imprimir la respuesta completa para depuración
            print(f"{obtener_timestamp()} - Respuesta del servicio: {response}")

        except Fault as e:
            escribir_log(f"{obtener_timestamp()} - Error durante la solicitud de CAE: {e}")
            print(f"{obtener_timestamp()} - Error durante la solicitud de CAE: {e}")
            validacion_actual = [False, "-", "-", "-"]