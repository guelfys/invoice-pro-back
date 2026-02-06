from backend.log import escribir_log, obtener_timestamp
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font
import win32com.client as win32
import os
import shutil
import re
import json
import base64
import datetime as _dt
import tempfile
from io import BytesIO
from facturasOuput.qr_rescue import asegurar_qr_en_factura

try:
    import qrcode  
except Exception:
    qrcode = None

_QR_BASE_URL = "https://www.arca.gob.ar/fe/qr/"

def _solo_digitos(v):
    if v is None:
        return ""
    return re.sub(r"\D+", "", str(v))

def _to_int_safe(v, default=None):
    s = _solo_digitos(v)
    if s == "":
        return default
    try:
        return int(s)
    except Exception:
        return default

def _to_float_safe(v, default=0.0):
    if v is None:
        return default
    if isinstance(v, str):
        s = v.strip()
        if s == "":
            return default
        # soporta "1.234,56" y "1234.56"
        s = s.replace(".", "").replace(",", ".") if ("," in s and "." in s) else s.replace(",", ".")
        try:
            return float(s)
        except Exception:
            return default
    try:
        return float(v)
    except Exception:
        return default

def _set_num(ws, cell_ref, value, fmt="#,##0.00"):

    try:
        if value is None:
            v = 0.0
        elif isinstance(value, str):
            s = value.strip()
            if s == "":
                v = 0.0
            else:
                # soporta "1.234,56" o "1234.56" o "1234,56"
                if "," in s and "." in s:
                    s = s.replace(".", "").replace(",", ".")
                else:
                    s = s.replace(",", ".")
                v = float(s)
        else:
            v = float(value)
    except Exception:
        v = 0.0

    c = ws[cell_ref]
    c.value = v
    c.number_format = fmt


def _fecha_full_date(v):

    if v is None:
        return None
    if isinstance(v, (_dt.date, _dt.datetime)):
        d = v.date() if isinstance(v, _dt.datetime) else v
        return d.strftime("%Y-%m-%d")
    s = str(v).strip()
    if not s:
        return None
    # YYYYMMDD
    if re.fullmatch(r"\d{8}", s):
        return f"{s[0:4]}-{s[4:6]}-{s[6:8]}"
    # YYYY-MM-DD (o con /)
    m = re.fullmatch(r"(\d{4})[-/](\d{1,2})[-/](\d{1,2})", s)
    if m:
        y, mo, d = m.group(1), int(m.group(2)), int(m.group(3))
        return f"{y}-{mo:02d}-{d:02d}"
    # DD/MM/YYYY
    m = re.fullmatch(r"(\d{1,2})[-/](\d{1,2})[-/](\d{4})", s)
    if m:
        d, mo, y = int(m.group(1)), int(m.group(2)), m.group(3)
        return f"{y}-{mo:02d}-{d:02d}"
    # fallback
    try:
        # último intento: parse ISO-like
        dt = _dt.datetime.fromisoformat(s)
        return dt.date().strftime("%Y-%m-%d")
    except Exception:
        return None

def _map_tipo_cmp(tipo_nota, tipo_factura):
    """
    Mapea a códigos AFIP/ARCA:
    - Factura: A=1, B=6, C=11
    - Nota Débito: A=2, B=7, C=12
    - Nota Crédito: A=3, B=8, C=13
    """
    tipo_factura = str(tipo_factura).upper().strip()
    tipo_nota = str(tipo_nota).strip().lower()

    fact = {"A": 1, "B": 6, "C": 11}
    deb = {"A": 2, "B": 7, "C": 12}
    cred = {"A": 3, "B": 8, "C": 13}

    if tipo_nota == "factura":
        return fact.get(tipo_factura)
    if tipo_nota == "debito":
        return deb.get(tipo_factura)
    if tipo_nota == "credito":
        return cred.get(tipo_factura)
    return None

def _guess_tipo_doc_rec_y_nro(doc_nro):
    """
    Heurística simple:
    - 11 dígitos => CUIT (80)
    - 8 dígitos  => DNI (96)
    - otro       => 99
    """
    s = _solo_digitos(doc_nro)
    if not s:
        return (None, None)
    if len(s) == 11:
        return (80, int(s))
    if len(s) == 8:
        return (96, int(s))
    return (99, int(s))

def _build_qr_url(*, fecha, cuit, pto_vta, tipo_cmp, nro_cmp, importe, moneda="PES", cotizacion=1.0,
                  tipo_doc_rec=None, nro_doc_rec=None, tipo_cod_aut="E", cod_aut=None):
    payload = {
        "ver": 1,
        "fecha": fecha,
        "cuit": int(cuit),
        "ptoVta": int(pto_vta),
        "tipoCmp": int(tipo_cmp),
        "nroCmp": int(nro_cmp),
        "importe": round(float(importe), 2),
        "moneda": str(moneda),
        "ctz": round(float(cotizacion), 6),
        "tipoCodAut": str(tipo_cod_aut),
        "codAut": int(cod_aut),
    }
    if tipo_doc_rec is not None and nro_doc_rec is not None:
        payload["tipoDocRec"] = int(tipo_doc_rec)
        payload["nroDocRec"] = int(nro_doc_rec)

    b64 = base64.b64encode(json.dumps(payload, separators=(",", ":"), ensure_ascii=False).encode("utf-8")).decode("ascii")
    return f"{_QR_BASE_URL}?p={b64}"

def _qr_png_bytes(qr_url):
    if qrcode is None:
        return None
    qr = qrcode.QRCode(error_correction=qrcode.constants.ERROR_CORRECT_M, box_size=6, border=1)
    qr.add_data(qr_url)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    # Asegurar PIL Image real
    try:
        img = img.get_image()
    except Exception:
        pass
    buf = BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()


def restaurar_imagenes_desde_plantilla_excel(plantilla_path, salida_xlsx_path, hojas=("Hoja1","Hoja2","Hoja3"), skip_cells_por_hoja=None):

    skip_cells_por_hoja = skip_cells_por_hoja or {}
    excel = None
    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        wb_src = excel.Workbooks.Open(os.path.abspath(plantilla_path))
        wb_dst = excel.Workbooks.Open(os.path.abspath(salida_xlsx_path))

        for sh_name in hojas:
            try:
                sh_src = wb_src.Worksheets(sh_name)
                sh_dst = wb_dst.Worksheets(sh_name)
            except Exception:
                continue

            skip = set(skip_cells_por_hoja.get(sh_name, set()))

            # Copiamos shapes (imágenes, etc.)
            for shp in list(sh_src.Shapes):
                try:
                    tl = shp.TopLeftCell.Address(False, False)  # "A39"
                except Exception:
                    tl = None
                if tl and tl in skip:
                    continue
                try:
                    shp.Copy()
                    sh_dst.Paste()
                    new_shp = sh_dst.Shapes(sh_dst.Shapes.Count)
                    new_shp.Left = shp.Left
                    new_shp.Top = shp.Top
                    new_shp.Width = shp.Width
                    new_shp.Height = shp.Height
                    try:
                        new_shp.PrintObject = True
                    except Exception:
                        pass
                except Exception:
                    continue

        wb_dst.Save()
        wb_dst.Close(SaveChanges=True)
        wb_src.Close(SaveChanges=False)
    finally:
        if excel is not None:
            excel.Quit()

def insertar_qr_en_excel(salida_xlsx_path, qr_png_bytes, hojas=("Hoja1","Hoja2","Hoja3"), anchor_cell="A39",
                         width=110, height=110):

    if not qr_png_bytes:
        return

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
    try:
        tmp.write(qr_png_bytes)
        tmp.close()

        excel = None
        try:
            excel = win32.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False

            wb = excel.Workbooks.Open(os.path.abspath(salida_xlsx_path))

            for sh_name in hojas:
                try:
                    ws = wb.Worksheets(sh_name)
                except Exception:
                    continue

                # Borra cualquier shape anclada en anchor_cell (placeholder QR)
                try:
                    for shp in list(ws.Shapes):
                        try:
                            tl = shp.TopLeftCell.Address(False, False)
                        except Exception:
                            tl = None
                        if tl == anchor_cell:
                            try:
                                shp.Delete()
                            except Exception:
                                pass
                except Exception:
                    pass

                # Inserta QR
                try:
                    rng = ws.Range(anchor_cell)
                    left, top = rng.Left, rng.Top
                    pic = ws.Shapes.AddPicture(tmp.name, False, True, left, top, width, height)
                    try:
                        pic.PrintObject = True
                    except Exception:
                        pass
                except Exception:
                    continue

            wb.Save()
            wb.Close(SaveChanges=True)
        finally:
            if excel is not None:
                excel.Quit()
    finally:
        try:
            os.remove(tmp.name)
        except Exception:
            pass

def post_procesar_imagenes_y_qr(plantilla_path, factura_output, *, config, tipo_nota, tipo_factura, comprobante, validacion):

    try:
        restaurar_imagenes_desde_plantilla_excel(
            plantilla_path,
            factura_output,
            skip_cells_por_hoja={"Hoja1": {"A39"}, "Hoja2": {"A39"}, "Hoja3": {"A39"}}
        )
        escribir_log(f"{obtener_timestamp()} - Imágenes del template restauradas correctamente.")
    except Exception as e:
        escribir_log(f"{obtener_timestamp()} - WARNING: No se pudieron restaurar imágenes del template: {e}")

    try:
        # Armar datos del QR
        cuit = _to_int_safe(config.get("Cuit") or config.get("CUIT") or config.get("cuit"))
        pto_vta = _to_int_safe(config.get("Punto Venta") or config.get("PtoVta") or config.get("ptoVta"))
        tipo_cmp = _map_tipo_cmp(tipo_nota, tipo_factura)

        if str(tipo_nota).lower() == "factura":
            fecha_emision = _fecha_full_date(comprobante.iloc[1])
            doc_nro = comprobante.iloc[7] if len(comprobante) > 7 else None
        else:
            fecha_emision = _fecha_full_date(comprobante.iloc[4] if len(comprobante) > 4 else None)
            if str(tipo_nota).lower() == "credito":
                doc_nro = comprobante.iloc[8] if len(comprobante) > 8 else None
            else:
                doc_nro = comprobante.iloc[7] if len(comprobante) > 7 else None

        nro_cmp = _to_int_safe(validacion[3])
        importe = _to_float_safe(comprobante.iloc[12] if len(comprobante) > 12 else 0.0, default=0.0)
        cod_aut = _to_int_safe(validacion[1])

        if not all([cuit, pto_vta, tipo_cmp, nro_cmp, cod_aut, fecha_emision]):
            escribir_log(f"{obtener_timestamp()} - WARNING: No se pudo armar QR (faltan datos). cuit={cuit} ptoVta={pto_vta} tipoCmp={tipo_cmp} nroCmp={nro_cmp} codAut={cod_aut} fecha={fecha_emision}")
            return

        tipo_doc_rec, nro_doc_rec = _guess_tipo_doc_rec_y_nro(doc_nro)

        qr_url = _build_qr_url(
            fecha=fecha_emision,
            cuit=cuit,
            pto_vta=pto_vta,
            tipo_cmp=tipo_cmp,
            nro_cmp=nro_cmp,
            importe=importe,
            moneda="PES",
            cotizacion=1.0,
            tipo_doc_rec=tipo_doc_rec,
            nro_doc_rec=nro_doc_rec,
            tipo_cod_aut="E",
            cod_aut=cod_aut
        )

        png = _qr_png_bytes(qr_url)
        if not png:
            escribir_log(f"{obtener_timestamp()} - WARNING: No se pudo generar PNG de QR (qrcode no instalado o falló).")
            return

        insertar_qr_en_excel(factura_output, png)
        escribir_log(f"{obtener_timestamp()} - QR ARCA insertado correctamente.")
    except Exception as e:
        escribir_log(f"{obtener_timestamp()} - WARNING: No se pudo insertar QR: {e}")

#! ------------------------------------------------------------------------------------------------------------------------------------------------------
#!
#! ARMAR EXCELS Y LUEGO GUARDAR COMO PDF'S
#!
#! ------------------------------------------------------------------------------------------------------------------------------------------------------

def qr_presente_en_xlsx(xlsx_path: str) -> bool:

    try:
        wb = load_workbook(xlsx_path)
        for sh in ("Hoja1", "Hoja2", "Hoja3"):
            if sh in wb.sheetnames:
                ws = wb[sh]
                imgs = getattr(ws, "_images", [])
                if imgs and len(imgs) > 0:
                    return True
        return False
    except Exception:

        return False

def completar_plantilla(input_path, plantilla_path, datos, cuerpo_solicitud, config, ListaValidacionCAE, tipo_factura, tipo_nota):
    # Convertir datos y cuerpo_solicitud a DataFrame si son listas
    contador_input = 2
    
    if isinstance(datos, list):
        datos = pd.DataFrame(datos)
    
    if isinstance(cuerpo_solicitud, list):
        cuerpo_solicitud = pd.DataFrame(cuerpo_solicitud)

    for i, validacion in enumerate(ListaValidacionCAE):
        if validacion[0] is False:
            escribir_log(f"{obtener_timestamp()} - Salta CAE inválido en índice: {i}")
            escribir_log("--------------------------------------------------")
            escribir_log(f"{obtener_timestamp()} - Se abre el excel de facturación con el input, para cargar los datos antes de armar la plantilla")  

            # Cargar la plantilla de Excel
            try:

                wb_input = load_workbook(input_path)
                escribir_log(f"{obtener_timestamp()} - Logró abrir correctamente el workbook, con la plantilla {input_path}")
            except FileNotFoundError:
                escribir_log(f"{obtener_timestamp()} - Error: El archivo de plantilla {input_path} no fue encontrado.")
                return
            except PermissionError:
                escribir_log(f"{obtener_timestamp()} - Error: No se tienen permisos para acceder al archivo de plantilla {input_path}.")
                return
            except Exception as e:
                escribir_log(f"{obtener_timestamp()} - Error al cargar la plantilla: {e}")
                return

            ws_input = wb_input[f"Factura {tipo_factura}"]
            ws_input[f"AD{contador_input}"] = "No realizo la Factura"
            ws_input[f"AE{contador_input}"] = "-"
            ws_input[f"AF{contador_input}"] = "-"
            ws_input[f"AG{contador_input}"] = "-"
            contador_input += 1

            wb_input.save(input_path)
            continue

        try:
            escribir_log("--------------------------------------------------")
            escribir_log(f"{obtener_timestamp()} - Se abre el excel de facturación con el input, para cargar los datos antes de armar la plantilla")  

            # Cargar la plantilla de Excel
            try:

                wb_input = load_workbook(input_path)
                escribir_log(f"{obtener_timestamp()} - Logró abrir correctamente el workbook, con la plantilla {input_path}")
            except FileNotFoundError:
                escribir_log(f"{obtener_timestamp()} - Error: El archivo de plantilla {input_path} no fue encontrado.")
                return
            except PermissionError:
                escribir_log(f"{obtener_timestamp()} - Error: No se tienen permisos para acceder al archivo de plantilla {input_path}.")
                return
            except Exception as e:
                escribir_log(f"{obtener_timestamp()} - Error al cargar la plantilla: {e}")
                return

            ws_input = wb_input[f"Factura {tipo_factura}"]
            ws_input[f"AD{contador_input}"] = "Realizo la Factura"
            ws_input[f"AE{contador_input}"] = validacion[1]
            ws_input[f"AF{contador_input}"] = validacion[2]
            ws_input[f"AG{contador_input}"] = validacion[3]
            contador_input += 1

            wb_input.save(input_path)

            escribir_log("--------------------------------------------------")

            escribir_log("Iniciando proceso para completar la plantilla...")
            escribir_log(f"{obtener_timestamp()} - Ruta de la plantilla: {plantilla_path}")
            escribir_log(f"{obtener_timestamp()} - Tipo de comprobante: {tipo_nota} {tipo_factura}")

            solicitud = cuerpo_solicitud.iloc[i]
            comprobante = datos.iloc[i]

            escribir_log(f"{obtener_timestamp()} - Logró completar los datos de solicitud y comprobante")

            # Comprobamos si 'plantilla_path' tiene un valor válido
            if not plantilla_path:
                escribir_log(f"{obtener_timestamp()} - Error: La ruta de la plantilla no está definida.")
                return

            # Cargar la plantilla de Excel
            try:
                wb = load_workbook(plantilla_path)
                escribir_log(f"{obtener_timestamp()} - Logró abrir correctamente el workbook, con la plantilla {plantilla_path}")
            except FileNotFoundError:
                escribir_log(f"{obtener_timestamp()} - Error: El archivo de plantilla {plantilla_path} no fue encontrado.")
                return
            except PermissionError:
                escribir_log(f"{obtener_timestamp()} - Error: No se tienen permisos para acceder al archivo de plantilla {plantilla_path}.")
                return
            except Exception as e:
                escribir_log(f"{obtener_timestamp()} - Error al cargar la plantilla: {e}")
                return

            hojas = ['Hoja1', 'Hoja2', 'Hoja3']

            # Rellenar los datos comunes de la factura
            for hoja_nombre in hojas:
                try:
                    ws = wb[hoja_nombre]

                    # Completa las celdas de la hoja (con ceros a la izquierda donde corresponde)
                    pto_vta_raw = config.get("Punto Venta", "-")
                    pto_vta_int = _to_int_safe(pto_vta_raw)
                    pto_vta_str = f"{pto_vta_int:05d}" if pto_vta_int is not None else str(pto_vta_raw)

                    nro_cmp_int = _to_int_safe(validacion[3])
                    nro_cmp_str = f"{nro_cmp_int:08d}" if nro_cmp_int is not None else str(validacion[3])

                    razon_social = str(config.get("Razón Social", "-") or "-")
                    domicilio_comercial = str(config.get("Domicilio Comercial", "-") or "-")
                    cuit_emisor = str(config.get("Cuit", "-") or "-")

                    ws['G3'] = pto_vta_str
                    ws['G5'] = cuit_emisor
                    ws['B4'] = razon_social
                    ws['C5'] = domicilio_comercial
                    ws['C7'] = config.get("Condición IVA", "-")
                    ws['I3'] = nro_cmp_str 


                    try:
                        rs_up = razon_social.upper()
                        a2_val = ws['A2'].value
                        if a2_val is None or str(a2_val).strip() == "":
                            ws['A2'] = rs_up
                        else:

                            if rs_up not in str(a2_val).upper():
                                ws['A2'] = f"{str(a2_val)} - {rs_up}"
                        ws['A2'].font = ws['A2'].font.copy(sz=10)
                    except Exception:
                        pass


                    try:
                        ws['C5'].font = ws['C5'].font.copy(sz=7)
                    except Exception:
                        pass

                    # Agregar CAE y fecha de vencimiento del CAE
                    ws['H38'] = validacion[1]  # CAE
                    ws['H39'] = validacion[2]  # Fecha vencimiento CAE

                    if tipo_nota == "Factura":
                        ws['G4'] = comprobante.iloc[1]  # Fecha Emisión
                        ws['C8'] = comprobante.iloc[2]  # Fecha periodo desde
                        ws['E8'] = comprobante.iloc[3]  # Fecha periodo hasta
                        ws['I8'] = comprobante.iloc[1]  # Fecha Vencimiento pago, por ahora puesto con la misma de fecha emisión 
                        ws['C11'] = comprobante.iloc[4]  # Condición frente al IVA del comprador
                        ws['F9'] = comprobante.iloc[6]  # Razón Social, Cliente
                        ws['B9'] = comprobante.iloc[7]  # CUIT / CUIL
                        ws['F11'] = comprobante.iloc[8]  # Domicilio
                        ws['C13'] = comprobante.iloc[13]  # Condición de venta

                    elif tipo_nota == "Credito":
                        ws['G4'] = comprobante.iloc[4]  # Fecha Emisión
                        ws['C8'] = comprobante.iloc[5]  # Fecha periodo desde
                        ws['E8'] = comprobante.iloc[6]  # Fecha periodo hasta
                        ws['I8'] = comprobante.iloc[7]  # Fecha Vencimiento pago, por ahora puesto con la misma de fecha emisión 
                        ws['C11'] = comprobante.iloc[3]  # Condición frente al IVA del comprador
                        ws['F9'] = comprobante.iloc[9]  # Razón Social, Cliente
                        ws['B9'] = comprobante.iloc[8]  # CUIT / CUIL
                        ws['F11'] = comprobante.iloc[11]  # Domicilio
                        ws['C13'] = comprobante.iloc[20]  # Condición de venta
                    
                    elif tipo_nota == "Debito":
                        ws['G4'] = comprobante.iloc[4]  # Fecha Emisión
                        ws['C8'] = comprobante.iloc[5]  # Fecha periodo desde
                        ws['E8'] = comprobante.iloc[6]  # Fecha periodo hasta
                        ws['I8'] = comprobante.iloc[7]  # Fecha Vencimiento pago, por ahora puesto con la misma de fecha emisión 
                        ws['C11'] = comprobante.iloc[3]  # Condición frente al IVA del comprador
                        ws['F9'] = comprobante.iloc[6]  # Razón Social, Cliente
                        ws['B9'] = comprobante.iloc[7]  # CUIT / CUIL
                        ws['F11'] = comprobante.iloc[8]  # Domicilio
                        ws['C13'] = comprobante.iloc[20]  # Condición de venta

                    # Función para verificar si una celda está combinada
                    def es_celda_combinada(ws, fila, columna):
                        for rango in ws.merged_cells.ranges:
                            if ws.cell(row=fila, column=columna).coordinate in rango:
                                return True
                        return False

                    # Inicializar la fila donde se empezarán a escribir los ítems
                    fila = 15

                    # Acceder a los elementos de la solicitud JSON
                    if 'arrayItems' in solicitud:  # Verificar que exista el array de ítems
                        # Recorrer todos los ítems de arrayItems
                        for idx, item_group in enumerate(solicitud['arrayItems'], start=1):
                            # Dentro de cada grupo de ítems, recorrer cada "item"
                            for item in item_group['item']:
                                # Verificar antes de escribir si la celda está combinada
                                if not es_celda_combinada(ws, fila, 1):
                                    ws[f'A{fila}'] = item['codigo']
                                if not es_celda_combinada(ws, fila, 2):
                                    ws[f'B{fila}'] = item['descripcion']
                                if not es_celda_combinada(ws, fila, 3):
                                    _set_num(ws, f'D{fila}', item['cantidad'])
                                if not es_celda_combinada(ws, fila, 4):
                                    _set_num(ws, f'G{fila}', item['precioUnitario'])
                                if not es_celda_combinada(ws, fila, 5):
                                    _set_num(ws, f'I{fila}', item['importeItem'])
                                if not es_celda_combinada(ws, fila, 6):
                                    _set_num(ws, f'H{fila}', item.get('importeBonificacion', 0))  # Importe de Bonificación (si existe)
                                if not es_celda_combinada(ws, fila, 7):
                                    _set_num(ws, f'E{fila}', item.get('importeIVA', 0))  # Importe de IVA (si existe)
                                if not es_celda_combinada(ws, fila, 8):
                                    if item['codigoCondicionIVA'] == "3":
                                        ws[f'F{fila}'] = "0%"  # Porcentaje Del IVA
                                    if item['codigoCondicionIVA'] == "4":
                                        ws[f'F{fila}'] = "10,50%"  # Porcentaje Del IVA
                                    if item['codigoCondicionIVA'] == "5":
                                        ws[f'F{fila}'] = "21%"  # Porcentaje Del IVA
                                    if item['codigoCondicionIVA'] == "6":
                                        ws[f'F{fila}'] = "27%"  # Porcentaje Del IVA
                                    if item['codigoCondicionIVA'] == "8":
                                        ws[f'F{fila}'] = "5%"  # Porcentaje Del IVA
                                    if item['codigoCondicionIVA'] == "9":
                                        ws[f'F{fila}'] = "2.50%"  # Porcentaje Del IVA
                                fila += 1  # Avanzar a la siguiente fila después de cada ítem
                    else:
                        # Si entro en este else, es porque solo tiene 1 item:
                        numero_producto = 1 

                        if tipo_factura == "A" or tipo_factura == "B":
                            # Primer producto que está directamente en el comprobante
                            ws[f'A{fila}'] = numero_producto
                            _set_num(ws, f'D{fila}', comprobante.iloc[9])  # Cantidad de cosas vendidas
                            ws[f'B{fila}'] = comprobante.iloc[10]  # Descripción del producto (producto / servicio)
                            _set_num(ws, f'G{fila}', int(comprobante.iloc[11]))  # Precio Unitario
                            _set_num(ws, f'I{fila}', int(comprobante.iloc[12]))  # Total de todo
                            if tipo_factura == "A":
                                _set_num(ws, f'E{fila}', comprobante.iloc[18])  # Importe IVA  
                                if comprobante.iloc[27] == "3":
                                    ws[f'F{fila}'] = "0%"  # Porcentaje Del IVA
                                if comprobante.iloc[27] == "4":
                                    ws[f'F{fila}'] = "10,50%"  # Porcentaje Del IVA
                                if comprobante.iloc[27] == "5":
                                    ws[f'F{fila}'] = "21%"  # Porcentaje Del IVA
                                if comprobante.iloc[27] == "6":
                                    ws[f'F{fila}'] = "27%"  # Porcentaje Del IVA
                                if comprobante.iloc[27] == "8":
                                    ws[f'F{fila}'] = "5%"  # Porcentaje Del IVA
                                if comprobante.iloc[27] == "9":
                                    ws[f'F{fila}'] = "2.50%"  # Porcentaje Del IVA
                            else:
                                ws[f'F{fila}'] = int(comprobante.iloc[15])  # Porcentaje (%) de Bonificación                                
                            _set_num(ws, f'H{fila}', int(comprobante.iloc[16]))  # Importe de bonificación

                        else:
                            # Primer producto que está directamente en el comprobante
                            ws[f'A{fila}'] = numero_producto
                            _set_num(ws, f'D{fila}', comprobante.iloc[9])  # Cantidad de cosas vendidas
                            ws[f'B{fila}'] = comprobante.iloc[10]  # Descripción del producto (producto / servicio)
                            _set_num(ws, f'F{fila}', int(comprobante.iloc[11]))  # Precio Unitario
                            _set_num(ws, f'I{fila}', int(comprobante.iloc[12]))  # Total de todo
                            ws[f'E{fila}'] = comprobante.iloc[14]  # Unidad de Medida     
                            ws[f'G{fila}'] = int(comprobante.iloc[15])  # Porcentaje (%) de Bonificación
                            _set_num(ws, f'H{fila}', int(comprobante.iloc[16]))  # Importe de bonificación
                            
                except KeyError:
                    escribir_log(f"{obtener_timestamp()} - Error: La hoja {hoja_nombre} no existe en el archivo de plantilla.")
                except Exception as e:
                    escribir_log(f"{obtener_timestamp()} - Error al intentar completar la hoja {hoja_nombre}: {e}")
            # Guardar la factura en un nuevo archivo (nombre estándar: CUIT_TIPO(3)_PTO(5)_NRO(8))
            try:
                cuit_arch = _solo_digitos(config.get("Cuit", ""))
                tipo_cmp = _map_tipo_cmp(tipo_nota, tipo_factura)
                tipo_cmp_str = f"{int(tipo_cmp):03d}" if tipo_cmp is not None else "000"

                pto_vta_int = _to_int_safe(config.get("Punto Venta"))
                pto_vta_str = f"{pto_vta_int:05d}" if pto_vta_int is not None else str(config.get("Punto Venta", "-"))

                nro_cmp_int = _to_int_safe(validacion[3])
                nro_cmp_str = f"{nro_cmp_int:08d}" if nro_cmp_int is not None else str(validacion[3])

                nombre_base = f"{cuit_arch}_{tipo_cmp_str}_{pto_vta_str}_{nro_cmp_str}"
            except Exception:
                nombre_base = f"Comprobante_{validacion[3]}"

            factura_output = os.path.join(os.path.dirname(plantilla_path), f"{nombre_base}.xlsx")
            wb.save(factura_output)
            escribir_log(f"{obtener_timestamp()} - Factura generada y guardada en {factura_output}.")

            # 1) Post-procesado general (copia imágenes + intenta QR)
            try:
                post_procesar_imagenes_y_qr(
                    plantilla_path,
                    factura_output,
                    config=config,
                    tipo_nota=tipo_nota,
                    tipo_factura=tipo_factura,
                    comprobante=comprobante,
                    validacion=validacion
                )
            except Exception as e:
                escribir_log(f"{obtener_timestamp()} - WARNING: Falló el post-procesado de imágenes/QR: {e}")

            # 2) Reintento SOLO si falta el QR (para que sea “sí o sí”)
            try:
                if not qr_presente_en_xlsx(factura_output):
                    asegurar_qr_en_factura(
                        factura_output,
                        config=config,
                        solicitud=solicitud,
                        validacion=validacion,
                    )
            except Exception as e:
                escribir_log(f"{obtener_timestamp()} - WARNING: Falló el reintento del QR: {e}")
        
        
        except Exception as e:
                escribir_log(f"{obtener_timestamp()} - Error al completar plantilla para índice {i}: {str(e)}")
                print(f"{obtener_timestamp()} - Error al completar plantilla para índice {i}: {str(e)}")

    os.remove(plantilla_path)

#! FUNCIONALIDAD PARA ARMAR PDF's Y LUEGO ALMACENARLOS EN LA CARPETA CORRESPONDIENTE
def convertir_xlsx_a_pdf(carpeta_origen, carpeta_destino):

    carpeta_dia = os.path.join(carpeta_destino, _dt.datetime.now().strftime("%d-%m-%Y"))
    os.makedirs(carpeta_dia, exist_ok=True)

    excel = win32.Dispatch('Excel.Application')
    excel.Visible = False
    try:
        # Evitar prompts de Excel
        try:
            excel.DisplayAlerts = False
        except Exception:
            pass

        for archivo in os.listdir(carpeta_origen):
            if not archivo.lower().endswith('.xlsx'):
                continue
            if archivo.startswith('~$'):
                continue  # temporales de Excel

            ruta_xlsx = os.path.join(carpeta_origen, archivo)
            ruta_pdf = os.path.join(carpeta_dia, os.path.splitext(archivo)[0] + '.pdf')

            workbook = None
            try:
                workbook = excel.Workbooks.Open(ruta_xlsx)
                workbook.ExportAsFixedFormat(0, ruta_pdf)
                workbook.Close(SaveChanges=False)

                print(f"{obtener_timestamp()} - Convertido: {archivo} a {ruta_pdf}")
                escribir_log("")
                escribir_log(f"{obtener_timestamp()} - Convertido: {archivo} a {ruta_pdf}")

            except Exception as e:
                print(f"{obtener_timestamp()} - Error al convertir {archivo}: {e}")
                escribir_log(f"{obtener_timestamp()} - Error al convertir {archivo}: {e}")
                try:
                    if workbook is not None:
                        workbook.Close(SaveChanges=False)
                except Exception:
                    pass

            # Mover el xlsx al folder del día (convertido o no)
            try:
                destino_xlsx = os.path.join(carpeta_dia, archivo)
                if os.path.exists(destino_xlsx):
                    base, ext = os.path.splitext(archivo)
                    n = 1
                    while os.path.exists(os.path.join(carpeta_dia, f"{base}_{n}{ext}")):
                        n += 1
                    destino_xlsx = os.path.join(carpeta_dia, f"{base}_{n}{ext}")
                shutil.move(ruta_xlsx, destino_xlsx)
            except Exception as e:
                escribir_log(f"{obtener_timestamp()} - WARNING: No se pudo mover {archivo} a carpeta del día: {e}")

    finally:
        try:
            excel.Quit()
        except Exception:
            pass
