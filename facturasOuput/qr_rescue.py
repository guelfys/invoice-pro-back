# FacturasOuput/qr_rescue.py
import json
import base64
from io import BytesIO
import datetime as _dt
import tempfile
import os

import win32com.client as win32

try:
    import qrcode
except Exception:
    qrcode = None

_QR_BASE_URL = "https://www.arca.gob.ar/fe/qr/"


def _to_int_safe(v):
    try:
        if v is None:
            return None
        if isinstance(v, (int, float)):
            return int(v)
        s = str(v).strip()
        if not s:
            return None
        # s puede venir como "20-..." o con puntos
        s = "".join(ch for ch in s if ch.isdigit())
        return int(s) if s else None
    except Exception:
        return None


def _to_float_safe(v):
    try:
        if v is None:
            return None
        if isinstance(v, (int, float)):
            return float(v)
        s = str(v).strip().replace(".", "").replace(",", ".")
        return float(s)
    except Exception:
        return None


def _guess_fecha_emision(solicitud: dict):

    if isinstance(solicitud, dict):
        for k in ("fecha_emision", "fechaEmision", "cbte_fch", "CbteFch", "fecha", "Fecha"):
            if k in solicitud and solicitud[k]:
                val = str(solicitud[k]).strip()

                # si viene "YYYY-MM-DD"
                if "-" in val and len(val) >= 10:
                    try:
                        d = _dt.datetime.fromisoformat(val[:10]).date()
                        return d.strftime("%Y%m%d")
                    except Exception:
                        pass

                # si viene "YYYYMMDD"
                if val.isdigit() and len(val) == 8:
                    return val
    return _dt.date.today().strftime("%Y%m%d")


def _infer_doc_tipo(doc_nro: int | None):

    if not doc_nro:
        return 99
    ln = len(str(doc_nro))
    if ln == 11:
        return 80
    if ln == 8:
        return 96
    return 99


def _build_qr_url(data: dict) -> str:
    payload = json.dumps(data, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
    b64 = base64.urlsafe_b64encode(payload).decode("ascii")
    return f"{_QR_BASE_URL}?p={b64}"


def _generar_qr_png_bytes(qr_url: str) -> bytes:
    if qrcode is None:
        raise RuntimeError(
            "No está disponible la librería 'qrcode'. "
            "En el EXE tenés que incluirla en el build (ver punto 3)."
        )

    qr = qrcode.QRCode(
        version=None,
        error_correction=qrcode.constants.ERROR_CORRECT_M,
        box_size=4,
        border=1
    )
    qr.add_data(qr_url)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")

    bio = BytesIO()
    img.save(bio, format="PNG")
    return bio.getvalue()


def insertar_qr_en_excel(
    salida_xlsx_path: str,
    qr_png_bytes: bytes,
    hojas=("Hoja1", "Hoja2", "Hoja3"),
    anchor_cell="A39",
    width=110,
    height=110
):

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


def asegurar_qr_en_factura(
    factura_output_xlsx: str,
    *,
    config: dict,
    solicitud: dict,
    validacion: list,
):

    # ---- datos base
    cuit = _to_int_safe(config.get("Cuit") or config.get("CUIT"))
    pto_vta = _to_int_safe(config.get("PuntoVenta") or config.get("pto_vta") or config.get("PtoVta"))
    tipo_cmp = _to_int_safe(config.get("CodigoTipoComprobante") or config.get("tipo_cmp") or config.get("CbteTipo"))
    nro_cmp = _to_int_safe(validacion[3] if len(validacion) > 3 else None)
    cae = str(validacion[1]).strip() if len(validacion) > 1 and validacion[1] else None

    fecha = _guess_fecha_emision(solicitud)

    doc_nro = None
    if isinstance(solicitud, dict):
        for k in ("doc_nro", "DocNro", "nro_doc", "nroDocRec", "numero_documento", "documento"):
            if k in solicitud and solicitud[k]:
                doc_nro = _to_int_safe(solicitud[k])
                if doc_nro:
                    break

    doc_tipo = None
    if isinstance(solicitud, dict):
        for k in ("doc_tipo", "DocTipo", "tipo_doc", "tipoDocRec"):
            if k in solicitud and solicitud[k]:
                doc_tipo = _to_int_safe(solicitud[k])
                if doc_tipo:
                    break
    if not doc_tipo:
        doc_tipo = _infer_doc_tipo(doc_nro)

    imp_total = None
    if isinstance(solicitud, dict):
        for k in ("importe_total", "importeTotal", "ImpTotal", "imp_total", "total", "importe"):
            if k in solicitud and solicitud[k] is not None:
                imp_total = _to_float_safe(solicitud[k])
                if imp_total is not None:
                    break

    if imp_total is None:
        imp_total = 0.0

    if not all([cuit, pto_vta, tipo_cmp, nro_cmp, cae, fecha]):

        raise RuntimeError(
            f"Faltan datos mínimos QR. cuit={cuit}, pto_vta={pto_vta}, tipo_cmp={tipo_cmp}, "
            f"nro_cmp={nro_cmp}, cae={cae}, fecha={fecha}"
        )


    qr_data = {
        "ver": 1,
        "fecha": fecha,
        "cuit": cuit,
        "ptoVta": pto_vta,
        "tipoCmp": tipo_cmp,
        "nroCmp": nro_cmp,
        "importe": float(imp_total),
        "moneda": "PES",
        "ctz": 1,
        "tipoDocRec": int(doc_tipo) if doc_tipo else 99,
        "nroDocRec": int(doc_nro) if doc_nro else 0,
        "tipoCodAut": "E",
        "codAut": str(cae),
    }
    qr_url = _build_qr_url(qr_data)


    qr_png = _generar_qr_png_bytes(qr_url)
    insertar_qr_en_excel(factura_output_xlsx, qr_png)

    return qr_url
