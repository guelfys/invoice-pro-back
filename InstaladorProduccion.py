import os
import shutil
import sys
import tkinter as tk
from tkinter import filedialog, messagebox

# =========================
# Nombres NUEVOS (sin acentos)
# =========================
AUTH_ABC_NEW = "Autorizacion_ABC_sin_item"
AUTH_AB_NEW  = "Autorizacion_AB_con_item"

# Nombres VIEJOS (por si el bundle aún los trae así)
AUTH_ABC_OLD = "Autorización A, B y C sin item"
AUTH_AB_OLD  = "Autorización A y B con item"

# Si querés que deje también las carpetas viejas para compatibilidad (recomendado por ahora)
CREAR_COMPAT_NOMBRES_VIEJOS = False


def get_bundle_dir() -> str:
    """Carpeta desde donde copiar recursos (PyInstaller: _MEIPASS / Script: dir actual)."""
    if hasattr(sys, "_MEIPASS"):
        return sys._MEIPASS  # type: ignore[attr-defined]
    return os.path.dirname(os.path.abspath(__file__))


def solicitar_carpeta_instalacion() -> str | None:
    root = tk.Tk()
    root.withdraw()
    carpeta = filedialog.askdirectory(title="Selecciona la carpeta de instalación")
    root.destroy()

    if not carpeta:
        messagebox.showerror("Error", "Debes seleccionar una carpeta de instalación.")
        return None
    return carpeta


def copy_dir_if_exists(src: str, dst: str) -> bool:
    if os.path.isdir(src):
        os.makedirs(dst, exist_ok=True)
        shutil.copytree(src, dst, dirs_exist_ok=True)
        print(f"[OK] Carpeta copiada: {dst}")
        return True
    print(f"[SKIP] No existe carpeta: {src}")
    return False


def copy_file_if_exists(src: str, dst: str) -> bool:
    if os.path.isfile(src):
        os.makedirs(os.path.dirname(dst), exist_ok=True)
        shutil.copy2(src, dst)
        print(f"[OK] Archivo copiado: {dst}")
        return True
    print(f"[SKIP] No existe archivo: {src}")
    return False


def crear_estructura_base(base_dir: str) -> None:
    carpetas = [
        os.path.join(base_dir, "log"),
        os.path.join(base_dir, "input", "Factura A"),
        os.path.join(base_dir, "input", "Factura B"),
        os.path.join(base_dir, "input", "Factura C"),
        os.path.join(base_dir, "output"),
        os.path.join(base_dir, "facturas"),
        os.path.join(base_dir, "Usuarios"),
        os.path.join(base_dir, "sources"),
    ]
    for c in carpetas:
        os.makedirs(c, exist_ok=True)
        print(f"[OK] Carpeta creada: {c}")


def instalar_autorizaciones(bundle_dir: str, base_dir: str) -> None:
    """
    Copia carpetas de autorización del bundle a destino:
      - instala SIEMPRE los nombres NUEVOS sin acentos
      - opcional: crea también copia con nombres viejos (compat temporal)
    """
    src_abc_old = os.path.join(bundle_dir, AUTH_ABC_OLD)
    src_ab_old  = os.path.join(bundle_dir, AUTH_AB_OLD)
    src_abc_new = os.path.join(bundle_dir, AUTH_ABC_NEW)
    src_ab_new  = os.path.join(bundle_dir, AUTH_AB_NEW)

    dst_abc_new = os.path.join(base_dir, AUTH_ABC_NEW)
    dst_ab_new  = os.path.join(base_dir, AUTH_AB_NEW)

    # 1) Copiar a nombres nuevos (preferimos fuentes viejas si son las que existen)
    ok_abc = copy_dir_if_exists(src_abc_old, dst_abc_new) or copy_dir_if_exists(src_abc_new, dst_abc_new)
    ok_ab  = copy_dir_if_exists(src_ab_old,  dst_ab_new)  or copy_dir_if_exists(src_ab_new,  dst_ab_new)

    if not ok_abc:
        print("[WARN] No se pudo copiar autorización ABC (no encontrada en bundle).")
    if not ok_ab:
        print("[WARN] No se pudo copiar autorización AB (no encontrada en bundle).")

    # 2) Compat: dejar también los nombres viejos para que el código actual no se rompa (por ahora)
    if CREAR_COMPAT_NOMBRES_VIEJOS:
        if os.path.isdir(dst_abc_new):
            copy_dir_if_exists(dst_abc_new, os.path.join(base_dir, AUTH_ABC_OLD))
        if os.path.isdir(dst_ab_new):
            copy_dir_if_exists(dst_ab_new, os.path.join(base_dir, AUTH_AB_OLD))


def instalar_sources(bundle_dir: str, base_dir: str) -> None:
    """
    Copia carpeta sources completa.
    Se espera:
      sources/demo/fev1
      sources/demo/mtxca
      sources/produccion/fev1
      sources/produccion/mtxca
    """
    copy_dir_if_exists(os.path.join(bundle_dir, "sources"), os.path.join(base_dir, "sources"))


def instalar_archivos_comunes(bundle_dir: str, base_dir: str) -> None:
    copy_dir_if_exists(os.path.join(bundle_dir, "Configuraciones"), os.path.join(base_dir, "Configuraciones"))
    copy_dir_if_exists(os.path.join(bundle_dir, "input", "Ejemplos Input"), os.path.join(base_dir, "input", "Ejemplos Input"))

    copy_file_if_exists(os.path.join(bundle_dir, "base.png"), os.path.join(base_dir, "base.png"))

    # Plantillas
    copy_file_if_exists(os.path.join(bundle_dir, "vacio_factura_a.xlsx"), os.path.join(base_dir, "vacio_factura_a.xlsx"))
    copy_file_if_exists(os.path.join(bundle_dir, "vacio_factura_b.xlsx"), os.path.join(base_dir, "vacio_factura_b.xlsx"))
    copy_file_if_exists(os.path.join(bundle_dir, "vacio_factura_c.xlsx"), os.path.join(base_dir, "vacio_factura_c.xlsx"))

    # Docs
    copy_file_if_exists(os.path.join(bundle_dir, "Documentación Funcional.docx"), os.path.join(base_dir, "Documentación Funcional.docx"))
    copy_file_if_exists(os.path.join(bundle_dir, "Documentación Tecnica.docx"), os.path.join(base_dir, "Documentación Tecnica.docx"))
    copy_file_if_exists(os.path.join(bundle_dir, "Manual Autorizaciones AFIP.docx"), os.path.join(base_dir, "Manual Autorizaciones AFIP.docx"))


def instalar_ejecutables(bundle_dir: str, base_dir: str) -> None:
    """
    Copia ejecutables si existen. Copiamos TODOS para tener demo + prod en una instalación.
    """
    candidatos = [
        "Desarrollo_APP_InvoicerPRO.exe",
        "Desarrollo_Programable_InvoicerPRO.exe",
        "Produccion_APP_InvoicerPRO.exe",
        "Produccion_Programable_InvoicerPRO.exe",
    ]
    for exe in candidatos:
        copy_file_if_exists(os.path.join(bundle_dir, exe), os.path.join(base_dir, exe))


def main():
    base_dir = solicitar_carpeta_instalacion()
    if not base_dir:
        return

    bundle_dir = get_bundle_dir()

    crear_estructura_base(base_dir)
    instalar_autorizaciones(bundle_dir, base_dir)
    instalar_sources(bundle_dir, base_dir)
    instalar_archivos_comunes(bundle_dir, base_dir)
    instalar_ejecutables(bundle_dir, base_dir)

    messagebox.showinfo("Listo", "Instalación Global completada.")


if __name__ == "__main__":
    main()
