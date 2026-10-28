#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Autoprueba del equipo: «ValidadorRenovaciones.exe --autoprueba» (o Revisar_Equipo.bat). Revisa que el ejecutable traiga
todo lo que usa la app (librerías, Excel, OCR, lector de PDF, conexión a la base) y que el equipo tenga Outlook abierto
y el Oracle Client. No se conecta a la base con usuario ni contraseña: solo prueba que el conector funcione.
Escribe autoprueba.txt junto a la app y muestra el resumen.
"""

import importlib
import json
import os
import platform
import re
import sys
import tempfile
import time

OK, AVISO, FALLA = "OK", "AVISO", "FALLA"


def _librerias():
    for modulo in ("customtkinter", "tkcalendar", "pandas", "openpyxl", "pythoncom", "win32clipboard", "win32com.client",
                   "win32cred", "win32timezone", "PIL.Image"):
        importlib.import_module(modulo)
    from babel.numbers import format_decimal
    from bs4 import BeautifulSoup
    format_decimal(1234.5, locale="es_MX")
    BeautifulSoup("<table><tr><td>1</td></tr></table>", "lxml").find("td")
    return OK, "CustomTkinter, calendario, pandas, Excel, Outlook (pywin32), lxml, Pillow e idioma español"


def _cifrado():
    # Lo que usa oracledb (modo thin) para entrar a la base: AES, PBKDF2 y SHA-512
    from cryptography.hazmat.primitives import hashes, padding
    from cryptography.hazmat.primitives.ciphers import Cipher, algorithms, modes
    from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
    llave = PBKDF2HMAC(algorithm=hashes.SHA512(), length=32, salt=b"x" * 16, iterations=1000).derive(b"prueba")
    relleno = padding.PKCS7(128).padder()
    datos = relleno.update(b"prueba") + relleno.finalize()
    cifrador = Cipher(algorithms.AES(llave), modes.CBC(b"\0" * 16)).encryptor()
    cifrador.update(datos) + cifrador.finalize()
    import cryptography
    return OK, f"cryptography {cryptography.__version__}"


def _oracle_thin():
    # Contra un puerto cerrado de este equipo: si el conector está completo, el error es «no responde», no DPY-3016
    import oracledb
    try:
        oracledb.connect(user="prueba", password="prueba", dsn="127.0.0.1:9/prueba", tcp_connect_timeout=3)
    except Exception as e:
        texto = str(e).splitlines()[0] if str(e) else type(e).__name__
        if "DPY-3016" in texto or "cryptography" in texto:
            return FALLA, f"oracledb {oracledb.__version__}: {texto}"
        if "DPY-6005" in texto or "DPY-6000" in texto or "refused" in texto.lower() or "10061" in texto:
            return OK, f"oracledb {oracledb.__version__} (modo thin listo)"
        return AVISO, f"oracledb {oracledb.__version__}: respuesta inesperada: {texto[:160]}"
    return AVISO, "respuesta inesperada: se conectó a un puerto cerrado"


def _oracle_cliente(config: dict):
    # El modo clásico, que la base de FONACOT pide (DPY-3015): el Oracle Client que trae la app o el instalado
    import oracledb
    import sql_renovaciones as sql
    carpetas = [config["sql_cliente_oracle"]] if config.get("sql_cliente_oracle") else sql._clientes_oracle()
    if not carpetas:
        return AVISO, ("no se encontró un Oracle Client de 64 bits en ningún disco (SQL Developer no lo trae): si "
                       "la base pide el modo clásico (DPY-3015), la validación en la base no funcionará")
    fallos = []
    for carpeta in carpetas:
        try:
            oracledb.init_oracle_client(lib_dir=carpeta)
        except Exception as e:
            fallos.append(sql._explicar_cliente_viejo(f"{carpeta}: {str(e).splitlines()[0]}", carpeta))
            continue
        version = ".".join(str(n) for n in oracledb.clientversion())
        try:
            oracledb.connect(user="prueba", password="prueba", dsn="127.0.0.1:9/prueba")
        except Exception as e:
            texto = str(e).splitlines()[0] if str(e) else type(e).__name__
            if not re.search(r"ORA-125\d\d|ORA-12170|TNS-", texto):
                return AVISO, f"Oracle Client {version} en {carpeta}: respuesta inesperada: {texto[:160]}"
        return OK, f"Oracle Client {version} en {carpeta} (modo clásico listo)"
    return FALLA, "no se pudo usar ningún Oracle Client: " + "  ·  ".join(fallos)


def _servidor(config: dict):
    # Solo abre y cierra el puerto del servidor de la base (sin usuario ni contraseña)
    import socket
    import sql_renovaciones as sql
    servidor = sql.servidor_de(config.get("sql_host", ""), config.get("sql_puerto", ""), config.get("sql_servicio", ""))
    if not servidor:
        return AVISO, "aún no hay servidor en «Validar en la base de datos»"
    host, puerto, _ = sql.separar_servidor(servidor)
    puerto = int(puerto or 1521)
    try:
        socket.create_connection((host, puerto), timeout=5).close()
    except OSError as e:
        return AVISO, (f"no responde {host}:{puerto} ({e.strerror or e}): revisa el Nombre del host y que estés en la "
                       f"red de FONACOT (VPN)")
    return OK, f"{host}:{puerto} responde"


def _excel():
    import pandas as pd
    with tempfile.TemporaryDirectory() as tmp:
        ruta = os.path.join(tmp, "prueba.xlsx")
        pd.DataFrame({"Cliente": ["7438176"], "Crédito": ["57815"]}).to_excel(ruta, index=False, engine="openpyxl")
        tam = os.path.getsize(ruta)
    return OK, f"se generó un Excel de prueba ({tam} bytes)"


def _ocr():
    from PIL import Image, ImageDraw, ImageFont
    import ocr_imagenes as ocr
    with tempfile.TemporaryDirectory() as tmp:
        ruta = os.path.join(tmp, "prueba.png")
        im = Image.new("RGB", (900, 160), "white")
        try:
            letra = ImageFont.truetype("arial.ttf", 48)
        except OSError:
            letra = ImageFont.load_default(size=48)
        ImageDraw.Draw(im).text((20, 50), "CLIENTE 7438176 CREDITO 57815", fill="black", font=letra)
        im.save(ruta)
        try:
            lecturas = ocr.leer_imagenes([ruta]).get(ruta, [])
        except ocr.OCRNoDisponible as e:
            return FALLA, f"el OCR de Windows no está disponible: {e}"
    texto = " ".join(p[0] for version in lecturas for linea in version for p in linea)
    if "7438176" in texto and "57815" in texto:
        return OK, "el OCR de Windows leyó la imagen de prueba"
    return AVISO, f"el OCR respondió pero no leyó bien la imagen de prueba («{texto[:80]}»)"


def _pdf(app_dir: str):
    import ocr_imagenes as ocr
    pdf = os.path.join(app_dir, "GUIA_DE_USO.pdf")
    if not os.path.isfile(pdf):
        return AVISO, "no está GUIA_DE_USO.pdf junto a la app para probar el lector de PDF"
    with tempfile.TemporaryDirectory() as tmp:
        notas = []
        paginas = ocr.pdf_a_imagenes([(pdf, os.path.join(tmp, "guia"))], notas)
        if paginas:
            return OK, f"el lector de PDF de Windows convirtió {len(paginas)} página(s) de la guía"
    return FALLA, "el lector de PDF de Windows no pudo abrir la guía" + (f" ({'; '.join(notas)})" if notas else "")


def _outlook():
    import pythoncom
    import win32com.client
    pythoncom.CoInitialize()
    try:
        try:
            app = win32com.client.GetActiveObject("Outlook.Application")
        except Exception:
            return AVISO, "Outlook no está abierto (ábrelo antes de «Extraer y validar»)"
        buzones = [app.Session.Stores.Item(i).DisplayName for i in range(1, app.Session.Stores.Count + 1)]
        return OK, f"Outlook abierto con {len(buzones)} buzón(es): {', '.join(buzones)[:150]}"
    finally:
        pythoncom.CoUninitialize()


def _carpeta(app_dir: str):
    prueba = os.path.join(app_dir, "autoprueba.tmp")
    try:
        with open(prueba, "w", encoding="utf-8") as f:
            f.write("ok")
        os.remove(prueba)
    except OSError as e:
        return FALLA, (f"no se puede escribir en {app_dir} ({e.strerror}): mueve la carpeta a Documentos o al "
                       f"Escritorio")
    return OK, f"se puede escribir en {app_dir}"


def _internet():
    import urllib.request
    from calendario_cedulas import URL_CALENDARIO
    pedido = urllib.request.Request(URL_CALENDARIO, headers={"User-Agent": "Mozilla/5.0"})
    try:
        with urllib.request.urlopen(pedido, timeout=10) as resp:
            return OK, f"página del calendario de cédulas de FONACOT (HTTP {resp.status})"
    except Exception as e:
        return AVISO, (f"no se pudo abrir la página del calendario de cédulas ({str(e)[:120]}): se usan las fechas "
                       f"que trae la app")


def revisar() -> list:
    """[(prueba, nivel, detalle)] de cada comprobación; ninguna detiene a las demás."""
    from core_renovaciones import get_app_dir
    app_dir = get_app_dir()
    try:
        with open(os.path.join(os.path.expanduser("~"), ".fonacot_validador_renovaciones_config.json"),
                  encoding="utf-8") as f:
            config = json.load(f)
    except (OSError, ValueError):
        config = {}
    pruebas = [
        ("Librerías", _librerias),
        ("Cifrado (para la base)", _cifrado),
        ("Conector de la base (modo thin)", _oracle_thin),
        ("Oracle Client (modo clásico)", lambda: _oracle_cliente(config)),
        ("Servidor de la base", lambda: _servidor(config)),
        ("Excel", _excel),
        ("OCR de capturas", _ocr),
        ("Lector de PDF", lambda: _pdf(app_dir)),
        ("Outlook", _outlook),
        ("Carpeta de la app", lambda: _carpeta(app_dir)),
        ("Internet (calendario)", _internet),
    ]
    resultados = []
    for nombre, prueba in pruebas:
        inicio = time.time()
        try:
            nivel, detalle = prueba()
        except Exception as e:
            nivel, detalle = FALLA, f"{type(e).__name__}: {e}"
        resultados.append((nombre, nivel, f"{detalle}  [{time.time() - inicio:.1f} s]"))
    return resultados


def reporte(resultados: list, version: str = "") -> str:
    lineas = [f"Autoprueba del Validador de Renovaciones v{version} · {time.strftime('%d/%m/%Y %H:%M')}",
              f"Equipo: {platform.node()} · Windows {platform.version()} · "
              f"{'ejecutable' if getattr(sys, 'frozen', False) else 'Python ' + platform.python_version()}", ""]
    for nombre, nivel, detalle in resultados:
        lineas.append(f"[{nivel:5}] {nombre}: {detalle}")
    fallas = sum(1 for _, n, _ in resultados if n == FALLA)
    avisos = sum(1 for _, n, _ in resultados if n == AVISO)
    lineas += ["", f"Resultado: {fallas} falla(s), {avisos} aviso(s)."]
    return "\n".join(lineas)


def main(version: str = "", con_ventana: bool = True) -> int:
    from core_renovaciones import get_app_dir
    texto = reporte(revisar(), version)
    ruta = os.path.join(get_app_dir(), "autoprueba.txt")
    try:
        with open(ruta, "w", encoding="utf-8") as f:
            f.write(texto + "\n")
    except OSError:
        ruta = os.path.join(tempfile.gettempdir(), "autoprueba.txt")
        with open(ruta, "w", encoding="utf-8") as f:
            f.write(texto + "\n")
    if con_ventana:
        import tkinter as tk
        from tkinter import messagebox
        raiz = tk.Tk()
        raiz.withdraw()
        messagebox.showinfo("Autoprueba del Validador", f"{texto}\n\nSe guardó en: {ruta}")
        raiz.destroy()
    return 1 if " 0 falla(s)" not in texto else 0


if __name__ == "__main__":
    sys.exit(main(con_ventana="--sin-ventana" not in sys.argv))
