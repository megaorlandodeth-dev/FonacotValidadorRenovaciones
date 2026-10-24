#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Lectura de capturas de pantalla (OCR) para los incidentes que en vez de escribir Cliente y Crédito mandan una
imagen (pantalla «Consulta de Crédito», resultado de SQL, etc.).

Usa el reconocimiento de texto que ya trae Windows 10/11 (Windows.Media.Ocr) llamado desde PowerShell: no hay
que instalar nada. Si el equipo no lo permite, la app sigue igual y avisa que la imagen no se pudo leer.
"""

import base64
import hashlib
import os
import re
import subprocess
import tempfile
import unicodedata

EXT_IMAGEN = ('.png', '.jpg', '.jpeg', '.bmp', '.gif', '.tif', '.tiff')
TAM_MIN_IMAGEN = 6 * 1024        # logos, iconos y firmas pesan menos
MAX_IMAGENES_POR_CORREO = 4
TIEMPO_MAX_OCR = 90              # segundos por llamada

_SEP = "\x1f"
_CACHE = {}                      # sha1 de la imagen -> líneas leídas (no se vuelve a leer en cada actualización)

# Imprime por cada imagen "IMG<tab>ruta", luego "LIN<tab>palabra<tab>palabra…" (palabra = texto␟x␟y␟ancho␟alto)
_SCRIPT_PS = r"""
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
try {
  Add-Type -AssemblyName System.Runtime.WindowsRuntime
  $null = [Windows.Storage.StorageFile, Windows.Storage, ContentType = WindowsRuntime]
  $null = [Windows.Media.Ocr.OcrEngine, Windows.Foundation, ContentType = WindowsRuntime]
  $null = [Windows.Graphics.Imaging.BitmapDecoder, Windows.Graphics, ContentType = WindowsRuntime]
  $asTask = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
    $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and
    $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' })[0]
  function Await($op, [Type]$t) { $task = $asTask.MakeGenericMethod($t).Invoke($null, @($op)); $null = $task.Wait(-1); $task.Result }
  $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages()
  if ($null -eq $engine) {
    $lang = [Windows.Media.Ocr.OcrEngine]::AvailableRecognizerLanguages | Select-Object -First 1
    if ($null -ne $lang) { $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage($lang) }
  }
  if ($null -eq $engine) { "ERR`tWindows no tiene instalado ningún idioma de reconocimiento de texto (OCR)"; exit }
} catch { "ERR`t" + $_.Exception.Message; exit }
$sep = [char]31
foreach ($path in Get-Content -LiteralPath '__LISTA__' -Encoding UTF8) {
  if (-not $path) { continue }
  "IMG`t" + $path
  try {
    $file = Await ([Windows.Storage.StorageFile]::GetFileFromPathAsync($path)) ([Windows.Storage.StorageFile])
    $stream = Await ($file.OpenAsync([Windows.Storage.FileAccessMode]::Read)) ([Windows.Storage.Streams.IRandomAccessStream])
    $decoder = Await ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) ([Windows.Graphics.Imaging.BitmapDecoder])
    $bitmap = Await ($decoder.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])
    $result = Await ($engine.RecognizeAsync($bitmap)) ([Windows.Media.Ocr.OcrResult])
    foreach ($l in $result.Lines) {
      "LIN`t" + (($l.Words | ForEach-Object {
        $r = $_.BoundingRect
        $_.Text + $sep + [int]$r.X + $sep + [int]$r.Y + $sep + [int]$r.Width + $sep + [int]$r.Height }) -join "`t")
    }
    $stream.Dispose()
  } catch { "ERR`t" + $_.Exception.Message }
}
"""


class OCRNoDisponible(Exception):
    """El equipo no permite usar el OCR de Windows (sin idioma, PowerShell bloqueado, etc.)."""


# ======================= Lectura de las imágenes =======================

# El OCR de Windows lee mejor a un tamaño medio: las capturas chicas se agrandan y las fotos grandes se reducen.
# Se lee a dos tamaños para comparar: si no coinciden, la lectura queda como dudosa.
LADOS_OCR = (1800, 2400)


def _preparar(ruta: str, carpeta: str) -> list:
    """Copias listas para el OCR (orientación de la foto corregida, en gris, a cada tamaño). Sin Pillow, la original."""
    try:
        from PIL import Image, ImageOps
    except ImportError:
        return [ruta]
    try:
        with Image.open(ruta) as im:
            im = ImageOps.exif_transpose(im).convert("L")
            base = os.path.join(carpeta, hashlib.sha1(ruta.encode('utf-8')).hexdigest()[:12])
            versiones = []
            for lado in LADOS_OCR:
                factor = lado / max(max(im.size), 1)
                copia = im.resize((max(1, int(im.width * factor)), max(1, int(im.height * factor))), Image.LANCZOS)
                versiones.append(f"{base}_{lado}.png")
                copia.save(versiones[-1])
            return versiones
    except Exception:
        return [ruta]


def _ejecutar_ocr(rutas: list) -> dict:
    """Llama al OCR de Windows una sola vez para todas las rutas. Devuelve {ruta: [[(texto, x, y, w, h), …], …]}."""
    with tempfile.TemporaryDirectory(prefix="fonacot_ocr_") as tmp:
        lista = os.path.join(tmp, "lista.txt")
        with open(lista, "w", encoding="utf-8") as f:
            f.write("\n".join(rutas))
        script = _SCRIPT_PS.replace("__LISTA__", lista.replace("'", "''"))
        # -EncodedCommand: no depende de la política de ejecución de scripts .ps1
        cifrado = base64.b64encode(script.encode("utf-16-le")).decode("ascii")
        try:
            proc = subprocess.run(
                ["powershell.exe", "-NoProfile", "-NonInteractive", "-EncodedCommand", cifrado],
                capture_output=True, timeout=TIEMPO_MAX_OCR,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0)
            )
        except (OSError, subprocess.TimeoutExpired) as e:
            raise OCRNoDisponible(f"No se pudo ejecutar el OCR de Windows: {e}")
    salida = proc.stdout.decode("utf-8", errors="replace")
    resultado, actual = {}, None
    for renglon in salida.splitlines():
        tipo, _, resto = renglon.partition("\t")
        if tipo == "ERR" and actual is None:
            raise OCRNoDisponible(resto.strip() or "El OCR de Windows no respondió")
        if tipo == "IMG":
            actual = resto.strip()
            resultado[actual] = []
        elif tipo == "LIN" and actual is not None:
            palabras = []
            for p in resto.split("\t"):
                partes = p.split(_SEP)
                if len(partes) == 5:
                    try:
                        palabras.append((partes[0], *(int(v) for v in partes[1:])))
                    except ValueError:
                        continue
            if palabras:
                resultado[actual].append(palabras)
    if not resultado and proc.returncode != 0:
        raise OCRNoDisponible(proc.stderr.decode("utf-8", errors="replace").strip()[:200] or "El OCR de Windows falló")
    return resultado


def leer_imagenes(rutas: list) -> dict:
    """
    {ruta original: [líneas de cada versión]} de cada imagen. Usa la caché: una imagen que ya se leyó no se vuelve
    a leer en la siguiente actualización.
    """
    resultado, pendientes = {}, {}
    with tempfile.TemporaryDirectory(prefix="fonacot_img_") as tmp:
        for ruta in rutas:
            try:
                with open(ruta, "rb") as f:
                    huella = hashlib.sha1(f.read()).hexdigest()
            except OSError:
                continue
            if huella in _CACHE:
                resultado[ruta] = _CACHE[huella]
            else:
                pendientes[ruta] = (huella, _preparar(ruta, tmp))
        if pendientes:
            leidas = _ejecutar_ocr([v for _, versiones in pendientes.values() for v in versiones])
            for ruta, (huella, versiones) in pendientes.items():
                lecturas = [leidas.get(v, []) for v in versiones]
                _CACHE[huella] = lecturas
                resultado[ruta] = lecturas
    return resultado


# ======================= Imágenes del correo =======================

def guardar_imagenes_del_correo(msg, carpeta: str) -> list:
    """Guarda en `carpeta` las imágenes del correo (adjuntas o pegadas en el cuerpo). Ignora logos y firmas."""
    rutas = []
    try:
        adjuntos = msg.Attachments
        total = adjuntos.Count
    except Exception:
        total = 0
    for i in range(1, total + 1):
        if len(rutas) >= MAX_IMAGENES_POR_CORREO:
            break
        try:
            att = adjuntos.Item(i)
            nombre = str(getattr(att, 'FileName', '') or '')
            if os.path.splitext(nombre)[1].lower() not in EXT_IMAGEN or int(getattr(att, 'Size', 0) or 0) < TAM_MIN_IMAGEN:
                continue
            ruta = os.path.join(carpeta, f"adj{i}_{re.sub(r'[^A-Za-z0-9._-]', '_', nombre)}")
            att.SaveAsFile(ruta)
            rutas.append(ruta)
        except Exception:
            continue
    # Imágenes pegadas dentro del HTML (data:image/…;base64)
    try:
        html = str(getattr(msg, 'HTMLBody', '') or '')
    except Exception:
        html = ''
    for n, m in enumerate(re.finditer(r'src=["\']data:image/(png|jpe?g|gif|bmp);base64,([A-Za-z0-9+/=\s]+)["\']', html, re.I)):
        if len(rutas) >= MAX_IMAGENES_POR_CORREO:
            break
        try:
            datos = base64.b64decode(re.sub(r'\s+', '', m.group(2)))
        except Exception:
            continue
        if len(datos) < TAM_MIN_IMAGEN:
            continue
        ruta = os.path.join(carpeta, f"pegada{n}.{m.group(1).lower().replace('jpeg', 'jpg')}")
        with open(ruta, "wb") as f:
            f.write(datos)
        rutas.append(ruta)
    return rutas


def leer_correo(msg) -> tuple:
    """
    Lee las imágenes del correo de un incidente. Devuelve (lecturas, número de imágenes): una lectura
    {'cliente', 'credito', 'imagen', 'texto', 'dudosa'} por imagen en la que se encontró algún dato.
    Lanza OCRNoDisponible si el OCR de Windows no se puede usar en este equipo.
    """
    with tempfile.TemporaryDirectory(prefix="fonacot_correo_") as tmp:
        rutas = guardar_imagenes_del_correo(msg, tmp)
        if not rutas:
            return [], 0
        leidas = leer_imagenes(rutas)
        lecturas = []
        for ruta in rutas:
            datos = extraer_de_versiones(leidas.get(ruta, []))
            if datos['cliente'] or datos['credito']:
                datos['imagen'] = os.path.basename(ruta).split("_", 1)[-1]
                lecturas.append(datos)
        return lecturas, len(rutas)


# ======================= Cliente y crédito a partir del texto leído =======================

def _norm(texto: str) -> str:
    t = unicodedata.normalize('NFKD', texto).encode('ascii', 'ignore').decode('ascii').lower()
    return re.sub(r'[^a-z_]', '', t)


def _numero(texto: str) -> str:
    """Dígitos de una palabra que es solo un número (sin decimales, fechas ni montos). '' si no lo es."""
    t = texto.strip("()[]'\"«»:;.")
    t = t.rstrip(",")
    return t if re.fullmatch(r'\d{3,9}', t) else ''


def _es_etiqueta_cliente(n: str) -> bool:
    return n in ('cliente', 'clientes', 'cte', 'cuenta', 'nocliente', 'numcliente', 'nocte') or n.startswith('client')


def _es_etiqueta_credito(n: str) -> bool:
    return (n in ('credito', 'cred', 'nocredito', 'credrto', 'credlto', 'creditohijo', 'renovacion', 'norenovacion')
            or n.startswith('credito') or n.endswith('renovacion'))


def extraer_de_versiones(versiones: list) -> dict:
    """
    Junta lo leído en cada tamaño de la misma imagen: gana el valor que más se repite. Si las lecturas no
    coinciden queda `dudosa` (pasa con fotos tomadas al monitor, donde un 6 se lee como 5).
    """
    datos = [extraer_cliente_credito(v) for v in versiones if v]
    if not datos:
        return {'cliente': '', 'credito': '', 'texto': '', 'dudosa': False}
    elegido = {'texto': max((d['texto'] for d in datos), key=len), 'dudosa': False}
    for campo in ('cliente', 'credito'):
        valores = [d[campo] for d in datos if d[campo]]
        elegido[campo] = max(valores, key=valores.count) if valores else ''
        if len(set(valores)) > 1:  # que un tamaño no lo lea no es contradicción; que lea otro número sí
            elegido['dudosa'] = True
    return elegido


def extraer_cliente_credito(lineas: list) -> dict:
    """
    Busca No. de cliente y crédito en el texto leído de una imagen. `lineas`: [[(texto, x, y, ancho, alto), …], …].
    1) Junto a su etiqueta: el número a la derecha («No. Cliente  4363419», «WHERE CUENTA IN ('1184282')») o debajo
       (encabezados de columna como CUENTA / NO_RENOVACION en un resultado de SQL, o «Datos del Crédito»).
    2) Sin etiqueta legible: un número de 6 a 8 dígitos seguido en el mismo renglón por un nombre en mayúsculas es
       el cliente (pantalla «Consulta de Crédito», donde las etiquetas son texto blanco sobre gris).
    """
    palabras = [p for ln in lineas for p in ln]
    numeros = [(p, _numero(p[0])) for p in palabras]
    numeros = [(p, n) for p, n in numeros if n]
    texto = "\n".join(" ".join(p[0] for p in ln) for ln in lineas)

    def mismo_renglon(a, b):
        ca, cb = a[2] + a[4] / 2, b[2] + b[4] / 2
        return abs(ca - cb) <= max(a[4], b[4]) * 0.7

    def valor_de(etiqueta, digitos_min, digitos_max):
        alto = max(etiqueta[4], 8)
        candidatos = [(p, n) for p, n in numeros if digitos_min <= len(n) <= digitos_max]
        derecha = [(p[1] - (etiqueta[1] + etiqueta[3]), n) for p, n in candidatos
                   if mismo_renglon(etiqueta, p) and p[1] >= etiqueta[1] + etiqueta[3] - 2
                   and p[1] - (etiqueta[1] + etiqueta[3]) <= alto * 25]
        if derecha:
            return min(derecha)[1]
        izq, der = etiqueta[1] - alto * 3, etiqueta[1] + etiqueta[3] + alto * 10
        abajo = [(p[2] - etiqueta[2], n) for p, n in candidatos
                 if 0 < p[2] - etiqueta[2] <= alto * 3.5 and p[1] + p[3] >= izq and p[1] <= der]
        return min(abajo)[1] if abajo else ''

    cliente = credito = ''
    for p in palabras:
        n = _norm(p[0])
        if not cliente and _es_etiqueta_cliente(n):
            cliente = valor_de(p, 5, 9)
        elif not credito and _es_etiqueta_credito(n):
            credito = valor_de(p, 3, 6)
    if not cliente:
        for p, n in numeros:
            if 6 <= len(n) <= 8:
                nombre = [q for q in palabras if q is not p and mismo_renglon(p, q) and q[1] > p[1]
                          and q[1] - (p[1] + p[3]) <= max(p[4], 8) * 30 and re.fullmatch(r"[A-ZÁÉÍÓÚÑ]{2,}", q[0])]
                if len(nombre) >= 2:
                    cliente = n
                    break
    if not (cliente and credito):
        # Texto corrido («Cliente: 4363419 Crédito: 17626») cuando las palabras no traen posición útil
        plano = unicodedata.normalize('NFKD', texto).encode('ascii', 'ignore').decode('ascii')
        if not cliente:
            m = re.search(r'(?:cliente|cuenta)\W{0,6}(\d{5,9})\b', plano, re.I)
            cliente = m.group(1) if m else ''
        if not credito:
            m = re.search(r'(?:credito|renovacion)\W{0,6}(\d{3,6})\b', plano, re.I)
            credito = m.group(1) if m else ''
    return {'cliente': cliente, 'credito': credito, 'texto': texto}
