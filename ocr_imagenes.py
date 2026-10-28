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
import html as html_mod
import os
import re
import subprocess
import tempfile
import unicodedata
from datetime import date, timedelta

EXT_IMAGEN = ('.png', '.jpg', '.jpeg', '.bmp', '.gif', '.tif', '.tiff')
TAM_MIN_IMAGEN = 6 * 1024        # logos, iconos y firmas pesan menos
MAX_IMAGENES_POR_CORREO = 4
MAX_PAGINAS_PDF = 2              # del estado de cuenta basta la primera hoja
TAM_MIN_PDF = 1024
TIEMPO_MAX_OCR = 90              # segundos por llamada
# Por ruta completa: en equipos con el PATH cambiado (políticas de TI) «powershell.exe» solo no se encuentra
_POWERSHELL = next((p for p in [os.path.join(os.environ.get("SystemRoot", r"C:\Windows"), "System32",
                                             "WindowsPowerShell", "v1.0", "powershell.exe")] if os.path.isfile(p)),
                   "powershell.exe")

_SEP = "\x1f"
_CACHE = {}                      # sha1 de la imagen -> líneas leídas (no se vuelve a leer en cada actualización)
_CACHE_PDF = {}                  # sha1 del PDF -> imágenes de sus páginas

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


# Convierte cada PDF de la lista («ruta<tab>prefijo») en imágenes PNG de sus primeras páginas con el lector de PDF que
# trae Windows (Windows.Data.Pdf): imprime "PNG<tab>ruta" por página o "ERR<tab>motivo".
_SCRIPT_PDF_PS = r"""
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
try {
  Add-Type -AssemblyName System.Runtime.WindowsRuntime
  $null = [Windows.Storage.StorageFile, Windows.Storage, ContentType = WindowsRuntime]
  $null = [Windows.Data.Pdf.PdfDocument, Windows.Data.Pdf, ContentType = WindowsRuntime]
  $null = [Windows.Storage.Streams.InMemoryRandomAccessStream, Windows.Storage.Streams, ContentType = WindowsRuntime]
  $metodos = [System.WindowsRuntimeSystemExtensions].GetMethods()
  $asTask = ($metodos | Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and
    $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' })[0]
  $asTaskA = ($metodos | Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and
    $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncAction' })[0]
  function Await($op, [Type]$t) { $task = $asTask.MakeGenericMethod($t).Invoke($null, @($op)); $null = $task.Wait(-1); $task.Result }
  function AwaitA($op) { $null = $asTaskA.Invoke($null, @($op)).Wait(-1) }
} catch { "ERR`t" + $_.Exception.Message; exit }
foreach ($renglon in Get-Content -LiteralPath '__LISTA__' -Encoding UTF8) {
  if (-not $renglon) { continue }
  $pdf, $prefijo = $renglon -split "`t", 2
  try {
    $file = Await ([Windows.Storage.StorageFile]::GetFileFromPathAsync($pdf)) ([Windows.Storage.StorageFile])
    $doc = Await ([Windows.Data.Pdf.PdfDocument]::LoadFromFileAsync($file)) ([Windows.Data.Pdf.PdfDocument])
    $paginas = [Math]::Min([int]$doc.PageCount, __PAGINAS__)
    for ($i = 0; $i -lt $paginas; $i++) {
      $page = $doc.GetPage($i)
      $opt = New-Object Windows.Data.Pdf.PdfPageRenderOptions
      $opt.DestinationWidth = [uint32]2400
      $ms = New-Object Windows.Storage.Streams.InMemoryRandomAccessStream
      AwaitA ($page.RenderToStreamAsync($ms, $opt))
      $salida = $prefijo + "_pag" + ($i + 1) + ".png"
      $net = [System.IO.WindowsRuntimeStreamExtensions]::AsStreamForRead($ms.GetInputStreamAt(0))
      $fs = [System.IO.File]::Create($salida)
      $net.CopyTo($fs)
      $fs.Close()
      $ms.Dispose()
      $page.Dispose()
      "PNG`t" + $salida
    }
  } catch { "ERR`t" + $_.Exception.Message }
}
"""


class OCRNoDisponible(Exception):
    """El equipo no permite usar el OCR de Windows (sin idioma, PowerShell bloqueado, etc.)."""


# ======================= Lectura de las imágenes =======================

# El OCR de Windows lee mejor a un tamaño medio: las capturas chicas se agrandan y las fotos grandes se reducen.
# Se lee a dos tamaños para comparar: si no coinciden, la lectura queda como dudosa.
LADOS_OCR = (1800, 2400)


def _preparar(ruta: str, carpeta: str, giro: int = 0) -> list:
    """
    Copias listas para el OCR (orientación de la foto corregida, en gris, a cada tamaño). Sin Pillow, la original.
    `giro`: grados para enderezar una foto tomada de lado (el OCR de Windows solo lee texto derecho).
    """
    try:
        from PIL import Image, ImageOps
    except ImportError:
        return [ruta] if not giro else []
    try:
        with Image.open(ruta) as im:
            im = ImageOps.exif_transpose(im).convert("L")
            if giro:
                im = im.rotate(giro, expand=True)
            base = os.path.join(carpeta, hashlib.sha1(f"{ruta}|{giro}".encode('utf-8')).hexdigest()[:12])
            versiones = []
            for lado in LADOS_OCR:
                factor = lado / max(max(im.size), 1)
                copia = im.resize((max(1, int(im.width * factor)), max(1, int(im.height * factor))), Image.LANCZOS)
                versiones.append(f"{base}_{lado}.png")
                copia.save(versiones[-1])
            return versiones
    except Exception:
        return [ruta] if not giro else []


MIN_PALABRAS_DERECHA = 25        # con menos, la foto pudo tomarse de lado: se prueba girándola


def _palabras_leidas(versiones: list) -> int:
    return max((sum(1 for ln in v for p in ln if len(p[0]) >= 3) for v in versiones), default=0)


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
                [_POWERSHELL, "-NoProfile", "-NonInteractive", "-EncodedCommand", cifrado],
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
            de_lado = {}
            for ruta, (huella, versiones) in pendientes.items():
                lecturas = [leidas.get(v, []) for v in versiones]
                resultado[ruta] = lecturas
                if _palabras_leidas(lecturas) < MIN_PALABRAS_DERECHA:
                    for giro in (90, 270):
                        de_lado[(ruta, giro)] = _preparar(ruta, tmp, giro)
            if de_lado:
                leidas = _ejecutar_ocr([v for versiones in de_lado.values() for v in versiones])
                for (ruta, giro), versiones in de_lado.items():
                    lecturas = [leidas.get(v, []) for v in versiones]
                    if _palabras_leidas(lecturas) > _palabras_leidas(resultado[ruta]):
                        resultado[ruta] = lecturas
            _leer_cajas_de_etiquetas(pendientes, resultado, tmp)
            for ruta, (huella, _) in pendientes.items():
                _CACHE[huella] = resultado[ruta]
    return resultado


BLANCO_CAJA = 245                # las cajas de los campos son blancas; el fondo de las pantallas, gris (≤ 240)
MAX_CAJAS_POR_LECTURA = 6


def _caja_a_la_derecha(im, etiqueta) -> tuple:
    """
    (x0, y0, x1, y1) de la primera caja blanca a la derecha de una etiqueta (los campos de «Consulta de Crédito»);
    None si no hay. Una columna es de la caja si tiene blanco en su franja: el borde, el fondo gris y el botón «?» no.
    """
    _, x, y, w, h = etiqueta
    arriba, abajo, izq = max(0, int(y - h * 0.5)), min(im.height, int(y + h * 1.5)), int(x + w)
    if abajo - arriba < 4 or izq >= im.width - 4:
        return None
    banda = im.crop((izq, arriba, im.width, abajo))
    pix, (ancho, alto) = banda.load(), banda.size
    minimo = max(2, int(alto * 0.15))
    inicio = None
    for i in range(ancho + 1):
        blanca = i < ancho and sum(1 for j in range(alto) if pix[i, j] >= BLANCO_CAJA) >= minimo
        if blanca and inicio is None:
            inicio = i
        elif not blanca and inicio is not None:
            if i - inicio >= h * 1.5:
                return izq + inicio, arriba, izq + i, abajo
            inicio = None
    return None


def _leer_cajas_de_etiquetas(pendientes: dict, resultado: dict, carpeta: str):
    """
    El OCR de Windows a veces se salta un renglón entero de la pantalla «Consulta de Crédito»: el del cliente, con el
    número, el botón «?» y el nombre (I-363027). Si una etiqueta Cliente o Crédito se leyó pero su número no, se lee
    sola la caja que tiene a la derecha y sus palabras se agregan a esa lectura, en su lugar.
    """
    try:
        from PIL import Image
    except ImportError:
        return
    recortes = {}
    for ruta, (_, versiones) in pendientes.items():
        if len(resultado.get(ruta, [])) != len(versiones):
            continue
        for k, (version, lineas) in enumerate(zip(versiones, resultado[ruta])):
            if not lineas or version == ruta:
                continue  # sin Pillow la «versión» es la imagen original: no se recorta
            datos = extraer_cliente_credito(lineas)
            etiquetas = [p for ln in lineas for p in ln
                         if (not datos['cliente'] and _es_etiqueta_cliente(_norm(p[0])))
                         or (not datos['credito'] and _es_etiqueta_credito(_norm(p[0])))][:MAX_CAJAS_POR_LECTURA]
            if not etiquetas:
                continue
            try:
                with Image.open(version) as im:
                    im = im.convert("L")
                    for n, etiqueta in enumerate(etiquetas):
                        caja = _caja_a_la_derecha(im, etiqueta)
                        if not caja:
                            continue
                        margen = 20
                        recorte = Image.new("L", (caja[2] - caja[0] + margen * 2, caja[3] - caja[1] + margen * 2), 255)
                        recorte.paste(im.crop(caja), (margen, margen))
                        destino = f"{version[:-4]}_caja{n}.png"
                        recorte.save(destino)
                        recortes[destino] = (ruta, k, caja[0] - margen, caja[1] - margen)
            except Exception:
                continue
    if not recortes:
        return
    try:
        leidas = _ejecutar_ocr(list(recortes))
    except OCRNoDisponible:
        return
    for destino, (ruta, k, dx, dy) in recortes.items():
        for ln in leidas.get(destino, []):
            resultado[ruta][k].append([(t, x + dx, y + dy, w, h) for t, x, y, w, h in ln])


def pdf_a_imagenes(pdfs: list, notas: list = None) -> list:
    """
    Convierte en imágenes las primeras páginas de cada PDF (`pdfs`: [(ruta del pdf, prefijo de las imágenes)]) para
    leerlas con el OCR: así se leen el estado de cuenta o la consulta que la sucursal manda en PDF, sea texto o un
    escaneo. Un PDF que ya se convirtió no se vuelve a convertir en la siguiente actualización.
    """
    rutas, pendientes = [], []
    for pdf, prefijo in pdfs:
        try:
            with open(pdf, "rb") as f:
                huella = hashlib.sha1(f.read()).hexdigest()
        except OSError:
            continue
        if huella in _CACHE_PDF:
            for n, datos in enumerate(_CACHE_PDF[huella], 1):
                with open(f"{prefijo}_pag{n}.png", "wb") as f:
                    f.write(datos)
                rutas.append(f"{prefijo}_pag{n}.png")
        else:
            pendientes.append((pdf, prefijo, huella))
    if not pendientes:
        return rutas
    with tempfile.TemporaryDirectory(prefix="fonacot_pdf_") as tmp:
        lista = os.path.join(tmp, "lista.txt")
        with open(lista, "w", encoding="utf-8") as f:
            f.write("\n".join(f"{pdf}\t{prefijo}" for pdf, prefijo, _ in pendientes))
        script = _SCRIPT_PDF_PS.replace("__LISTA__", lista.replace("'", "''")).replace("__PAGINAS__", str(MAX_PAGINAS_PDF))
        cifrado = base64.b64encode(script.encode("utf-16-le")).decode("ascii")
        try:
            proc = subprocess.run(
                [_POWERSHELL, "-NoProfile", "-NonInteractive", "-EncodedCommand", cifrado],
                capture_output=True, timeout=TIEMPO_MAX_OCR,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0)
            )
        except (OSError, subprocess.TimeoutExpired) as e:
            if notas is not None:
                notas.append(f"no se pudo abrir el PDF adjunto ({type(e).__name__})")
            return rutas
    hechas = []
    for renglon in proc.stdout.decode("utf-8", errors="replace").splitlines():
        tipo, _, resto = renglon.partition("\t")
        if tipo == "PNG" and os.path.exists(resto.strip()):
            hechas.append(resto.strip())
        elif tipo == "ERR" and notas is not None:
            notas.append(f"no se pudo abrir el PDF adjunto ({resto.strip()[:80]})")
    for pdf, prefijo, huella in pendientes:
        paginas = [r for r in hechas if r.startswith(prefijo + "_pag")]
        if paginas:
            datos = []
            for r in paginas:
                with open(r, "rb") as f:
                    datos.append(f.read())
            _CACHE_PDF[huella] = datos
        rutas += paginas
    return rutas


def nombre_mostrado(ruta: str) -> str:
    """«adj2_EDO CTA.pdf_pag1.png» -> «EDO CTA.pdf (pág. 1)»; «adj1_captura.png» -> «captura.png»."""
    nombre = os.path.basename(ruta).split("_", 1)[-1]
    return re.sub(r'\.pdf_pag(\d+)\.png$', r'.pdf (pág. \1)', nombre, flags=re.I)


# ======================= Imágenes del correo =======================

# Imágenes ligadas en el correo (las que iTop deja en su servidor): solo de estos servidores, nunca de fuera
SERVIDORES_IMAGENES = ("fonacot.gob.mx", "itop")
TIEMPO_MAX_DESCARGA = 8
TIPOS_IMAGEN = {"image/png": "png", "image/jpeg": "jpg", "image/jpg": "jpg", "image/gif": "gif", "image/bmp": "bmp"}


def _descargar_imagen(url: str, ruta_sin_ext: str, notas: list = None) -> str:
    """Baja una imagen ligada en el correo (iTop). La ruta guardada o '' (no es imagen, pide sesión, sin red…)."""
    import urllib.request
    from urllib.parse import urlparse
    host = (urlparse(url).hostname or "").lower()
    if not any(s in host for s in SERVIDORES_IMAGENES):
        return ""
    pedido = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    try:
        with urllib.request.urlopen(pedido, timeout=TIEMPO_MAX_DESCARGA) as resp:
            tipo = (resp.headers.get_content_type() or "").lower()
            datos = resp.read(15 * 1024 * 1024)
    except Exception as e:
        if notas is not None:
            notas.append(f"no se pudo bajar la imagen de {host} ("
                         + ("certificado no válido" if "CERTIFICATE" in str(e).upper() else type(e).__name__) + ")")
        return ""
    if tipo not in TIPOS_IMAGEN:
        if notas is not None:
            notas.append(f"la imagen de {host} pide iniciar sesión en iTop" if "html" in tipo
                         else f"{host} no devolvió una imagen ({tipo or 'sin tipo'})")
        return ""
    if len(datos) < TAM_MIN_IMAGEN:
        return ""
    ruta = f"{ruta_sin_ext}.{TIPOS_IMAGEN[tipo]}"
    with open(ruta, "wb") as f:
        f.write(datos)
    return ruta


def guardar_imagenes_del_correo(msg, carpeta: str, notas: list = None, _nivel: int = 0) -> list:
    """
    Guarda en `carpeta` las imágenes del correo: adjuntas, pegadas en el cuerpo, ligadas al servidor de iTop, las
    páginas de los PDF adjuntos (estado de cuenta) y las de un correo reenviado como adjunto (.msg). Ignora logos
    y firmas.
    """
    rutas, pdfs = [], []
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
            if _nivel == 0 and (int(getattr(att, 'Type', 1) or 1) == 5 or nombre.lower().endswith('.msg')):
                # Correo de la sucursal reenviado como adjunto: sus imágenes también cuentan
                ruta_msg = os.path.join(carpeta, f"adj{i}.msg")
                att.SaveAsFile(ruta_msg)
                interno = msg.Session.OpenSharedItem(ruta_msg)
                try:
                    sub = os.path.join(carpeta, f"msg{i}")
                    os.makedirs(sub, exist_ok=True)
                    rutas += guardar_imagenes_del_correo(interno, sub, notas, _nivel + 1)
                finally:
                    try:
                        interno.Close(1)  # olDiscard
                    except Exception:
                        pass
                continue
            if nombre.lower().endswith('.pdf') and int(getattr(att, 'Size', 0) or 0) >= TAM_MIN_PDF:
                ruta_pdf = os.path.join(carpeta, f"adj{i}.pdf")
                att.SaveAsFile(ruta_pdf)
                pdfs.append((ruta_pdf, os.path.join(carpeta, f"adj{i}_{re.sub(r'[^A-Za-z0-9._-]', '_', nombre)}")))
                continue
            if os.path.splitext(nombre)[1].lower() not in EXT_IMAGEN or int(getattr(att, 'Size', 0) or 0) < TAM_MIN_IMAGEN:
                continue
            ruta = os.path.join(carpeta, f"adj{i}_{re.sub(r'[^A-Za-z0-9._-]', '_', nombre)}")
            att.SaveAsFile(ruta)
            rutas.append(ruta)
        except Exception:
            continue
    if pdfs and len(rutas) < MAX_IMAGENES_POR_CORREO:
        rutas += pdf_a_imagenes(pdfs, notas)[:MAX_IMAGENES_POR_CORREO - len(rutas)]
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
    # Ligadas al servidor (iTop guarda ahí las imágenes de la descripción del incidente)
    vistas = set()
    for n, m in enumerate(re.finditer(r'<img\b[^>]*?\bsrc=["\'](https?://[^"\']+)["\']', html, re.I)):
        url = html_mod.unescape(m.group(1))
        if len(rutas) >= MAX_IMAGENES_POR_CORREO or url in vistas:
            continue
        vistas.add(url)
        ruta = _descargar_imagen(url, os.path.join(carpeta, f"ligada{n}"), notas)
        if ruta:
            rutas.append(ruta)
    return rutas[:MAX_IMAGENES_POR_CORREO]


def leer_correo(msg, notas: list = None, ref: date = None) -> tuple:
    """
    Lee las imágenes (y los PDF) del correo de un incidente. Devuelve (lecturas, número de imágenes): una lectura
    {'cliente', 'credito', 'imagen', 'texto', 'dudosa'} por imagen en la que se encontró algún dato; de un estado de
    cuenta, una por cada renovación autorizada cerca de `ref` (la fecha del correo), con 'fuente' que lo explica.
    En `notas` deja lo que no se pudo bajar o abrir (para explicarlo en pantalla).
    Lanza OCRNoDisponible si el OCR de Windows no se puede usar en este equipo.
    """
    with tempfile.TemporaryDirectory(prefix="fonacot_correo_") as tmp:
        rutas = guardar_imagenes_del_correo(msg, tmp, notas)
        if not rutas:
            return [], 0
        leidas = leer_imagenes(rutas)
        lecturas = []
        for ruta in rutas:
            datos = extraer_de_versiones(leidas.get(ruta, []))
            datos['imagen'] = nombre_mostrado(ruta)
            if datos.get('edo_cuenta'):
                elegida = renovacion_del_estado(datos['edo_cuenta'], ref)
                if elegida['nota'] and notas is not None:
                    notas.append(elegida['nota'])
                for cred in elegida['creditos']:
                    lecturas.append(dict(datos, credito=cred, dudosa=datos['dudosa'] or elegida['dudosa'],
                                         fuente=f"del estado de cuenta «{datos['imagen']}» (autorizado el "
                                                f"{elegida['fecha']:%d/%m/%Y})"))
                if elegida['creditos'] or not datos['cliente']:
                    continue
                datos['credito'] = ''  # el primer crédito de la tabla no es la renovación
            if datos['cliente'] or datos['credito']:
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
        return {'cliente': '', 'credito': '', 'texto': '', 'dudosa': False, 'edo_cuenta': []}
    elegido = {'texto': max((d['texto'] for d in datos), key=len), 'dudosa': False,
               'edo_cuenta': [d['edo_cuenta'] for d in datos if d.get('edo_cuenta')]}
    if elegido['edo_cuenta']:
        for d in datos:
            d['credito'] = ''  # en un estado de cuenta el crédito sale de su tabla, no de la etiqueta
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
            m = re.search(r'(?:credito|renovacion)\W{0,6}(\d{3,7})\b', plano, re.I)
            credito = m.group(1) if m else ''
    res = {'cliente': cliente, 'credito': credito, 'texto': texto}
    edo = creditos_estado_de_cuenta(lineas)
    if edo:
        res['edo_cuenta'] = edo  # solo si es un estado de cuenta
    return res

# ======================= Estado de cuenta (FORM5061) =======================
# Trae al cliente y TODOS sus créditos (NO. CREDITO · PRODUCTO · FECHA AUTORIZADO · …): la renovación a cancelar es
# la que se autorizó en esos días (pagos 0/12). Lo de la etiqueta «NO. CREDITO» sería el primer crédito de la tabla.

DIAS_ESTADO_DE_CUENTA = 31       # la renovación se autorizó a lo más un mes antes del correo


def creditos_estado_de_cuenta(lineas: list) -> list:
    """
    [(crédito, fecha autorizado)] de los renglones de la tabla de un estado de cuenta: el primer número del renglón
    (3 a 6 dígitos), luego el producto (1 a 3 dígitos) y la primera fecha. [] si el texto no es un estado de cuenta.
    """
    palabras = [p for ln in lineas for p in ln]
    plano = _norm("".join(p[0] for p in palabras))
    if not ('decuenta' in plano or 'autoriz' in plano or ('producto' in plano and 'credito' in plano)):
        return []

    def mismo_renglon(a, b):
        ca, cb = a[2] + a[4] / 2, b[2] + b[4] / 2
        return abs(ca - cb) <= max(a[4], b[4]) * 0.6

    res = []
    for p in palabras:
        m = re.fullmatch(r'(\d{1,2})/(\d{1,2})/(\d{4})', p[0].strip(".,;:"))
        if not m:
            continue
        try:
            fecha = date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
        except ValueError:
            continue
        renglon = sorted((q for q in palabras if q is not p and mismo_renglon(p, q) and q[1] < p[1]), key=lambda q: q[1])
        if any(re.search(r'\d{1,2}/\d{1,2}/\d{4}', q[0]) for q in renglon):
            continue  # no es la primera fecha del renglón (FECHA EJERCIDO)
        nums = [_numero(q[0]) or (q[0] if re.fullmatch(r'\d{1,2}', q[0]) else '') for q in renglon]
        if len(nums) >= 2 and nums[0] and 3 <= len(nums[0]) <= 6 and nums[1] and len(nums[1]) <= 3 and all(
                re.fullmatch(r'\d+', q[0]) for q in renglon):
            res.append((nums[0], fecha))
    return list(dict.fromkeys(res))


def renovacion_del_estado(versiones_edo: list, ref: date = None) -> dict:
    """
    De lo leído de un estado de cuenta (una lista de renglones por cada tamaño leído) elige la renovación: los
    créditos con la fecha de autorización más reciente, si es de hasta un mes antes del correo (`ref`).
    {'creditos': [...], 'fecha': date|None, 'dudosa': bool, 'nota': str}
    """
    renglones = max(versiones_edo, key=len) if versiones_edo else []
    if not renglones:
        return {'creditos': [], 'fecha': None, 'dudosa': False, 'nota': ''}
    hasta = (ref or max(f for _, f in renglones)) + timedelta(days=1)
    validos = [(c, f) for c, f in renglones if f <= hasta]
    if not validos:
        return {'creditos': [], 'fecha': None, 'dudosa': False, 'nota': ''}
    fecha = max(f for _, f in validos)
    if ref and fecha < ref - timedelta(days=DIAS_ESTADO_DE_CUENTA):
        ultimo = next(c for c, f in validos if f == fecha)
        return {'creditos': [], 'fecha': None, 'dudosa': False,
                'nota': (f"el estado de cuenta no trae ningún crédito autorizado cerca del {ref:%d/%m/%Y} "
                         f"(el último: {ultimo} del {fecha:%d/%m/%Y})")}
    creditos = [c for c, f in validos if f == fecha]
    otras = [{c for c, f in v if f == fecha} for v in versiones_edo if v]
    dudosa = any(o and o != set(creditos) for o in otras)
    return {'creditos': creditos, 'fecha': fecha, 'dudosa': dudosa, 'nota': ''}

