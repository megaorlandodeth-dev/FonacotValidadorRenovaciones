#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Convierte GUIA_DE_USO.md en GUIA_DE_USO.pdf para repartirla con el ejecutable (los analistas no tienen lector de .md).
El PDF lo imprime Microsoft Edge, que ya trae Windows 10/11. Lo llama Construir_Ejecutable.bat; también se puede correr
solo: python guia_pdf.py
"""

import os
import subprocess
import sys
import tempfile

import markdown

AQUI = os.path.dirname(os.path.abspath(__file__))
ORIGEN = os.path.join(AQUI, "GUIA_DE_USO.md")
DESTINO = os.path.join(AQUI, "GUIA_DE_USO.pdf")
EDGE = [os.path.join(os.environ.get(var, ""), r"Microsoft\Edge\Application\msedge.exe")
        for var in ("ProgramFiles(x86)", "ProgramFiles", "LOCALAPPDATA")]

ESTILO = """
@page { size: Letter; margin: 18mm 16mm; }
body { font-family: "Segoe UI", Arial, sans-serif; font-size: 10.5pt; line-height: 1.45; color: #1f2328; }
h1 { font-size: 20pt; color: #0b3d91; border-bottom: 2px solid #0b3d91; padding-bottom: 4px; margin-top: 0; }
h2 { font-size: 14pt; color: #0b3d91; margin-top: 22px; page-break-after: avoid; }
h3 { font-size: 11.5pt; margin-top: 16px; page-break-after: avoid; }
h2:not(:first-of-type) { page-break-before: always; }
table { border-collapse: collapse; width: 100%; margin: 8px 0; font-size: 9.5pt; page-break-inside: avoid; }
th, td { border: 1px solid #c8ccd1; padding: 4px 6px; vertical-align: top; text-align: left; }
th { background: #eef2f8; }
code { font-family: Consolas, monospace; font-size: 9pt; background: #f2f3f5; padding: 0 3px; border-radius: 3px; }
pre { background: #f2f3f5; padding: 6px 8px; border-radius: 4px; }
pre code { padding: 0; }
hr { display: none; }
a { color: #0b3d91; text-decoration: none; }
li { margin: 2px 0; }
"""


def html_de_la_guia() -> str:
    with open(ORIGEN, encoding="utf-8") as f:
        cuerpo = markdown.markdown(f.read(), extensions=["tables", "fenced_code", "toc"])
    return (f'<!doctype html><html lang="es"><head><meta charset="utf-8"><title>Guía del Validador de Renovaciones'
            f'</title><style>{ESTILO}</style></head><body>{cuerpo}</body></html>')


def main() -> int:
    edge = next((e for e in EDGE if os.path.isfile(e)), None)
    if not edge:
        print("No se encontró Microsoft Edge para imprimir la guía en PDF.")
        return 1
    with tempfile.TemporaryDirectory() as tmp:
        html = os.path.join(tmp, "guia.html")
        with open(html, "w", encoding="utf-8") as f:
            f.write(html_de_la_guia())
        if os.path.exists(DESTINO):
            os.remove(DESTINO)
        subprocess.run([edge, "--headless", "--disable-gpu", "--no-pdf-header-footer", f"--user-data-dir={tmp}",
                        f"--print-to-pdf={DESTINO}", "file:///" + html.replace("\\", "/")],
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=120)
    if not os.path.isfile(DESTINO):
        print("Edge no generó el PDF de la guía.")
        return 1
    print(f"Guía en PDF: {DESTINO}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
