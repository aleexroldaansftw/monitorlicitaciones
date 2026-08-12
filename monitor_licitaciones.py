# -*- coding: utf-8 -*-
"""
Monitor de licitaciones - e-Oficialía Hidalgo  (versión web)
============================================================
Revisa https://eoficialia.hidalgo.gob.mx/LICITACIONES/VISTAS/WebFrmLC004.aspx
y genera:
  - docs/index.html  -> tablero web (se publica con GitHub Pages)
  - data/licitaciones.xlsx -> Excel acumulado
  - data/estado.json -> registro interno

El WhatsApp es OPCIONAL: solo se envía si defines WHATSAPP_PHONE y
CALLMEBOT_APIKEY como variables de entorno. Si no, se omite sin error.
"""

import json
import os
import re
import sys
from datetime import datetime, date, timedelta
from pathlib import Path

import requests
from bs4 import BeautifulSoup

URL = "https://eoficialia.hidalgo.gob.mx/LICITACIONES/VISTAS/WebFrmLC004.aspx"
BASE = "https://eoficialia.hidalgo.gob.mx"

RAIZ = Path(__file__).parent
DATA_DIR = RAIZ / "data"
DOCS_DIR = RAIZ / "docs"
ESTADO_PATH = DATA_DIR / "estado.json"
EXCEL_PATH = DATA_DIR / "licitaciones.xlsx"
HTML_PATH = DOCS_DIR / "index.html"

DIAS_COMO_NUEVA = 3  # cuántos días una licitación conserva la etiqueta "nueva"

MESES = {
    "ene": 1, "feb": 2, "mar": 3, "abr": 4, "may": 5, "jun": 6,
    "jul": 7, "ago": 8, "sep": 9, "oct": 10, "nov": 11, "dic": 12,
}

RE_LICITACION = re.compile(r"^(EA-\d+-N\d+-\d{4})\s*\|\s*(.+)$", re.I)
RE_CONVOCATORIA = re.compile(r"^CONVOCATORIA\s*-\s*(\d+)\s*\|\s*(.+)$", re.I)
RE_EVENTO = re.compile(
    r"^(Junta de Aclaraciones|Apertura de Proposiciones|Fallo|Diferimiento de Fallo)"
    r"\b.*?\|\s*(.+)$",
    re.I,
)


def parse_fecha(txt):
    txt = txt.strip()
    m = re.match(r"([A-Za-z]{3})\w*\s+(\d{1,2})\s+(\d{4})", txt)
    if not m:
        return None
    mes = MESES.get(m.group(1).lower()[:3])
    if not mes:
        return None
    try:
        return date(int(m.group(3)), mes, int(m.group(2)))
    except ValueError:
        return None


def descargar_pagina():
    headers = {"User-Agent": "Mozilla/5.0 (monitor-licitaciones; uso personal)"}
    try:
        r = requests.get(URL, headers=headers, timeout=60)
        r.raise_for_status()
        return r.text
    except requests.exceptions.SSLError:
        import urllib3
        urllib3.disable_warnings()
        r = requests.get(URL, headers=headers, timeout=60, verify=False)
        r.raise_for_status()
        return r.text


def extraer_nodos(html):
    soup = BeautifulSoup(html, "html.parser")
    nodos = []
    for el in soup.find_all(["a", "span", "td"]):
        txt = " ".join(el.get_text(" ", strip=True).split())
        if not txt or "|" not in txt:
            continue
        href = el.get("href") if el.name == "a" else None
        if nodos and nodos[-1][0] == txt:
            if href and not nodos[-1][1]:
                nodos[-1] = (txt, href)
            continue
        nodos.append((txt, href))
    return nodos


def parsear_licitaciones(nodos):
    licitaciones = {}
    convocatoria_actual = None
    lic_actual = None
    for txt, href in nodos:
        m = RE_CONVOCATORIA.match(txt)
        if m:
            convocatoria_actual = m.group(1)
            lic_actual = None
            continue
        m = RE_LICITACION.match(txt)
        if m:
            numero = m.group(1).upper()
            fecha_pub = parse_fecha(m.group(2))
            lic = licitaciones.setdefault(numero, {
                "numero": numero,
                "convocatoria": convocatoria_actual,
                "fecha_publicacion": fecha_pub.isoformat() if fecha_pub else m.group(2).strip(),
                "junta_aclaraciones": None,
                "apertura": None,
                "fallo": None,
                "url_pdf": None,
                "objeto": None,
            })
            if href and href.lower().endswith(".pdf") and not lic["url_pdf"]:
                lic["url_pdf"] = href if href.startswith("http") else BASE + "/" + href.lstrip("/")
            lic_actual = numero
            continue
        m = RE_EVENTO.match(txt)
        if m and lic_actual:
            tipo = m.group(1).lower()
            fecha = parse_fecha(m.group(2))
            valor = fecha.isoformat() if fecha else m.group(2).strip()
            lic = licitaciones[lic_actual]
            if "junta" in tipo:
                lic["junta_aclaraciones"] = valor
            elif "apertura" in tipo:
                lic["apertura"] = valor
            elif "diferimiento" in tipo:
                pass
            elif "fallo" in tipo:
                lic["fallo"] = valor
    return licitaciones


def extraer_objeto_pdf(url_pdf):
    if not url_pdf:
        return None
    try:
        import pdfplumber
        from io import BytesIO
        r = requests.get(url_pdf, timeout=60, verify=False)
        r.raise_for_status()
        with pdfplumber.open(BytesIO(r.content)) as pdf:
            texto = ""
            for page in pdf.pages[:2]:
                texto += (page.extract_text() or "") + "\n"
        m = re.search(r"(?:relativa? a l?a?\s+|objeto[:\s]+)(.{20,300}?)(?:\.|\n)", texto, re.I | re.S)
        if m:
            return " ".join(m.group(1).split())
        for linea in texto.splitlines():
            if re.search(r"adquisici[oó]n|servicio|arrendamiento|contrataci[oó]n", linea, re.I) and len(linea) > 30:
                return linea.strip()[:300]
    except Exception as e:
        print(f"  [aviso] No se pudo extraer objeto de {url_pdf}: {e}")
    return None


def cargar_estado():
    if ESTADO_PATH.exists():
        return json.loads(ESTADO_PATH.read_text(encoding="utf-8"))
    return {"vistas": {}}


def guardar_estado(estado):
    DATA_DIR.mkdir(exist_ok=True)
    ESTADO_PATH.write_text(json.dumps(estado, ensure_ascii=False, indent=2), encoding="utf-8")


def guardar_excel(licitaciones):
    import pandas as pd
    filas = sorted(licitaciones.values(), key=lambda x: x["numero"])
    df = pd.DataFrame(filas)[[
        "numero", "convocatoria", "fecha_publicacion",
        "junta_aclaraciones", "apertura", "fallo", "objeto", "url_pdf",
    ]]
    df.columns = [
        "No. Licitación", "Convocatoria", "Publicada",
        "Junta de Aclaraciones", "Apertura", "Fallo", "Objeto", "PDF",
    ]
    DATA_DIR.mkdir(exist_ok=True)
    df.to_excel(EXCEL_PATH, index=False)
    print(f"Excel actualizado: {EXCEL_PATH} ({len(df)} licitaciones)")


def generar_html(licitaciones):
    plantilla = (RAIZ / "plantilla.html").read_text(encoding="utf-8")
    datos = sorted(licitaciones.values(), key=lambda x: x["numero"], reverse=True)
    payload = {
        "actualizado": datetime.now().strftime("%d/%m/%Y %H:%M"),
        "hoy": date.today().isoformat(),
        "dias_nueva": DIAS_COMO_NUEVA,
        "fuente": URL,
        "licitaciones": datos,
    }
    html = plantilla.replace("__DATOS__", json.dumps(payload, ensure_ascii=False))
    DOCS_DIR.mkdir(exist_ok=True)
    HTML_PATH.write_text(html, encoding="utf-8")
    print(f"Tablero web generado: {HTML_PATH}")


def enviar_whatsapp(mensaje):
    phone = os.environ.get("WHATSAPP_PHONE")
    apikey = os.environ.get("CALLMEBOT_APIKEY")
    if not phone or not apikey:
        return
    for bloque in [mensaje[i:i + 1500] for i in range(0, len(mensaje), 1500)]:
        r = requests.get(
            "https://api.callmebot.com/whatsapp.php",
            params={"phone": phone, "text": bloque, "apikey": apikey},
            timeout=60,
        )
        print(f"WhatsApp -> {r.status_code}")


def main():
    print(f"=== Monitor de licitaciones — {datetime.now():%Y-%m-%d %H:%M} ===")
    html = descargar_pagina()
    licitaciones = parsear_licitaciones(extraer_nodos(html))
    if not licitaciones:
        print("ERROR: no se encontraron licitaciones. ¿Cambió la estructura de la página?")
        sys.exit(1)
    print(f"Licitaciones en la página: {len(licitaciones)}")

    estado = cargar_estado()
    vistas = estado["vistas"]
    hoy = date.today().isoformat()
    primera_corrida = len(vistas) == 0

    nuevas = []
    for numero, lic in licitaciones.items():
        previa = vistas.get(numero)
        if previa is None:
            lic["fecha_detectada"] = hoy if not primera_corrida else lic.get("fecha_publicacion", hoy)
            lic["objeto"] = extraer_objeto_pdf(lic["url_pdf"])
            nuevas.append(lic)
        else:
            lic["fecha_detectada"] = previa.get("fecha_detectada", hoy)
            lic["objeto"] = previa.get("objeto") or lic.get("objeto")
        vistas[numero] = lic

    guardar_estado(estado)
    guardar_excel(licitaciones)
    generar_html(vistas)  # el tablero muestra todo lo acumulado

    if nuevas and not primera_corrida:
        msj = f"📋 {len(nuevas)} licitación(es) nueva(s) hoy. Revisa tu tablero."
        enviar_whatsapp(msj)
    print(f"Nuevas detectadas hoy: {len(nuevas)}")


if __name__ == "__main__":
    main()
