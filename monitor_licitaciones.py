# -*- coding: utf-8 -*-
"""
Monitor de licitaciones - e-Oficialía Hidalgo  (v3)
===================================================
Novedades v3:
- Descarga el PDF de cada convocatoria simulando el clic del portal (postback
  ASP.NET) o vía enlace directo si existe.
- Extrae del PDF: objeto de la licitación y fechas PROGRAMADAS (con hora) de
  junta de aclaraciones, apertura y fallo.
- Guarda una copia del PDF en docs/pdfs/ para que el tablero enlace directo.
"""

import json
import os
import re
import sys
import unicodedata
from datetime import datetime, date, timedelta
from io import BytesIO
from pathlib import Path
from urllib.parse import urljoin

import requests
from bs4 import BeautifulSoup

URL = "https://eoficialia.hidalgo.gob.mx/LICITACIONES/VISTAS/WebFrmLC004.aspx"
BASE = "https://eoficialia.hidalgo.gob.mx"

RAIZ = Path(__file__).parent
DATA_DIR = RAIZ / "data"
DOCS_DIR = RAIZ / "docs"
PDF_DIR = DOCS_DIR / "pdfs"
ESTADO_PATH = DATA_DIR / "estado.json"
EXCEL_PATH = DATA_DIR / "licitaciones.xlsx"
HTML_PATH = DOCS_DIR / "index.html"

DIAS_COMO_NUEVA = 3
DIAS_BACKFILL = 20        # al estrenar v3, procesa PDFs de licitaciones publicadas en los últimos N días
MAX_PDFS_POR_CORRIDA = 30 # tope para no saturar el portal
MAX_INTENTOS_PDF = 3

MESES3 = {"ene":1,"feb":2,"mar":3,"abr":4,"may":5,"jun":6,"jul":7,"ago":8,"sep":9,"oct":10,"nov":11,"dic":12}
MESES_LARGOS = {"enero":1,"febrero":2,"marzo":3,"abril":4,"mayo":5,"junio":6,"julio":7,
                "agosto":8,"septiembre":9,"setiembre":9,"octubre":10,"noviembre":11,"diciembre":12}

RE_LICITACION = re.compile(r"^(EA-\d+-N\d+-\d{4})\s*\|\s*(.+)$", re.I)
RE_CONVOCATORIA = re.compile(r"^CONVOCATORIA\s*-\s*(\d+)\s*\|\s*(.+)$", re.I)
RE_EVENTO = re.compile(r"^(Junta de Aclaraciones|Apertura de Proposiciones|Fallo|Diferimiento de Fallo)\b.*?\|\s*(.+)$", re.I)
RE_POSTBACK = re.compile(r"__doPostBack\('([^']*)'\s*,\s*'([^']*)'\)")

S = requests.Session()
S.headers.update({"User-Agent": "Mozilla/5.0 (monitor-licitaciones; uso personal)"})
VERIFY = True


def parse_fecha_corta(txt):
    m = re.match(r"([A-Za-z]{3})\w*\s+(\d{1,2})\s+(\d{4})", txt.strip())
    if not m:
        return None
    mes = MESES3.get(m.group(1).lower()[:3])
    try:
        return date(int(m.group(3)), mes, int(m.group(2))) if mes else None
    except (ValueError, TypeError):
        return None


def get(url, **kw):
    global VERIFY
    try:
        r = S.get(url, timeout=60, verify=VERIFY, **kw)
    except requests.exceptions.SSLError:
        import urllib3; urllib3.disable_warnings()
        VERIFY = False
        r = S.get(url, timeout=60, verify=False, **kw)
    r.raise_for_status()
    return r


# ------------------------- parsing del árbol -------------------------

def extraer_nodos(html):
    """Devuelve lista de (texto, href) en orden de documento.

    El portal arma cada fila como <tr> con el texto ("NUMERO | fecha") en un
    <td> y el botón de descarga como <input onclick="__doPostBack(...)">
    en otro <td> de la misma fila (no hay <a href> reales en la página, solo
    se conserva ese caso como respaldo por si el portal cambia).
    """
    soup = BeautifulSoup(html, "html.parser")
    nodos = []
    for tr in soup.find_all("tr"):
        txt = None
        for el in tr.find_all(["td", "span"]):
            t = " ".join(el.get_text(" ", strip=True).split())
            if t and "|" in t:
                txt = t
                break
        if not txt:
            continue
        href = None
        a = tr.find("a", href=True)
        if a:
            href = a["href"]
        else:
            inp = tr.find("input", onclick=re.compile(r"__doPostBack"))
            if inp:
                href = inp.get("onclick")
        if nodos and nodos[-1][0] == txt:
            if href and not nodos[-1][1]:
                nodos[-1] = (txt, href)
            continue
        nodos.append((txt, href))
    return nodos


def parsear_licitaciones(nodos):
    licitaciones, conv, actual = {}, None, None
    for txt, href in nodos:
        m = RE_CONVOCATORIA.match(txt)
        if m:
            conv, actual = m.group(1), None
            continue
        m = RE_LICITACION.match(txt)
        if m:
            numero = m.group(1).upper()
            f = parse_fecha_corta(m.group(2))
            licitaciones.setdefault(numero, {
                "numero": numero, "convocatoria": conv,
                "fecha_publicacion": f.isoformat() if f else m.group(2).strip(),
                "junta_aclaraciones": None, "apertura": None, "fallo": None,
                "junta_prog": None, "apertura_prog": None, "fallo_prog": None,
                "url_pdf": None, "objeto": None, "href_nodo": href,
            })
            if href and not licitaciones[numero]["href_nodo"]:
                licitaciones[numero]["href_nodo"] = href
            actual = numero
            continue
        m = RE_EVENTO.match(txt)
        if m and actual:
            tipo = m.group(1).lower()
            f = parse_fecha_corta(m.group(2))
            valor = f.isoformat() if f else m.group(2).strip()
            lic = licitaciones[actual]
            if "junta" in tipo:
                lic["junta_aclaraciones"] = valor
            elif "apertura" in tipo:
                lic["apertura"] = valor
            elif "diferimiento" not in tipo:
                lic["fallo"] = valor
    return licitaciones


# ------------------------- descarga del PDF -------------------------

def campos_ocultos(html):
    soup = BeautifulSoup(html, "html.parser")
    data = {}
    for inp in soup.find_all("input", {"type": "hidden"}):
        if inp.get("name"):
            data[inp["name"]] = inp.get("value", "")
    return data


def descargar_pdf_licitacion(html_pagina, href_nodo, numero):
    """
    Intenta obtener el PDF de la convocatoria de una licitación.
    Cubre dos escenarios: enlace directo, o postback de ASP.NET.
    Devuelve bytes del PDF o None.
    """
    if not href_nodo:
        print(f"  [diag] {numero}: el nodo no tiene href")
        return None

    # Caso 1: enlace directo
    if "__doPostBack" not in href_nodo:
        url = href_nodo if href_nodo.startswith("http") else urljoin(URL, href_nodo)
        try:
            r = get(url)
            if r.content[:4] == b"%PDF":
                return r.content
            print(f"  [diag] {numero}: enlace directo no devolvió PDF ({r.headers.get('content-type')})")
        except Exception as e:
            print(f"  [diag] {numero}: error enlace directo: {e}")
        return None

    # Caso 2: postback
    m = RE_POSTBACK.search(href_nodo.replace("\\'", "'"))
    if not m:
        print(f"  [diag] {numero}: no pude interpretar el postback: {href_nodo[:120]}")
        return None
    target, arg = m.group(1), m.group(2)
    data = campos_ocultos(html_pagina)
    data["__EVENTTARGET"] = target
    data["__EVENTARGUMENT"] = arg
    try:
        r = S.post(URL, data=data, timeout=90, verify=VERIFY, allow_redirects=True)
        r.raise_for_status()
        if r.content[:4] == b"%PDF":
            return r.content
        # A veces el postback devuelve una página que abre/enlaza el PDF
        m2 = re.search(r"['\"]([^'\"]+\.pdf[^'\"]*)['\"]", r.text, re.I)
        if m2:
            url = m2.group(1)
            url = url if url.startswith("http") else urljoin(URL, url)
            r2 = get(url)
            if r2.content[:4] == b"%PDF":
                return r2.content
        print(f"  [diag] {numero}: postback no devolvió PDF (ct={r.headers.get('content-type')}, len={len(r.content)})")
    except Exception as e:
        print(f"  [diag] {numero}: error en postback: {e}")
    return None


# ------------------------- lectura del PDF -------------------------

def _normaliza(t):
    t = unicodedata.normalize("NFD", t)
    return "".join(c for c in t if unicodedata.category(c) != "Mn").lower()


RE_FECHA_LARGA = re.compile(
    r"(\d{1,2})\s*(?:de\s+|/|-)?\s*"
    r"(enero|febrero|marzo|abril|mayo|junio|julio|agosto|septiembre|setiembre|octubre|noviembre|diciembre)"
    r"\s*(?:de\s+|del\s+|/|-)?\s*(\d{4})", re.I)
RE_HORA = re.compile(r"(\d{1,2}):(\d{2})\s*(?:horas|hrs|h\b|am|pm)?", re.I)


def _fecha_cerca(texto_norm, texto_orig, claves, ventana=260):
    """Busca la primera fecha (y hora) que aparece poco después de alguna clave."""
    for clave in claves:
        for m in re.finditer(clave, texto_norm):
            tramo = texto_orig[m.end(): m.end() + ventana]
            f = RE_FECHA_LARGA.search(tramo)
            if f:
                mes = MESES_LARGOS.get(_normaliza(f.group(2)))
                try:
                    fecha = date(int(f.group(3)), mes, int(f.group(1)))
                except (ValueError, TypeError):
                    continue
                h = RE_HORA.search(tramo[f.end(): f.end() + 60])
                hora = f"{int(h.group(1)):02d}:{h.group(2)}" if h else None
                return fecha.isoformat(), hora
    return None, None


def analizar_pdf(contenido):
    """Extrae objeto y fechas programadas del PDF de la convocatoria."""
    import pdfplumber
    texto = ""
    try:
        with pdfplumber.open(BytesIO(contenido)) as pdf:
            for page in pdf.pages[:3]:
                texto += (page.extract_text() or "") + "\n"
    except Exception as e:
        print(f"  [diag] pdfplumber falló: {e}")
        return {}
    if not texto.strip():
        print("  [diag] PDF sin texto extraíble (¿escaneado?)")
        return {}
    norm = _normaliza(texto)
    res = {}
    res["junta_prog"], res["junta_hora"] = _fecha_cerca(norm, texto, [r"junta\s+de\s+aclaraciones"])
    res["apertura_prog"], res["apertura_hora"] = _fecha_cerca(
        norm, texto, [r"apertura\s+de\s+prop", r"presentacion\s+y\s+apertura"])
    res["fallo_prog"], res["fallo_hora"] = _fecha_cerca(norm, texto, [r"\bfallo\b"])

    m = re.search(r"(?:relativa?\s+a\s+l?a?\s+|objeto[:\s]+|contratacion\s+de\s+|adquisicion\s+de\s+|servicio\s+de\s+|arrendamiento\s+de\s+)", norm)
    if m:
        inicio = m.start()
        frag = texto[inicio: inicio + 320]
        frag = re.split(r"[\n]|(?<=\.)\s", frag, 1)[0]
        res["objeto"] = " ".join(frag.split())[:300]
    return res


# ------------------------- persistencia y salida -------------------------

def cargar_estado():
    if ESTADO_PATH.exists():
        return json.loads(ESTADO_PATH.read_text(encoding="utf-8"))
    return {"vistas": {}}


def guardar_estado(estado):
    DATA_DIR.mkdir(exist_ok=True)
    for lic in estado["vistas"].values():
        lic.pop("href_nodo", None)  # no persistir (cambia entre corridas)
    ESTADO_PATH.write_text(json.dumps(estado, ensure_ascii=False, indent=2), encoding="utf-8")


def guardar_excel(licitaciones):
    import pandas as pd
    filas = sorted(licitaciones.values(), key=lambda x: x["numero"])
    cols = ["numero","convocatoria","fecha_publicacion","junta_prog","junta_hora",
            "apertura_prog","apertura_hora","fallo_prog","fallo_hora",
            "junta_aclaraciones","apertura","fallo","objeto","url_pdf"]
    df = pd.DataFrame(filas)
    for c in cols:
        if c not in df.columns:
            df[c] = None
    df = df[cols]
    df.columns = ["No. Licitación","Convocatoria","Publicada","Junta (prog.)","Hora",
                  "Apertura (prog.)","Hora","Fallo (prog.)","Hora",
                  "Acta Junta","Acta Apertura","Acta Fallo","Objeto","PDF"]
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
    phone, apikey = os.environ.get("WHATSAPP_PHONE"), os.environ.get("CALLMEBOT_APIKEY")
    if not phone or not apikey:
        return
    for bloque in [mensaje[i:i+1500] for i in range(0, len(mensaje), 1500)]:
        requests.get("https://api.callmebot.com/whatsapp.php",
                     params={"phone": phone, "text": bloque, "apikey": apikey}, timeout=60)


# ------------------------- flujo principal -------------------------

def main():
    print(f"=== Monitor de licitaciones v3 — {datetime.now():%Y-%m-%d %H:%M} ===")
    html = get(URL).text
    licitaciones = parsear_licitaciones(extraer_nodos(html))
    if not licitaciones:
        print("ERROR: no se encontraron licitaciones. ¿Cambió la estructura de la página?")
        sys.exit(1)
    print(f"Licitaciones en la página: {len(licitaciones)}")

    estado = cargar_estado()
    vistas = estado["vistas"]
    hoy = date.today()
    primera_corrida = len(vistas) == 0
    limite_backfill = (hoy - timedelta(days=DIAS_BACKFILL)).isoformat()

    nuevas, candidatas_pdf = [], []
    for numero, lic in licitaciones.items():
        previa = vistas.get(numero)
        if previa is None:
            lic["fecha_detectada"] = hoy.isoformat() if not primera_corrida else lic.get("fecha_publicacion", hoy.isoformat())
            lic["intentos_pdf"] = 0
            nuevas.append(lic)
        else:
            for campo in ("fecha_detectada","objeto","url_pdf","junta_prog","junta_hora",
                          "apertura_prog","apertura_hora","fallo_prog","fallo_hora","intentos_pdf"):
                if previa.get(campo) is not None and lic.get(campo) is None:
                    lic[campo] = previa[campo]
            lic.setdefault("intentos_pdf", 0)
        # ¿Le falta procesar su PDF y es reciente?
        pub = lic.get("fecha_publicacion") or ""
        if (not lic.get("url_pdf")
                and lic.get("intentos_pdf", 0) < MAX_INTENTOS_PDF
                and (previa is None or pub >= limite_backfill)):
            candidatas_pdf.append(lic)
        vistas[numero] = lic

    # Priorizar las licitaciones más recientes primero, ya que el tope por
    # corrida (MAX_PDFS_POR_CORRIDA) puede dejar candidatas sin procesar.
    candidatas_pdf.sort(key=lambda l: (l.get("fecha_publicacion") or "", l["numero"]), reverse=True)

    # Descargar y analizar PDFs (con tope por corrida)
    PDF_DIR.mkdir(parents=True, exist_ok=True)
    procesadas = 0
    for lic in candidatas_pdf:
        if procesadas >= MAX_PDFS_POR_CORRIDA:
            print(f"Tope de {MAX_PDFS_POR_CORRIDA} PDFs alcanzado; el resto queda para mañana.")
            break
        lic["intentos_pdf"] = lic.get("intentos_pdf", 0) + 1
        contenido = descargar_pdf_licitacion(html, lic.get("href_nodo"), lic["numero"])
        if not contenido:
            continue
        nombre = lic["numero"].replace("/", "-") + ".pdf"
        (PDF_DIR / nombre).write_bytes(contenido)
        lic["url_pdf"] = f"pdfs/{nombre}"
        datos = analizar_pdf(contenido)
        for k in ("objeto","junta_prog","junta_hora","apertura_prog","apertura_hora","fallo_prog","fallo_hora"):
            if datos.get(k):
                lic[k] = datos[k]
        procesadas += 1
        print(f"  ✓ {lic['numero']}: PDF guardado"
              f" | junta {lic.get('junta_prog') or '—'} | apertura {lic.get('apertura_prog') or '—'}"
              f" | fallo {lic.get('fallo_prog') or '—'} | objeto {'sí' if lic.get('objeto') else 'no'}")

    guardar_estado(estado)
    guardar_excel(vistas)
    generar_html(vistas)

    if nuevas and not primera_corrida:
        enviar_whatsapp(f"📋 {len(nuevas)} licitación(es) nueva(s) hoy. Revisa tu tablero.")
    print(f"Resumen: {len(nuevas)} nuevas | {procesadas} PDFs procesados en esta corrida")


if __name__ == "__main__":
    main()
