"""Lectura de los boletines web de CORPOCESAR (para la pestaña 'Buscar en la CAR').

Índice del año:  https://www.corpocesar.gov.co/boletin-<año>.html
  → enlaces a los boletines mensuales y a los de las seccionales.
Cada boletín: tablas con 'número/dd/mm/aaaa', descripción (incluye el titular) y enlace al PDF.
"""

from __future__ import annotations

import re
import unicodedata
from datetime import date

import requests
import urllib3
from bs4 import BeautifulSoup

BASE = "https://www.corpocesar.gov.co/"
MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto",
         "septiembre", "octubre", "noviembre", "diciembre"]
FECHA = r"(\d{1,2})\s*(?:de|del)?\s+([a-z]{4,10})\s+(?:de|del)?\s*(\d{4})"
PATRON_NUMERO = re.compile(r"^\s*(\d{1,5})\s*/\s*(\d{1,2})\s*/\s*(\d{1,2})\s*/\s*(\d{4})")
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


def norm(s) -> str:
    s = unicodedata.normalize("NFD", str(s or "").upper())
    return re.sub(r"\s+", " ", "".join(c for c in s if unicodedata.category(c) != "Mn"))


def descargar(url: str, timeout: int = 45) -> requests.Response:
    """El sitio no envía el certificado intermedio: si falla TLS se reintenta sin verificar."""
    try:
        r = requests.get(url, timeout=timeout)
    except requests.exceptions.SSLError:
        r = requests.get(url, timeout=timeout, verify=False)
    r.raise_for_status()
    return r


def _html(url: str) -> BeautifulSoup:
    r = descargar(url)
    if "charset" not in r.headers.get("Content-Type", "").lower():
        r.encoding = "utf-8"
    return BeautifulSoup(r.text, "html.parser")


def paginas_del_anio(anio: int) -> list[tuple[str, str]]:
    soup = _html(f"{BASE}boletin-{anio}.html")
    paginas, vistos = [], set()
    for a in soup.find_all("a", href=True):
        href = a["href"].strip()
        if not href.endswith(".html") or str(anio) not in href or href.endswith(f"boletin-{anio}.html"):
            continue
        if not re.search(r"bolet|curumani|aguachica|chimichagua|jagua", href, re.I):
            continue
        completo = href if href.startswith("http") else BASE + href.lstrip("/")
        if completo not in vistos:
            vistos.add(completo)
            paginas.append((completo, a.get_text(" ", strip=True) or completo.rsplit("/", 1)[-1]))
    return paginas


def leer_pagina(url: str, etiqueta: str) -> list[dict]:
    soup = _html(url)
    actos = []
    for tabla in soup.find_all("table"):
        titulo = tabla.find_previous(["strong", "b"])
        categoria = titulo.get_text(" ", strip=True) if titulo else ""
        for fila in tabla.find_all("tr"):
            celdas = fila.find_all("td")
            if len(celdas) < 2:
                continue
            m = PATRON_NUMERO.match(celdas[0].get_text(" ", strip=True))
            if not m:
                continue
            try:
                fecha_acto = date(int(m.group(4)), int(m.group(3)), int(m.group(2))).isoformat()
            except ValueError:
                continue
            enlace = fila.find("a", href=True)
            href = enlace["href"].strip() if enlace else ""
            if href and not href.startswith("http"):
                href = BASE + href.lstrip("/")
            actos.append({
                "Fecha": fecha_acto, "Número": m.group(1), "Categoría": categoria,
                "Descripción": celdas[1].get_text(" ", strip=True), "PDF": href, "Boletín": etiqueta,
            })
    return actos


def a_fecha(d, m, a):
    m = m.lower()
    mes = next((i for i, x in enumerate(MESES, 1) if m.startswith(x[:4])), None)
    try:
        return date(int(a), mes, int(d)).isoformat() if mes else None
    except ValueError:
        return None


def leer_pdf(url: str) -> dict:
    """Fechas clave y señales de requerimientos dentro de la resolución/auto."""
    import pymupdf
    with pymupdf.open(stream=descargar(url, 90).content, filetype="pdf") as doc:
        texto = " ".join(p.get_text() for p in doc)
    t = norm(texto).lower()
    out = {}
    m = re.search(r"exp\s*[-.:]?\s*([a-z]{2,6}\s*-\s*\d{2,4}\s*-\s*\d{4})", t)
    out["Expediente"] = re.sub(r"\s", "", m.group(1)).upper() if m else ""
    m = re.search(r"radicad\w*\s+(?:interno\s+|internamente\s+)?(?:bajo\s+)?(?:el\s+)?no\.?\s*\d[\d\.\s\-]{0,10}"
                  r"(?:del\s+dia|de\s+fecha|del|de)\s+" + FECHA, t)
    out["Radicado"] = a_fecha(*m.groups()[-3:]) if m else None
    out["Auto de inicio"] = None
    for m in re.finditer(r"auto\s+(?:no\.?|n°|nro\.?)\s*\d+\s+(?:del|de)\s+" + FECHA, t):
        if re.search(r"inici|admit|avoc", t[m.end(): m.end() + 250]):
            out["Auto de inicio"] = a_fecha(*m.groups())
            break
    m = re.search(r"(?:visita\s+(?:tecnica\s+)?(?:realizada|practicada|efectuada)\s+el(?:\s+dia)?\s+|el\s+dia\s+)" + FECHA, t)
    out["Visita"] = a_fecha(*m.groups()) if m else None
    out["¿Le pidieron información adicional?"] = (
        "Sí" if re.search(r"requiri\w*\s+(?:al|a\s+la)|informacion\s+(?:adicional|complementaria)|requerimiento", t) else "No aparece")
    out["Páginas"] = texto.count("\f") + 1 if texto else None
    return out
