"""
Seguimientos CARs — CORPOCESAR, todo en una sola app de Streamlit Cloud.

No necesitas instalar nada en tu computador: esto corre completo en el
servidor de Streamlit Cloud (scraping, OCR y visualización).

La única "instalación" que hace falta es el archivo `packages.txt` en el
mismo repo, con estas tres líneas (le dice a Streamlit Cloud qué paquetes de
sistema instalar en SU servidor, no en el tuyo):
    tesseract-ocr
    tesseract-ocr-spa
    poppler-utils

Ojo con esto: el archivo de base de datos (SQLite) vive en el disco del
servidor de Streamlit Cloud, que se reinicia de vez en cuando (por
inactividad, o cuando actualizas el código). Cada vez que eso pase, los
datos que hayas acumulado se pierden y hay que volver a correr el scraping.
Para un uso exploratorio como este está bien; si más adelante quieres que
los datos queden guardados de forma permanente, ahí sí toca conectar una
base de datos externa (por ejemplo Supabase, que tiene plan gratis) — no es
necesario todavía.
"""

from __future__ import annotations

import re
import sqlite3
from dataclasses import asdict, dataclass
from datetime import date

import pandas as pd
import pytesseract
import requests
import streamlit as st
from bs4 import BeautifulSoup
from pdf2image import convert_from_bytes

st.set_page_config(page_title="Tiempos CORPOCESAR", layout="wide")

RUTA_DB = "actos_corpocesar.db"

# ---------------------------------------------------------------------------
# 1. Scraper del boletín mensual
# ---------------------------------------------------------------------------


@dataclass
class Acto:
    categoria: str
    tipo_acto: str
    numero: str
    fecha_acto: str
    descripcion: str
    fecha_publicacion: str
    url_pdf: str
    empresa: str


PALABRAS_CLAVE_PROPIO = ("UNERGY",)


def _detectar_empresa(descripcion: str) -> str:
    descripcion_norm = descripcion.upper()
    if any(palabra in descripcion_norm for palabra in PALABRAS_CLAVE_PROPIO):
        return "propio_unergy"
    return "otro"


def _texto_categoria_a_tipo(categoria: str) -> str:
    categoria_norm = categoria.upper()
    if "RESOLUC" in categoria_norm:
        return "resolucion"
    if "AUTO" in categoria_norm:
        return "auto"
    return "otro"


def parsear_boletin(html: str) -> list[Acto]:
    soup = BeautifulSoup(html, "html.parser")
    actos: list[Acto] = []
    categoria_actual = ""

    for bloque in soup.select('div[id^="wb_LayoutGrid"]'):
        titulo = bloque.find("strong")
        tabla = bloque.find("table")

        if titulo and not tabla:
            categoria_actual = titulo.get_text(strip=True)
            continue

        if tabla and categoria_actual:
            for fila in tabla.find_all("tr")[1:]:
                celdas = fila.find_all("td")
                if len(celdas) < 3:
                    continue
                enlace = celdas[0].find("a")
                if enlace is None:
                    continue
                texto_enlace = enlace.get_text(strip=True)
                if "/" not in texto_enlace:
                    continue
                numero, _, fecha_acto = texto_enlace.partition("/")
                descripcion = celdas[1].get_text(strip=True)
                actos.append(
                    Acto(
                        categoria=categoria_actual,
                        tipo_acto=_texto_categoria_a_tipo(categoria_actual),
                        numero=numero.strip(),
                        fecha_acto=fecha_acto.strip(),
                        descripcion=descripcion,
                        fecha_publicacion=celdas[2].get_text(strip=True),
                        url_pdf=requests.compat.urljoin("https://www.corpocesar.gov.co/", enlace.get("href", "")),
                        empresa=_detectar_empresa(descripcion),
                    )
                )
    return actos


# ---------------------------------------------------------------------------
# 2. OCR de la resolución: radicación + auto de inicio
# ---------------------------------------------------------------------------

MESES = {
    "enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5, "junio": 6,
    "julio": 7, "agosto": 8, "septiembre": 9, "octubre": 10, "noviembre": 11,
    "diciembre": 12,
}
PATRON_FECHA_TEXTO = r"(\d{1,2})\s+de\s+(\w+)\s+de[l]?\s+(\d{4})"
PATRON_RADICACION = re.compile(r"[Rr]adicad\w*[^.]{0,40}?No\.?\s*(\d+)[^.]{0,20}?" + PATRON_FECHA_TEXTO)
PATRON_AUTO = re.compile(r"Auto\s*No\.?\s*(\d+)[^.]{0,20}?" + PATRON_FECHA_TEXTO, re.IGNORECASE)


@dataclass
class Antecedentes:
    radicado_numero: str | None = None
    fecha_radicacion: date | None = None
    auto_inicio_numero: str | None = None
    fecha_auto_inicio: date | None = None


def _texto_a_fecha(dia: str, mes_texto: str, anio: str) -> date | None:
    mes = MESES.get(mes_texto.lower())
    if mes is None:
        return None
    try:
        return date(int(anio), mes, int(dia))
    except ValueError:
        return None


def ocr_primeras_paginas(pdf_bytes: bytes, num_paginas: int = 3) -> str:
    paginas = convert_from_bytes(pdf_bytes, dpi=200, first_page=1, last_page=num_paginas)
    return "\n".join(pytesseract.image_to_string(p, lang="spa") for p in paginas)


def extraer_antecedentes(texto: str) -> Antecedentes:
    resultado = Antecedentes()

    coincidencia = PATRON_RADICACION.search(texto)
    if coincidencia:
        numero, dia, mes_texto, anio = coincidencia.groups()
        resultado.radicado_numero = numero
        resultado.fecha_radicacion = _texto_a_fecha(dia, mes_texto, anio)

    for coincidencia in PATRON_AUTO.finditer(texto):
        numero, dia, mes_texto, anio = coincidencia.groups()
        ventana = texto[coincidencia.end(): coincidencia.end() + 200].lower()
        if "inici" in ventana:
            resultado.auto_inicio_numero = numero
            resultado.fecha_auto_inicio = _texto_a_fecha(dia, mes_texto, anio)
            break

    return resultado


def procesar_resolucion(url_pdf: str) -> Antecedentes:
    pdf_bytes = requests.get(url_pdf, timeout=60).content
    texto = ocr_primeras_paginas(pdf_bytes)
    return extraer_antecedentes(texto)


# ---------------------------------------------------------------------------
# 3. Base de datos (SQLite)
# ---------------------------------------------------------------------------

ESQUEMA = """
CREATE TABLE IF NOT EXISTS actos (
    url_pdf TEXT PRIMARY KEY, categoria TEXT, tipo_acto TEXT, numero TEXT,
    fecha_acto TEXT, descripcion TEXT, fecha_publicacion TEXT, empresa TEXT,
    radicado_numero TEXT, fecha_radicacion TEXT, auto_inicio_numero TEXT,
    fecha_auto_inicio TEXT, dias_radicacion_a_inicio INTEGER,
    dias_radicacion_a_resolucion INTEGER
);
"""


def conectar() -> sqlite3.Connection:
    conn = sqlite3.connect(RUTA_DB)
    conn.execute(ESQUEMA)
    conn.commit()
    return conn


def _fecha_dmy_a_iso(fecha_dmy: str) -> str | None:
    if not fecha_dmy or fecha_dmy.count("/") != 2:
        return None
    dia, mes, anio = fecha_dmy.split("/")
    try:
        return date(int(anio), int(mes), int(dia)).isoformat()
    except ValueError:
        return None


def guardar_acto(conn: sqlite3.Connection, acto: dict) -> None:
    fecha_acto_iso = _fecha_dmy_a_iso(acto.get("fecha_acto", ""))
    dias_a_inicio = dias_a_resolucion = None
    fecha_radicacion = acto.get("fecha_radicacion")
    if fecha_radicacion:
        fr = date.fromisoformat(fecha_radicacion)
        if acto.get("fecha_auto_inicio"):
            dias_a_inicio = (date.fromisoformat(acto["fecha_auto_inicio"]) - fr).days
        if acto.get("tipo_acto") == "resolucion" and fecha_acto_iso:
            dias_a_resolucion = (date.fromisoformat(fecha_acto_iso) - fr).days

    conn.execute(
        """
        INSERT INTO actos (url_pdf, categoria, tipo_acto, numero, fecha_acto, descripcion,
            fecha_publicacion, empresa, radicado_numero, fecha_radicacion, auto_inicio_numero,
            fecha_auto_inicio, dias_radicacion_a_inicio, dias_radicacion_a_resolucion)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(url_pdf) DO UPDATE SET
            categoria=excluded.categoria, tipo_acto=excluded.tipo_acto, numero=excluded.numero,
            fecha_acto=excluded.fecha_acto, descripcion=excluded.descripcion,
            fecha_publicacion=excluded.fecha_publicacion, empresa=excluded.empresa,
            radicado_numero=excluded.radicado_numero, fecha_radicacion=excluded.fecha_radicacion,
            auto_inicio_numero=excluded.auto_inicio_numero, fecha_auto_inicio=excluded.fecha_auto_inicio,
            dias_radicacion_a_inicio=excluded.dias_radicacion_a_inicio,
            dias_radicacion_a_resolucion=excluded.dias_radicacion_a_resolucion
        """,
        (
            acto.get("url_pdf"), acto.get("categoria"), acto.get("tipo_acto"), acto.get("numero"),
            fecha_acto_iso, acto.get("descripcion"), acto.get("fecha_publicacion"), acto.get("empresa"),
            acto.get("radicado_numero"), fecha_radicacion, acto.get("auto_inicio_numero"),
            acto.get("fecha_auto_inicio"), dias_a_inicio, dias_a_resolucion,
        ),
    )
    conn.commit()


def cargar_datos() -> pd.DataFrame:
    conn = conectar()
    df = pd.read_sql("SELECT * FROM actos", conn)
    conn.close()
    for columna in ("fecha_acto", "fecha_radicacion", "fecha_auto_inicio"):
        if columna in df.columns:
            df[columna] = pd.to_datetime(df[columna], errors="coerce")
    return df


# ---------------------------------------------------------------------------
# 4. Interfaz
# ---------------------------------------------------------------------------


def procesar_boletin(url_boletin: str) -> None:
    conn = conectar()
    respuesta = requests.get(url_boletin, timeout=30)
    respuesta.raise_for_status()
    actos = parsear_boletin(respuesta.text)

    resoluciones = [a for a in actos if a.tipo_acto == "resolucion"]
    barra = st.progress(0, text="Procesando resoluciones...")

    for i, acto in enumerate(actos):
        registro = asdict(acto)
        if acto.tipo_acto == "resolucion":
            try:
                antecedentes = procesar_resolucion(acto.url_pdf)
                registro["radicado_numero"] = antecedentes.radicado_numero
                registro["fecha_radicacion"] = (
                    antecedentes.fecha_radicacion.isoformat() if antecedentes.fecha_radicacion else None
                )
                registro["auto_inicio_numero"] = antecedentes.auto_inicio_numero
                registro["fecha_auto_inicio"] = (
                    antecedentes.fecha_auto_inicio.isoformat() if antecedentes.fecha_auto_inicio else None
                )
            except Exception as error:
                st.warning(f"No pude procesar la resolución {acto.numero}: {error}")
            barra.progress(
                min((resoluciones.index(acto) + 1) / max(len(resoluciones), 1), 1.0),
                text=f"Resolución {acto.numero} procesada",
            )
        guardar_acto(conn, registro)

    barra.empty()
    conn.close()
    st.success(f"Listo: {len(actos)} actos guardados ({len(resoluciones)} resoluciones con OCR).")
    cargar_datos.clear()


def main() -> None:
    st.title("Tiempos de trámite — CORPOCESAR")

    with st.sidebar:
        st.header("Actualizar datos")
        url_boletin = st.text_input(
            "URL del boletín mensual",
            value="https://www.corpocesar.gov.co/boletin-septiembre-2026.html",
        )
        if st.button("Procesar boletín", type="primary"):
            with st.spinner("Descargando y procesando (el OCR tarda un poco por resolución)..."):
                procesar_boletin(url_boletin)

    df = cargar_datos()
    if df.empty:
        st.info("Todavía no hay datos. Pon la URL de un boletín en la barra lateral y dale 'Procesar boletín'.")
        return

    st.sidebar.header("Filtros")
    tipos = sorted(df["tipo_acto"].dropna().unique())
    tipo_sel = st.sidebar.multiselect("Tipo de acto", tipos, default=tipos)
    empresas = sorted(df["empresa"].dropna().unique())
    empresa_sel = st.sidebar.multiselect("Empresa", empresas, default=empresas)
    df_filtrado = df[df["tipo_acto"].isin(tipo_sel) & df["empresa"].isin(empresa_sel)]

    tab_actos, tab_tiempos = st.tabs(["Actos", "Tiempos por empresa"])

    with tab_actos:
        c1, c2, c3 = st.columns(3)
        c1.metric("Actos (filtro actual)", len(df_filtrado))
        c2.metric("Resoluciones", int((df_filtrado["tipo_acto"] == "resolucion").sum()))
        c3.metric("De Unergy", int((df_filtrado["empresa"] == "propio_unergy").sum()))
        columnas = [
            "fecha_acto", "tipo_acto", "numero", "empresa", "descripcion",
            "fecha_radicacion", "fecha_auto_inicio", "dias_radicacion_a_inicio", "dias_radicacion_a_resolucion",
        ]
        st.dataframe(df_filtrado[columnas].sort_values("fecha_acto", ascending=False), use_container_width=True, hide_index=True)

    with tab_tiempos:
        resoluciones_con_dato = df_filtrado[df_filtrado["dias_radicacion_a_resolucion"].notna()]
        if resoluciones_con_dato.empty:
            st.info("Todavía no hay resoluciones con radicación detectada en el filtro actual.")
        else:
            st.subheader("Promedio de días, radicación → resolución, por empresa")
            st.bar_chart(resoluciones_con_dato.groupby("empresa")["dias_radicacion_a_resolucion"].mean())

            con_inicio = df_filtrado[df_filtrado["dias_radicacion_a_inicio"].notna()]
            if not con_inicio.empty:
                st.subheader("Promedio de días, radicación → auto de inicio, por empresa")
                st.bar_chart(con_inicio.groupby("empresa")["dias_radicacion_a_inicio"].mean())

            st.subheader("Detalle por resolución")
            columnas_t = [
                "numero", "empresa", "fecha_radicacion", "fecha_auto_inicio",
                "fecha_acto", "dias_radicacion_a_inicio", "dias_radicacion_a_resolucion",
            ]
            st.dataframe(
                resoluciones_con_dato[columnas_t].sort_values("dias_radicacion_a_resolucion", ascending=False),
                use_container_width=True, hide_index=True,
            )


if __name__ == "__main__":
    main()
