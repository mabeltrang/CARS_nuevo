"""
Tiempos de trámite CORPOCESAR — enfocado en empresas/entidades.

El Excel vive en el repo (Fechas_CARS.xlsx, junto a este archivo) y se carga
solo; el uploader de la barra lateral es solo para reemplazarlo puntualmente.
"""

from __future__ import annotations

import os
import re
import unicodedata

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Tiempos CORPOCESAR", layout="wide")

RUTA_EXCEL_REPO = os.path.join(os.path.dirname(__file__), "Fechas_CARS.xlsx")
COLUMNAS_FECHA = ["Fecha radicado inicio trámite", "Fecha auto", "Fecha visita", "Fecha resolución"]

COLOR_PRIMARIO = "#0F9D58"  # placeholder tipo "energía limpia" — cámbialo aquí si tienes el verde/azul oficial de Unergy
COLOR_ALERTA = "#dc2626"
LOGO_CORPOCESAR = "https://www.corpocesar.gov.co/images/LogoCorpocesar%20SIN%20FONDO.png"


def quitar_acentos(texto: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFKD", texto) if not unicodedata.combining(c))


def normalizar_titular(t):
    if pd.isna(t):
        return t
    t = re.sub(r"\s+", " ", str(t).replace("\n", " ")).strip()
    return t.rstrip(",").strip()


@st.cache_data(show_spinner=False)
def cargar_datos(contenido: bytes) -> pd.DataFrame:
    df = pd.read_excel(pd.io.common.BytesIO(contenido))
    df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")]

    for c in COLUMNAS_FECHA:
        if c in df.columns:
            df[c] = pd.to_datetime(df[c], errors="coerce")

    df["Titular"] = df["Titular"].apply(normalizar_titular)
    df["_titular_clave"] = (
        df["Titular"].astype(str).apply(quitar_acentos).str.upper()
        .str.replace(r"[.,]", " ", regex=True).str.replace(r"\s+", " ", regex=True).str.strip()
    )

    df["dias_radicado_auto"] = (df["Fecha auto"] - df["Fecha radicado inicio trámite"]).dt.days
    df["dias_auto_visita"] = (df["Fecha visita"] - df["Fecha auto"]).dt.days
    df["dias_visita_resolucion"] = (df["Fecha resolución"] - df["Fecha visita"]).dt.days
    df["dias_radicado_resolucion"] = (df["Fecha resolución"] - df["Fecha radicado inicio trámite"]).dt.days

    intervalos = ["dias_radicado_auto", "dias_auto_visita", "dias_visita_resolucion", "dias_radicado_resolucion"]
    df["_anomalia"] = False
    for c in intervalos:
        df["_anomalia"] |= df[c] < 0

    df["es_unergy"] = df["_titular_clave"].str.contains("UNERGY", na=False)
    df["mes_radicado"] = df["Fecha radicado inicio trámite"].dt.to_period("M").astype(str)
    df["num_arboles_num"] = pd.to_numeric(df["# Arboles"], errors="coerce")

    return df


def obtener_unidad() -> str:
    return st.query_params.get("unidad", "dias")


def _href_alternar_unidad() -> str:
    return "?unidad=meses" if obtener_unidad() == "dias" else "?unidad=dias"


def formatear_dias(valor: float | None, unidad: str) -> str:
    if valor is None or pd.isna(valor):
        return "—"
    if unidad == "meses":
        return f"{valor / 30:.1f}"
    return f"{valor:.0f}"


def circulo(valor: str, etiqueta: str, color: str = COLOR_PRIMARIO, clicable: bool = False) -> str:
    contenido = f"""
      <div style="width:118px;height:118px;border-radius:50%;background:{color}18;
                  border:5px solid {color};display:flex;align-items:center;justify-content:center;">
        <span style="font-size:26px;font-weight:700;color:{color};">{valor}</span>
      </div>
      <div style="font-size:13px;text-align:center;color:#555;max-width:140px;line-height:1.3;">{etiqueta}</div>
    """
    envoltura_abre = f'<a href="{_href_alternar_unidad()}" target="_self" style="text-decoration:none;cursor:pointer;">' if clicable else ""
    envoltura_cierra = "</a>" if clicable else ""
    return f"""
    <div style="display:flex;flex-direction:column;align-items:center;gap:10px;padding:8px;">
      {envoltura_abre}{contenido}{envoltura_cierra}
    </div>
    """


def fila_de_circulos(items: list[tuple[str, str, str]], clicable: bool = False) -> None:
    cols = st.columns(len(items))
    for col, (valor, etiqueta, color) in zip(cols, items):
        with col:
            st.markdown(circulo(valor, etiqueta, color, clicable=clicable), unsafe_allow_html=True)


def main() -> None:
    col_logo, col_titulo = st.columns([1, 6])
    with col_logo:
        st.image(LOGO_CORPOCESAR, width=90)
    with col_titulo:
        st.title("Tiempos de trámite — CORPOCESAR")

    archivo = st.sidebar.file_uploader("Reemplazar el Excel del repo (opcional)", type=["xlsx"])
    if archivo:
        contenido = archivo.read()
        st.sidebar.caption("Usando el archivo que acabas de subir.")
    elif os.path.exists(RUTA_EXCEL_REPO):
        with open(RUTA_EXCEL_REPO, "rb") as f:
            contenido = f.read()
        fecha_mod = pd.Timestamp(os.path.getmtime(RUTA_EXCEL_REPO), unit="s")
        st.sidebar.caption(f"Excel del repo — actualizado {fecha_mod:%Y-%m-%d %H:%M}")
    else:
        st.info("No encontré 'Fechas_CARS.xlsx' en el repo. Súbelo en la barra lateral.")
        return

    df = cargar_datos(contenido)

    with st.sidebar:
        st.header("Filtros")
        solo_empresas = st.checkbox("Solo empresas / entidades", value=True)
        categorias_validas = ["Aprovechamiento forestal", "Ocupación de cauce"]
        cat_sel = [c for c in categorias_validas if st.checkbox(c, value=True)]

    base = df[df["Categoría Trámite"].isin(cat_sel)]
    if solo_empresas:
        base = base[base["Tipo de persona"] == "Empresa / Entidad"]
    limpio = base[~base["_anomalia"]]

    if base["_anomalia"].sum():
        st.caption(f"({int(base['_anomalia'].sum())} trámite(s) con fechas inconsistentes se excluyen de los promedios)")

    unidad = obtener_unidad()
    st.subheader(f"Tiempo promedio entre etapas ({'días' if unidad == 'dias' else 'meses'})")
    st.caption("Haz clic en cualquier círculo para alternar entre días y meses.")
    prom = lambda c: formatear_dias(limpio[c].mean(), unidad) if limpio[c].notna().any() else "—"
    fila_de_circulos([
        (prom("dias_radicado_auto"), "Radicado → Auto", COLOR_PRIMARIO),
        (prom("dias_auto_visita"), "Auto → Visita", COLOR_PRIMARIO),
        (prom("dias_visita_resolucion"), "Visita → Resolución", COLOR_PRIMARIO),
        (prom("dias_radicado_resolucion"), "Radicado → Resolución (total)", COLOR_ALERTA),
    ], clicable=True)

    st.divider()

    tab_modalidad, tab_ranking, tab_tendencia, tab_detalle = st.tabs(
        ["Aislado vs. Único", "¿Quién sale más rápido?", "Tendencia", "Detalle"]
    )

    comparable = limpio[limpio["dias_radicado_resolucion"].notna()]

    with tab_modalidad:
        st.subheader("Radicado → Resolución, según modalidad de aprovechamiento")
        forestal = comparable[comparable["Categoría Trámite"] == "Aprovechamiento forestal"]
        forestal = forestal[forestal["Tipo de aprovechamiento"].notna()]
        if forestal.empty:
            st.info("No hay suficientes trámites forestales resueltos en este filtro.")
        else:
            por_modalidad = forestal.groupby("Tipo de aprovechamiento")["dias_radicado_resolucion"].agg(["mean", "count"])
            items = [
                (f"{fila['mean']:.0f}", f"{modalidad} (n={int(fila['count'])})", COLOR_PRIMARIO)
                for modalidad, fila in por_modalidad.iterrows()
            ]
            fila_de_circulos(items)
            st.caption("Único (el trámite forestal completo) tarda bastante más que Aislado (árboles urbanos aislados) — tiene sentido por la complejidad, pero ayuda a poner expectativas.")

        num_arboles_valido = forestal[["num_arboles_num", "dias_radicado_resolucion"]].dropna()
        if len(num_arboles_valido) >= 5:
            correlacion = num_arboles_valido.corr().iloc[0, 1]
            st.caption(f"A más árboles solicitados, más tarda el trámite (correlación de {correlacion:.2f} sobre {len(num_arboles_valido)} casos).")

    with tab_ranking:
        st.subheader("Radicado → Resolución, por titular (más rápido primero)")
        if comparable.empty:
            st.info("Todavía no hay trámites resueltos en este filtro.")
        else:
            minimo_casos = st.slider("Mostrar solo titulares con al menos N trámites resueltos", 1, 5, 1)
            ranking = (
                comparable.groupby("_titular_clave")
                .agg(Titular=("Titular", "first"), Días=("dias_radicado_resolucion", "mean"),
                     Casos=("dias_radicado_resolucion", "count"))
                .round(0)
            )
            ranking = ranking[ranking["Casos"] >= minimo_casos].sort_values("Días")
            ranking = ranking.set_index("Titular")[["Días", "Casos"]]

            col_a, col_b = st.columns(2)
            with col_a:
                st.markdown("**Más rápidos**")
                st.dataframe(ranking.head(10), use_container_width=True)
            with col_b:
                st.markdown("**Más lentos**")
                st.dataframe(ranking.tail(10).sort_values("Días", ascending=False), use_container_width=True)
            st.caption("Cuando 'Casos' es 1, el número de 'Días' es ese único trámite, no un promedio real — súbelo con el control de arriba para ver solo titulares con más de un caso.")

            fila_unergy = ranking[ranking.index.str.contains("UNERGY", case=False, na=False)]
            if not fila_unergy.empty:
                resto = ranking[~ranking.index.isin(fila_unergy.index)]
                fila_de_circulos([
                    (f"{fila_unergy['Días'].mean():.0f}", f"Unergy (n={int(fila_unergy['Casos'].sum())})", COLOR_ALERTA),
                    (f"{resto['Días'].mean():.0f}", f"Resto de titulares (n={len(resto)})", COLOR_PRIMARIO),
                ])

    with tab_tendencia:
        st.subheader("Radicado → Resolución, por mes de radicación")
        if comparable.empty:
            st.info("Todavía no hay trámites resueltos en este filtro.")
        else:
            tendencia = comparable.groupby("mes_radicado")["dias_radicado_resolucion"].mean().sort_index()
            st.line_chart(tendencia)
            st.caption("Los trámites radicados en 2024 tardaban varios cientos de días; los más recientes se están resolviendo mucho más rápido.")

    with tab_detalle:
        columnas = [
            "Archivo", "Titular", "Categoría Trámite", "Tipo de aprovechamiento", "Seccional",
            "Fecha radicado inicio trámite", "Fecha resolución", "dias_radicado_resolucion", "Requerimientos",
        ]
        columnas = [c for c in columnas if c in base.columns]
        st.dataframe(
            base[columnas].sort_values("Fecha resolución", ascending=False),
            use_container_width=True, hide_index=True,
        )


if __name__ == "__main__":
    main()
