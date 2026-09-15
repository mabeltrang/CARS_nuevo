"""
Tiempos de trámite CORPOCESAR — a partir del Excel de seguimiento
(el que genera el pipeline de extracción, o el que llevas a mano).

No hace scraping ni OCR: simplemente subes el .xlsx con estas columnas
(las mismas que ya usas) y la app calcula los tiempos entre etapas.

Columnas esperadas:
  Archivo, CAR, Tipo, Categoría Trámite, Tipo de aprovechamiento, Titular,
  Tipo de persona, Fecha radicado inicio trámite, Fecha auto, Fecha visita,
  Fecha resolución, # Arboles, Vol (m3), Seccional, Requerimientos
"""

from __future__ import annotations

import os
import re

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Tiempos CORPOCESAR", layout="wide")

# El Excel vive en el repo, junto a este archivo — así la app no depende de
# que alguien lo suba manualmente cada vez que la abre. Para actualizar los
# datos: reemplaza este archivo en el repo (arrastrándolo en GitHub o con
# git) y haz commit — Streamlit Cloud recoge el cambio solo.
RUTA_EXCEL_REPO = os.path.join(os.path.dirname(__file__), "Fechas_CARS.xlsx")

COLUMNAS_FECHA = [
    "Fecha radicado inicio trámite", "Fecha auto", "Fecha visita", "Fecha resolución",
]


def normalizar_titular(t):
    """Une nombres partidos por salto de línea y quita comas/espacios sobrantes,
    para que 'UNERGY\\nENERGÍA...' y 'UNERGY ENERGÍA...,' cuenten como el mismo."""
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
    # clave de agrupación insensible a puntos ("S.A.S" vs "SAS", "E.S.P" vs "ESP")
    # y mayúsculas/minúsculas, para no separar la misma entidad en dos filas
    df["_titular_clave"] = (
        df["Titular"].astype(str).str.upper().str.replace(".", "", regex=False).str.strip()
    )

    df["dias_radicado_auto"] = (df["Fecha auto"] - df["Fecha radicado inicio trámite"]).dt.days
    df["dias_auto_visita"] = (df["Fecha visita"] - df["Fecha auto"]).dt.days
    df["dias_visita_resolucion"] = (df["Fecha resolución"] - df["Fecha visita"]).dt.days
    df["dias_radicado_resolucion"] = (df["Fecha resolución"] - df["Fecha radicado inicio trámite"]).dt.days

    intervalos = ["dias_radicado_auto", "dias_auto_visita", "dias_visita_resolucion", "dias_radicado_resolucion"]

    def marcar_anomalia(row):
        for c in intervalos:
            v = row[c]
            if pd.notna(v) and v < 0:
                return f"{c} negativo"
        return None

    df["anomalia"] = df.apply(marcar_anomalia, axis=1)
    df["es_unergy"] = df["_titular_clave"].str.contains("UNERGY", na=False)

    return df


def main() -> None:
    st.title("Tiempos de trámite — CORPOCESAR")

    archivo = st.sidebar.file_uploader(
        "Sube un Excel para reemplazar el del repo (opcional)", type=["xlsx"]
    )

    if archivo:
        contenido = archivo.read()
        st.sidebar.caption("Usando el archivo que acabas de subir.")
    elif os.path.exists(RUTA_EXCEL_REPO):
        with open(RUTA_EXCEL_REPO, "rb") as f:
            contenido = f.read()
        fecha_mod = pd.Timestamp(os.path.getmtime(RUTA_EXCEL_REPO), unit="s")
        st.sidebar.caption(f"Usando el Excel del repo (actualizado: {fecha_mod:%Y-%m-%d %H:%M}).")
    else:
        st.info(
            "No encontré 'Fechas_CARS.xlsx' en el repo y no has subido ninguno. "
            "Sube el Excel en la barra lateral, o agrégalo al repo con ese nombre."
        )
        return

    df = cargar_datos(contenido)

    with st.sidebar:
        st.header("Filtros")
        categorias = sorted(df["Categoría Trámite"].dropna().unique())
        cat_sel = st.multiselect("Categoría trámite", categorias, default=categorias)
        tipos_persona = sorted(df["Tipo de persona"].dropna().unique())
        persona_sel = st.multiselect("Tipo de persona", tipos_persona, default=tipos_persona)

    df_f = df[df["Categoría Trámite"].isin(cat_sel) & df["Tipo de persona"].isin(persona_sel)]

    con_anomalia = df_f[df_f["anomalia"].notna()]
    if not con_anomalia.empty:
        with st.expander(f"⚠️ {len(con_anomalia)} fila(s) con fechas sospechosas — excluidas de los promedios"):
            st.dataframe(
                con_anomalia[["Archivo", "Titular", "anomalia"] + COLUMNAS_FECHA],
                use_container_width=True, hide_index=True,
            )
    limpio = df_f[df_f["anomalia"].isna()]

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Trámites (filtro actual)", len(df_f))
    c2.metric("Con resolución", int(limpio["Fecha resolución"].notna().sum()))
    c3.metric("De Unergy", int(df_f["es_unergy"].sum()))
    c4.metric("Con fechas sospechosas", len(con_anomalia))

    tab_resumen, tab_persona, tab_titular, tab_detalle = st.tabs(
        ["Tiempos promedio", "Empresa vs. persona natural", "¿Quién sale más rápido?", "Detalle"]
    )

    with tab_resumen:
        st.subheader("Tiempo promedio entre etapas (días)")
        intervalos = {
            "dias_radicado_auto": "Radicado → Auto",
            "dias_auto_visita": "Auto → Visita",
            "dias_visita_resolucion": "Visita → Resolución",
            "dias_radicado_resolucion": "Radicado → Resolución (total)",
        }
        resumen = pd.DataFrame({
            "Etapa": list(intervalos.values()),
            "Promedio (días)": [limpio[c].mean() for c in intervalos],
            "Mediana (días)": [limpio[c].median() for c in intervalos],
            "Casos con dato": [limpio[c].notna().sum() for c in intervalos],
        }).round(1)
        st.dataframe(resumen, use_container_width=True, hide_index=True)
        st.bar_chart(resumen.set_index("Etapa")["Promedio (días)"])

    with tab_persona:
        st.subheader("Radicado → Resolución, por tipo de persona")
        comparable = limpio[limpio["dias_radicado_resolucion"].notna()]
        if comparable.empty:
            st.info("Todavía no hay trámites resueltos en este filtro.")
        else:
            por_persona = comparable.groupby("Tipo de persona")["dias_radicado_resolucion"].agg(
                ["mean", "median", "count"]
            ).round(1).rename(columns={"mean": "Promedio", "median": "Mediana", "count": "Casos"})
            st.dataframe(por_persona, use_container_width=True)
            st.bar_chart(por_persona["Promedio"])

            st.subheader("Por categoría de trámite")
            por_categoria = comparable.groupby("Categoría Trámite")["dias_radicado_resolucion"].agg(
                ["mean", "median", "count"]
            ).round(1).rename(columns={"mean": "Promedio", "median": "Mediana", "count": "Casos"})
            st.dataframe(por_categoria, use_container_width=True)

    with tab_titular:
        st.subheader("Ranking de titulares — Radicado → Resolución (más rápido primero)")
        comparable = limpio[limpio["dias_radicado_resolucion"].notna()]
        if comparable.empty:
            st.info("Todavía no hay trámites resueltos en este filtro.")
        else:
            minimo_casos = st.slider("Mostrar solo titulares con al menos N trámites resueltos", 1, 5, 1)
            ranking = (
                comparable.groupby("_titular_clave")
                .agg(Titular=("Titular", "first"), Promedio=("dias_radicado_resolucion", "mean"),
                     Casos=("dias_radicado_resolucion", "count"))
                .round(1)
            )
            ranking = ranking[ranking["Casos"] >= minimo_casos].sort_values("Promedio")
            ranking = ranking.set_index("Titular")[["Promedio", "Casos"]]

            col_a, col_b = st.columns(2)
            with col_a:
                st.markdown("**Más rápidos**")
                st.dataframe(ranking.head(10), use_container_width=True)
            with col_b:
                st.markdown("**Más lentos**")
                st.dataframe(ranking.tail(10).sort_values("Promedio", ascending=False), use_container_width=True)

            fila_unergy = ranking[ranking.index.str.contains("UNERGY", case=False, na=False)]
            if not fila_unergy.empty:
                st.subheader("Unergy frente al resto")
                promedio_otros = ranking[~ranking.index.isin(fila_unergy.index)]["Promedio"].mean()
                cu1, cu2 = st.columns(2)
                cu1.metric("Promedio Unergy (días)", f"{fila_unergy['Promedio'].mean():.0f}", help=f"n={int(fila_unergy['Casos'].sum())}")
                cu2.metric("Promedio del resto de titulares (días)", f"{promedio_otros:.0f}")

    with tab_detalle:
        columnas = [
            "Archivo", "Titular", "Tipo de persona", "Categoría Trámite", "Tipo de aprovechamiento",
            "Seccional", "Fecha radicado inicio trámite", "Fecha auto", "Fecha visita", "Fecha resolución",
            "dias_radicado_auto", "dias_auto_visita", "dias_visita_resolucion", "dias_radicado_resolucion",
            "Requerimientos",
        ]
        columnas = [c for c in columnas if c in df_f.columns]
        st.dataframe(df_f[columnas].sort_values("Fecha resolución", ascending=False), use_container_width=True, hide_index=True)


if __name__ == "__main__":
    main()
    
