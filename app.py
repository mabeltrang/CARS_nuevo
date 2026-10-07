"""
Tiempos de trámite CORPOCESAR — enfocado en empresas/entidades.

El Excel vive en el repo (Fechas_CARS.xlsx, junto a este archivo) y se carga
solo; el uploader de la barra lateral es solo para reemplazarlo puntualmente.
"""

from __future__ import annotations

import os
import re
import unicodedata

import altair as alt
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


def unir_actos_en_tramites(df: pd.DataFrame) -> pd.DataFrame:
    """El Excel tiene una fila por ACTO (auto, resolución...), así que un mismo trámite
    puede aparecer varias veces. Para no contarlo doble:

    1. Filas con las mismas fechas (radicado, auto, visita, resolución) del mismo titular
       y categoría son el mismo trámite: se dejan en una sola fila.
    2. Un auto sin resolución cuyo radicado y fecha de auto coinciden con un trámite ya
       resuelto es el auto de ESE trámite: se quita (uno por cada resolución).
    No se unen trámites distintos aunque se hayan radicado el mismo día (Unergy radica
    varios proyectos a la vez y el Excel no trae el número de expediente).
    """
    df = df.copy()
    fechas = [c for c in COLUMNAS_FECHA if c in df.columns]
    df["_k_tit"] = df["_titular_clave"].fillna("") + "|" + df["Categoría Trámite"].fillna("")
    clave_completa = ["_k_tit"] + fechas
    archivos = (df.assign(_a=df["Archivo"].fillna("").astype(str))
                .groupby(clave_completa, dropna=False)["_a"]
                .agg(lambda s: " + ".join(x for x in s if x)))
    tipos = (df.assign(_t=df["Tipo"].fillna("").astype(str))
             .groupby(clave_completa, dropna=False)["_t"]
             .agg(lambda s: " + ".join(sorted(set(x for x in s if x)))))
    # resoluciones sin fecha de resolución: casi siempre la fecha quedó en otra columna
    df["_revisar"] = df["Tipo"].astype(str).str.lower().str.startswith("resol") & df["Fecha resolución"].isna()
    unidos = df.drop_duplicates(subset=clave_completa).copy()
    idx = pd.MultiIndex.from_frame(unidos[clave_completa])
    unidos["Archivo"] = archivos.reindex(idx).values
    unidos["Tipo"] = tipos.reindex(idx).values

    resueltos = unidos[unidos["Fecha resolución"].notna()]
    clave_auto = ["_k_tit", "Fecha radicado inicio trámite", "Fecha auto"]
    cupo = resueltos.groupby(clave_auto, dropna=True).size().to_dict()
    quitar = []
    for i, fila in unidos[unidos["Fecha resolución"].isna() & ~unidos["_revisar"]].iterrows():
        k = tuple(fila[c] for c in clave_auto)
        if cupo.get(k, 0) > 0:
            cupo[k] -= 1
            quitar.append(i)
    unidos = unidos.drop(index=quitar).drop(columns=["_k_tit"])
    return unidos.reset_index(drop=True)


@st.cache_data(show_spinner=False)
def cargar_datos(contenido: bytes) -> pd.DataFrame:
    df = pd.read_excel(pd.io.common.BytesIO(contenido))
    df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")]

    for c in COLUMNAS_FECHA:
        if c in df.columns:
            df[c] = pd.to_datetime(df[c], errors="coerce")

    df["Titular"] = df["Titular"].apply(normalizar_titular)
    df["_titular_clave"] = df["Titular"].astype(str).apply(quitar_acentos).str.upper()
    # colapsa siglas tipo "S.A.S", "E.S.P" quitándoles los puntos SIN dejar espacio
    # (si no, "E.S.P" -> "E S P" y "ESP" -> "ESP" quedan como cosas distintas)
    df["_titular_clave"] = df["_titular_clave"].str.replace(
        r"\b(?:[A-Z]\.){1,}[A-Z]?\.?", lambda m: m.group(0).replace(".", ""), regex=True
    )
    df["_titular_clave"] = (
        df["_titular_clave"].str.replace(r"[.,]", " ", regex=True)
        .str.replace(r"\s+", " ", regex=True).str.strip()
    )

    df = unir_actos_en_tramites(df)

    df["dias_radicado_auto"] = (df["Fecha auto"] - df["Fecha radicado inicio trámite"]).dt.days
    df["dias_auto_visita"] = (df["Fecha visita"] - df["Fecha auto"]).dt.days
    df["dias_visita_resolucion"] = (df["Fecha resolución"] - df["Fecha visita"]).dt.days
    df["dias_radicado_resolucion"] = (df["Fecha resolución"] - df["Fecha radicado inicio trámite"]).dt.days

    intervalos = ["dias_radicado_auto", "dias_auto_visita", "dias_visita_resolucion", "dias_radicado_resolucion"]
    df["_anomalia"] = False
    for c in intervalos:
        df["_anomalia"] |= df[c] < 0

    df["es_unergy"] = df["_titular_clave"].str.contains("UNERGY", na=False)
    clave = df["_titular_clave"].fillna("")
    # filas sin "Tipo de persona": si la razón social es de sociedad, se toma como empresa
    sin_tipo = df["Tipo de persona"].isna() | (df["Tipo de persona"].astype(str).str.strip() == "")
    es_sociedad = clave.str.contains(r"\b(?:SAS|SA|ESP|LTDA|SCA|SCS|CONSORCIO|UNION TEMPORAL)\b", regex=True)
    df.loc[sin_tipo & es_sociedad, "Tipo de persona"] = "Empresa / Entidad"
    df["es_esp"] = clave.str.contains(r"\bESP\b", regex=True)
    df["es_energia"] = df["es_esp"] | clave.str.contains(r"SOLAR|SUN\b|ENERG|RENOVABL|ELECTRI|FOTOVOLT|EOLIC", regex=True)
    df["es_publica"] = clave.str.contains(
        r"ALCALD|MUNICIPIO|GOBERNACION|DEPARTAMENTO|SECRETARI|^ESE |HOSPITAL|UNIVERSIDAD|JUNTA DE ACCION", regex=True
    )
    df["mes_radicado"] = df["Fecha radicado inicio trámite"].dt.to_period("M").astype(str)
    df["num_arboles_num"] = pd.to_numeric(df["# Arboles"], errors="coerce")
    for c in ("# Arboles", "Vol (m3)"):
        if c in df.columns:
            df[c] = df[c].astype("string")

    return df


def formatear_dias(valor: float | None, unidad: str) -> str:
    if valor is None or pd.isna(valor):
        return "—"
    if unidad == "Meses":
        return f"{valor / 30:.1f}"
    return f"{valor:.0f}"


def circulo(valor: str, etiqueta: str, color: str = COLOR_PRIMARIO) -> str:
    return f"""
    <div style="display:flex;flex-direction:column;align-items:center;gap:10px;padding:8px;">
      <div style="width:118px;height:118px;border-radius:50%;background:{color}18;
                  border:5px solid {color};display:flex;align-items:center;justify-content:center;">
        <span style="font-size:26px;font-weight:700;color:{color};">{valor}</span>
      </div>
      <div style="font-size:13px;text-align:center;color:#555;max-width:140px;line-height:1.3;">{etiqueta}</div>
    </div>
    """


def fila_de_circulos(items: list[tuple[str, str, str]]) -> None:
    cols = st.columns(len(items))
    for col, (valor, etiqueta, color) in zip(cols, items):
        with col:
            st.markdown(circulo(valor, etiqueta, color), unsafe_allow_html=True)




def seccion_pares(limpio: pd.DataFrame, unidad: str) -> None:
    """Compara a Unergy contra un grupo de pares, con la misma modalidad de trámite."""
    divisor = 30 if unidad == "Meses" else 1
    empresas = limpio[(limpio["Tipo de persona"] == "Empresa / Entidad") & ~limpio["es_publica"]]

    with st.container(border=True):
        st.subheader("¿Cuánto se demora Unergy frente a las demás empresas?")
        forestal_todo = empresas[empresas["Categoría Trámite"] == "Aprovechamiento forestal"]
        cauce = empresas[empresas["Categoría Trámite"] == "Ocupación de cauce"]
        unergy = forestal_todo[forestal_todo["es_unergy"]]
        modalidades_unergy = sorted(unergy["Tipo de aprovechamiento"].dropna().unique())
        misma_modalidad = st.checkbox(
            f"En aprovechamiento forestal, comparar solo la modalidad de Unergy ({', '.join(modalidades_unergy) or '—'})",
            value=bool(modalidades_unergy),
            help="Un aprovechamiento Único tarda mucho más que uno Aislado: compararlos juntos no es justo. "
                 "No afecta a ocupación de cauce.",
        )
        pares = forestal_todo[~forestal_todo["es_unergy"]]
        if misma_modalidad and modalidades_unergy:
            pares = pares[pares["Tipo de aprovechamiento"].isin(modalidades_unergy)]
            unergy = unergy[unergy["Tipo de aprovechamiento"].isin(modalidades_unergy)]

        res_u = unergy[unergy["dias_radicado_resolucion"].notna()]
        res_p = pares[pares["dias_radicado_resolucion"].notna()]
        res_c = cauce[cauce["dias_radicado_resolucion"].notna()]
        if res_u.empty and res_p.empty and res_c.empty:
            st.info("No hay trámites resueltos para comparar.")
            return
        res_e = res_p[res_p["es_energia"]]

        # --- Resumen (aprovechamiento forestal, que es el trámite de Unergy) ---
        def circulo_de(d, nombre, color):
            if not len(d):
                return ("—", nombre, color)
            auto = d["dias_radicado_auto"].mean()
            total = d["dias_radicado_resolucion"].mean()
            desglose = (f"<br><span style='font-size:12px'>{formatear_dias(auto, unidad)} hasta el auto + "
                        f"{formatear_dias(total - auto, unidad)} hasta la resolución</span>") if pd.notna(auto) else ""
            return (formatear_dias(total, unidad), f"{nombre} (n={len(d)}){desglose}", color)
        fila_de_circulos([
            circulo_de(res_u, "Unergy", COLOR_ALERTA),
            circulo_de(res_e, "Empresas de energía", "#475569"),
            circulo_de(res_p, "Todas las empresas privadas", "#94a3b8"),
        ])
        st.caption(f"Aprovechamiento forestal: {unidad.lower()} promedio de radicado a resolución, "
                   "y cuánto de eso fue hasta el auto de inicio. n = trámites resueltos.")

        # --- Por empresa: barra partida en radicado→auto y auto→resolución ---
        st.divider()
        st.markdown(f"**Radicado → Auto → Resolución, por empresa ({unidad.lower()} promedio)**")
        comp = pd.concat([res_u, res_p, res_c])
        comp = comp.assign(
            total=comp["dias_radicado_resolucion"] / divisor,
            hasta_auto=(comp["dias_radicado_auto"] / divisor),
            Trámite=comp["Categoría Trámite"],
        )
        comp["Empresa"] = comp["_titular_clave"].map(comp.groupby("_titular_clave")["Titular"].first())
        por_empresa = (
            comp.groupby(["Empresa", "Trámite"], as_index=False)
            .agg(promedio=("total", "mean"), hasta_auto=("hasta_auto", "mean"), casos=("total", "count"),
                 unergy=("es_unergy", "any"), energia=("es_energia", "any"))
        )
        por_empresa["etiqueta"] = por_empresa.apply(
            lambda r: f"{r['promedio']:.0f}" + (f"  ({r['casos']} casos)" if r["casos"] > 1 else ""), axis=1)
        # una fila por empresa y trámite; si una empresa tiene los dos, la de cauce lleva "(cauce)"
        dobles = set(por_empresa["Empresa"][por_empresa["Empresa"].duplicated()])
        por_empresa["Fila"] = por_empresa.apply(
            lambda r: r["Empresa"] + ("  (cauce)" if r["Empresa"] in dobles and r["Trámite"] == "Ocupación de cauce" else ""),
            axis=1)
        orden = por_empresa.sort_values("promedio")["Fila"].tolist()
        energia = set(por_empresa.loc[por_empresa["energia"], "Fila"])
        unergy_nombres = set(por_empresa.loc[por_empresa["unergy"], "Fila"])

        # formato largo: dos tramos por barra
        corto = {"Aprovechamiento forestal": "Forestal", "Ocupación de cauce": "Cauce"}
        tramos = []
        for _, r in por_empresa.iterrows():
            auto = r["hasta_auto"] if pd.notna(r["hasta_auto"]) else None
            partes = ([("radicado → auto", auto), ("auto → resolución", r["promedio"] - auto)] if auto is not None
                      else [("sin fecha de auto", r["promedio"])])
            for orden_t, (nombre, valor) in enumerate(partes):
                tramos.append({**r.to_dict(), "Tramo": f"{corto[r['Trámite']]}: {nombre}",
                               "valor": max(valor, 0), "orden_t": orden_t, "dias_tramo": valor})
        tramos = pd.DataFrame(tramos)
        dominio = ["Forestal: radicado → auto", "Forestal: auto → resolución",
                   "Cauce: radicado → auto", "Cauce: auto → resolución",
                   "Forestal: sin fecha de auto", "Cauce: sin fecha de auto"]
        colores = ["#a3c76d", "#4d7c0f", "#93c5fd", "#2563eb", "#9ca3af", "#9ca3af"]
        presentes = [d for d in dominio if d in set(tramos["Tramo"])]
        escala = alt.Scale(domain=presentes, range=[colores[dominio.index(d)] for d in presentes])

        lista = lambda nombres: "[" + ",".join(repr(x) for x in nombres) + "]"
        color_label = (f"indexof({lista(unergy_nombres)}, datum.value) >= 0 ? '{COLOR_ALERTA}' : "
                       f"indexof({lista(energia)}, datum.value) >= 0 ? '#0f172a' : '#64748b'")
        peso_label = f"indexof({lista(unergy_nombres | energia)}, datum.value) >= 0 ? 'bold' : 'normal'"
        eje_y = alt.Y("Fila:N", sort=orden, title=None,
                      scale=alt.Scale(paddingInner=0.35, paddingOuter=0.2),
                      axis=alt.Axis(labelLimit=420, labelOverlap=False, labelFontSize=12, ticks=False,
                                    domain=False, labelColor=alt.expr(color_label),
                                    labelFontWeight=alt.expr(peso_label)))
        barras = alt.Chart(tramos).mark_bar().encode(
            y=eje_y,
            x=alt.X("sum(valor):Q", title=unidad, axis=alt.Axis(grid=False)),
            color=alt.Color("Tramo:N", scale=escala, legend=alt.Legend(orient="top", title=None, columns=2)),
            order=alt.Order("orden_t:Q"),
            tooltip=[alt.Tooltip("Empresa:N"), alt.Tooltip("Tramo:N"),
                     alt.Tooltip("dias_tramo:Q", title=f"{unidad} del tramo", format=".0f"),
                     alt.Tooltip("promedio:Q", title=f"{unidad} en total", format=".0f"),
                     alt.Tooltip("casos:Q", title="Trámites")],
        )
        borde_unergy = alt.Chart(por_empresa[por_empresa["unergy"]]).mark_bar(
            fill=None, stroke=COLOR_ALERTA, strokeWidth=3).encode(y=eje_y, x="promedio:Q")
        textos = alt.Chart(por_empresa).mark_text(align="left", dx=5, fontSize=11, color="#334155").encode(
            y=eje_y, x="promedio:Q", text="etiqueta:N")
        st.altair_chart((barras + borde_unergy + textos).properties(height=36 * len(por_empresa) + 60),
                        width="stretch")
        st.caption("Cada barra es el total de radicado a resolución. El tramo claro es hasta el auto de inicio "
                   "y el oscuro, del auto a la resolución. Verde = aprovechamiento forestal, azul = ocupación de cauce. "
                   "Nombre en rojo = Unergy; en negrita = empresas de energía.")

        # --- Trámites de pares en curso ---
        en_curso = pares[pares["Fecha resolución"].isna() & pares["Fecha radicado inicio trámite"].notna()]
        if not en_curso.empty:
            with st.expander(f"Trámites de pares todavía sin resolución ({len(en_curso)})"):
                hoy = pd.Timestamp.today().normalize()
                tabla = en_curso.assign(
                    **{f"{unidad} desde radicado": ((hoy - en_curso["Fecha radicado inicio trámite"]).dt.days / divisor).round(0)}
                )[["Titular", "Tipo de aprovechamiento", "Fecha radicado inicio trámite", "Fecha auto",
                   f"{unidad} desde radicado", "Archivo"]]
                st.dataframe(tabla.sort_values(f"{unidad} desde radicado", ascending=False),
                             width="stretch", hide_index=True)


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
    comparable = limpio[limpio["dias_radicado_resolucion"].notna()]

    if base["_anomalia"].sum():
        st.caption(f"⚠️ {int(base['_anomalia'].sum())} trámite(s) con fechas inconsistentes se excluyen de los promedios")

    por_revisar = base[base["_revisar"]]
    if not por_revisar.empty:
        with st.expander(f"⚠️ {len(por_revisar)} resolución(es) sin 'Fecha resolución' en el Excel — revisar"):
            st.caption("Son filas de tipo Resolución sin fecha de resolución; no entran en los promedios. "
                       "Muchas veces la fecha quedó en otra columna.")
            st.dataframe(por_revisar[["Titular", "Archivo", "Fecha radicado inicio trámite", "Fecha auto",
                                      "Fecha resolución"]], hide_index=True, width="stretch")

    unidad = st.segmented_control("Unidad", ["Días", "Meses"], default="Días", label_visibility="collapsed")
    unidad = unidad or "Días"

    # ---- Tarjeta 1: resumen general ----
    with st.container(border=True):
        st.subheader(f"Tiempo promedio entre etapas ({unidad.lower()})")
        prom = lambda c: formatear_dias(limpio[c].mean(), unidad) if limpio[c].notna().any() else "—"
        fila_de_circulos([
            (prom("dias_radicado_auto"), "Radicado → Auto", COLOR_PRIMARIO),
            (prom("dias_auto_visita"), "Auto → Visita", COLOR_PRIMARIO),
            (prom("dias_visita_resolucion"), "Visita → Resolución", COLOR_PRIMARIO),
            (prom("dias_radicado_resolucion"), "Radicado → Resolución (total)", COLOR_ALERTA),
        ])

        forestal = comparable[comparable["Categoría Trámite"] == "Aprovechamiento forestal"]
        forestal = forestal[forestal["Tipo de aprovechamiento"].notna()]
        if not forestal.empty:
            st.divider()
            st.markdown(f"**Según modalidad ({unidad.lower()})**")
            por_modalidad = forestal.groupby("Tipo de aprovechamiento")["dias_radicado_resolucion"].agg(["mean", "count"])
            items = [
                (formatear_dias(fila["mean"], unidad), f"{modalidad} (n={int(fila['count'])})", COLOR_PRIMARIO)
                for modalidad, fila in por_modalidad.iterrows()
            ]
            fila_de_circulos(items)
            notas = ["Único (el trámite forestal completo) tarda bastante más que Aislado (árboles urbanos aislados)."]
            num_arboles_valido = forestal[["num_arboles_num", "dias_radicado_resolucion"]].dropna()
            if len(num_arboles_valido) >= 5:
                correlacion = num_arboles_valido.corr().iloc[0, 1]
                notas.append(f"A más árboles solicitados, más tarda el trámite (correlación de {correlacion:.2f} sobre {len(num_arboles_valido)} casos).")
            st.caption(" ".join(notas))

    tab_pares, tab_ranking, tab_tendencia, tab_detalle = st.tabs(
        ["Unergy vs. pares", "Ranking de titulares", "Tendencia", "Detalle"]
    )

    with tab_pares:
        seccion_pares(limpio, unidad)

    with tab_ranking:
        with st.container(border=True):
            st.subheader("Radicado → Resolución, por titular (más rápido primero)")
            if comparable.empty:
                st.info("Todavía no hay trámites resueltos en este filtro.")
            else:
                solo_repetidos = st.checkbox("Mostrar solo titulares con más de un trámite resuelto")
                col_unidad = f"{unidad} (radicado→resolución)"
                ranking = (
                    comparable.groupby("_titular_clave")
                    .agg(Titular=("Titular", "first"), _dias=("dias_radicado_resolucion", "mean"),
                         Casos=("dias_radicado_resolucion", "count"))
                )
                if solo_repetidos:
                    ranking = ranking[ranking["Casos"] > 1]
                ranking[col_unidad] = ranking["_dias"].apply(lambda v: formatear_dias(v, unidad))
                ranking = ranking.sort_values("_dias").set_index("Titular")[[col_unidad, "Casos"]]
                if ranking.empty:
                    st.info("Ningún titular tiene más de un trámite resuelto en este filtro.")
                else:
                    st.dataframe(ranking, width="stretch")
                    if not solo_repetidos:
                        st.caption("Cuando 'Casos' es 1, el valor es ese único trámite, no un promedio.")

    with tab_tendencia:
        with st.container(border=True):
            st.subheader(f"Radicado → Resolución, tendencia ({unidad.lower()})")
            if len(comparable) < 3:
                st.info("Todavía no hay suficientes trámites resueltos en este filtro para ver una tendencia.")
            else:
                divisor = 30 if unidad == "Meses" else 1
                datos = comparable[["Fecha radicado inicio trámite", "dias_radicado_resolucion"]].copy()
                datos = datos.sort_values("Fecha radicado inicio trámite")
                datos["valor"] = datos["dias_radicado_resolucion"] / divisor
                ventana = min(5, len(datos))
                datos["promedio_movil"] = datos["valor"].rolling(ventana, min_periods=1).mean()

                grafica = (
                    alt.Chart(datos)
                    .mark_line(point=False, color=COLOR_PRIMARIO, strokeWidth=3)
                    .encode(
                        x=alt.X("Fecha radicado inicio trámite:T", title=None, axis=alt.Axis(format="%b %Y", tickCount=6)),
                        y=alt.Y("promedio_movil:Q", title=unidad),
                        tooltip=[alt.Tooltip("Fecha radicado inicio trámite:T", format="%d %b %Y"), alt.Tooltip("promedio_movil:Q", title=unidad, format=".1f")],
                    )
                    .properties(height=350)
                )
                st.altair_chart(grafica, width="stretch")
                st.caption(f"Promedio móvil de los últimos {ventana} trámites, ordenados por fecha de radicación. Los de 2024 tardaban varios cientos de días; los más recientes se resuelven mucho más rápido.")

    with tab_detalle:
        with st.container(border=True):
            columnas = [
                "Archivo", "Titular", "Categoría Trámite", "Tipo de aprovechamiento", "Seccional",
                "Fecha radicado inicio trámite", "Fecha resolución", "dias_radicado_resolucion", "Requerimientos",
            ]
            columnas = [c for c in columnas if c in base.columns]
            st.dataframe(
                base[columnas].sort_values("Fecha resolución", ascending=False),
                width="stretch", hide_index=True,
            )


if __name__ == "__main__":
    main()
