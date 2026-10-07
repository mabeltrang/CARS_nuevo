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
# nombre con tilde para mostrar, según cómo venga escrito en el Excel (sin tildes, en mayúscula)
NOMBRES_CAR = {"CORPOBOYACA": "CORPOBOYACÁ", "CARSUCRE": "CARSUCRE", "CORPOCESAR": "CORPOCESAR"}
LOGOS_CAR = {"CORPOCESAR": "https://www.corpocesar.gov.co/images/LogoCorpocesar%20SIN%20FONDO.png"}
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
    # Si la columna "Archivo" trae el número de expediente (CORPOBOYACÁ: AFAA-00034-26),
    # todos los actos de ese expediente son un solo trámite: se unen por ahí.
    exp = df["Archivo"].astype(str).str.upper().str.replace(r"[^A-Z0-9]", "", regex=True)
    es_exp = df["Archivo"].notna() & ~df["Archivo"].astype(str).str.lower().str.endswith(".pdf") & \
        exp.str.match(r"^[A-Z]{2,6}\d{3,}")
    df["_expediente"] = (df["CAR"] + "|" + exp).where(es_exp)
    if es_exp.any():
        con_exp = df[es_exp]
        agg = {c: "first" for c in con_exp.columns if c != "_expediente"}
        agg.update({c: "min" for c in COLUMNAS_FECHA if c in con_exp.columns})
        agg["Tipo"] = lambda s: " + ".join(sorted(set(s.dropna().astype(str))))
        con_exp = con_exp.groupby("_expediente", as_index=False).agg(agg)
        df = pd.concat([df[~es_exp], con_exp], ignore_index=True)
    # titular genérico ("Persona natural", "Alcaldía", vacío): no se puede saber si dos filas
    # son el mismo trámite, así que cada fila cuenta sola
    generico = df["_titular_clave"].isin(["PERSONA NATURAL", "ALCALDIA", "NAN", ""]) | df["_titular_clave"].isna()
    df["_unico"] = ""
    df.loc[generico | df["_expediente"].notna(), "_unico"] = [f"f{i}" for i in range(int((generico | df["_expediente"].notna()).sum()))]
    fechas = [c for c in COLUMNAS_FECHA if c in df.columns]
    df["_k_tit"] = (df["CAR"] + "|" + df["_titular_clave"].fillna("") + "|" + df["Categoría Trámite"].fillna("")
                    + "|" + df["_unico"])
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

    if "CAR" not in df.columns:
        df["CAR"] = "CORPOCESAR"
    car = df["CAR"].fillna("CORPOCESAR").astype(str).apply(quitar_acentos).str.upper().str.strip()
    df["CAR"] = car.map(lambda c: NOMBRES_CAR.get(c, c))

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
    # auto posterior a la resolución (la visita puede faltar, así que se revisa aparte)
    df["_anomalia"] |= (df["Fecha resolución"] - df["Fecha auto"]).dt.days < 0

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




def seccion_pares(limpio: pd.DataFrame, unidad: str, personas: pd.DataFrame | None = None) -> None:
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

        personas = personas if personas is not None else limpio.iloc[0:0]
        if misma_modalidad and modalidades_unergy:
            personas = personas[(personas["Categoría Trámite"] != "Aprovechamiento forestal")
                                | personas["Tipo de aprovechamiento"].isin(modalidades_unergy)]
        res_n = personas[personas["dias_radicado_resolucion"].notna()]

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
        NOMBRE_PERSONAS = "PERSONAS NATURALES (promedio)"
        comp = pd.concat([res_u, res_p, res_c, res_n.assign(_es_persona=True)])
        comp["_es_persona"] = comp["_es_persona"].fillna(False).astype(bool)
        comp = comp.assign(
            total=comp["dias_radicado_resolucion"] / divisor,
            hasta_auto=(comp["dias_radicado_auto"] / divisor),
            hasta_visita=((comp["Fecha visita"] - comp["Fecha radicado inicio trámite"]).dt.days / divisor),
            Trámite=comp["Categoría Trámite"],
        )
        comp["Empresa"] = comp["_titular_clave"].map(comp.groupby("_titular_clave")["Titular"].first())
        comp.loc[comp["_es_persona"], "Empresa"] = NOMBRE_PERSONAS
        por_empresa = (
            comp.groupby(["Empresa", "Trámite"], as_index=False)
            .agg(promedio=("total", "mean"), hasta_auto=("hasta_auto", "mean"), casos=("total", "count"),
                 hasta_visita=("hasta_visita", "mean"), n_visita=("hasta_visita", "count"),
                 unergy=("es_unergy", "any"), energia=("es_energia", "any"), persona=("_es_persona", "any"))
        )
        por_empresa["etiqueta"] = por_empresa.apply(
            lambda r: f"{r['promedio']:.0f}" + (f"  ({r['casos']} casos)" if r["casos"] > 1
                                                 else "  (1 caso)" if r["persona"] else ""), axis=1)
        # una fila por empresa y trámite; si una empresa tiene los dos, la de cauce lleva "(cauce)"
        dobles = set(por_empresa["Empresa"][por_empresa["Empresa"].duplicated()])
        por_empresa["Fila"] = por_empresa.apply(
            lambda r: r["Empresa"] + ("  (cauce)" if r["Empresa"] in dobles and r["Trámite"] == "Ocupación de cauce" else ""),
            axis=1)
        orden = por_empresa.sort_values("promedio")["Fila"].tolist()
        energia = set(por_empresa.loc[por_empresa["energia"] & ~por_empresa["persona"], "Fila"])
        unergy_nombres = set(por_empresa.loc[por_empresa["unergy"], "Fila"])
        personas_nombres = set(por_empresa.loc[por_empresa["persona"], "Fila"])

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
                       f"indexof({lista(personas_nombres)}, datum.value) >= 0 ? '#7c3aed' : "
                       f"indexof({lista(energia)}, datum.value) >= 0 ? '#0f172a' : '#64748b'")
        peso_label = f"indexof({lista(unergy_nombres | energia | personas_nombres)}, datum.value) >= 0 ? 'bold' : 'normal'"
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
        visitas = por_empresa[por_empresa["n_visita"] > 0]
        rayita = alt.Chart(visitas).mark_tick(color="#111827", thickness=3, size=26).encode(
            y=eje_y, x="hasta_visita:Q",
            tooltip=["Empresa", alt.Tooltip("hasta_visita:Q", title=f"{unidad} de radicado a visita", format=".0f"),
                     alt.Tooltip("n_visita:Q", title="Trámites con fecha de visita")],
        )
        textos = alt.Chart(por_empresa).mark_text(align="left", dx=5, fontSize=11, color="#334155").encode(
            y=eje_y, x="promedio:Q", text="etiqueta:N")
        st.altair_chart((barras + borde_unergy + rayita + textos).properties(height=36 * len(por_empresa) + 60),
                        width="stretch")
        st.caption("Cada barra es el total de radicado a resolución. El tramo claro es hasta el auto de inicio "
                   "y el oscuro, del auto a la resolución. Verde = aprovechamiento forestal, azul = ocupación de cauce. "
                   "Nombre en rojo = Unergy; en negrita = empresas de energía; en morado = promedio de personas naturales. "
                   "La rayita negra marca, en promedio, cuándo fue la visita (solo donde el Excel tiene fecha de visita).")

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


COLORES_CAR = ["#15803d", "#1d4ed8", "#b45309", "#7c3aed", "#be123c", "#0e7490"]


def logo_de(car: str) -> str | None:
    """Logo de la CAR: primero logos/<CAR>.png en el repo (sin tilde), luego la URL conocida."""
    for nombre in {car, quitar_acentos(car)}:
        for ext in ("png", "jpg", "jpeg", "svg", "webp"):
            ruta = os.path.join(os.path.dirname(__file__), "logos", f"{nombre}.{ext}")
            if os.path.exists(ruta):
                return ruta
    return LOGOS_CAR.get(car)


def insignia(car: str, color: str) -> str:
    """Círculo con las iniciales, para la CAR que todavía no tiene logo."""
    letras = re.sub(r"^CORPO", "", quitar_acentos(car))[:2] or car[:2]
    return (f"<div style='width:44px;height:44px;border-radius:50%;background:{color};color:white;"
            f"display:flex;align-items:center;justify-content:center;font-weight:700;font-size:15px'>{letras}</div>")


def selector_car(cars: list[str]) -> str:
    """Un botón por CAR, con su logo, en la barra lateral."""
    if st.session_state.get("car_sel") not in cars:
        st.session_state["car_sel"] = cars[0]
    st.markdown("**Corporación**")
    for i, car in enumerate(cars):
        c_logo, c_boton = st.columns([1, 3], vertical_alignment="center")
        ruta = logo_de(car)
        if ruta:
            c_logo.image(ruta, width=44)
        else:
            c_logo.markdown(insignia(car, COLORES_CAR[i % len(COLORES_CAR)]), unsafe_allow_html=True)
        activo = st.session_state["car_sel"] == car
        if c_boton.button(car, key=f"car_{car}", type="primary" if activo else "secondary", width="stretch"):
            st.session_state["car_sel"] = car
            st.rerun()
    st.divider()
    return st.session_state["car_sel"]


@st.cache_data(ttl=24 * 3600, show_spinner=False)
def _boletines_corpocesar(anios: tuple[int, ...]) -> tuple[list[dict], list[str]]:
    import corpocesar_web as web
    actos, errores = [], []
    for anio in anios:
        try:
            paginas = web.paginas_del_anio(anio)
        except Exception as e:
            errores.append(f"Índice {anio}: {e.__class__.__name__}")
            continue
        for url, etiqueta in paginas:
            try:
                actos += web.leer_pagina(url, f"{etiqueta} {anio}")
            except Exception as e:
                errores.append(f"{etiqueta} {anio}: {e.__class__.__name__}")
    return actos, errores


@st.cache_data(ttl=7 * 24 * 3600, show_spinner=False)
def _leer_pdf_corpocesar(url: str) -> dict:
    import corpocesar_web as web
    try:
        return web.leer_pdf(url)
    except Exception as e:
        return {"Error": e.__class__.__name__}


def _nucleo_nombre(titular: str) -> str:
    """'PARQUE SOLAR EL UNIÓN S.A.S E.S.P' → 'PARQUE SOLAR EL UNION' (para buscar en el boletín)."""
    n = quitar_acentos(str(titular)).upper()
    n = re.sub(r"\b(S\.?\s?A\.?\s?S\.?|S\.?\s?A\.?|E\.?\s?S\.?\s?P\.?|LTDA\.?|S\.?C\.?A\.?|SAS|ESP)\b", " ", n)
    return re.sub(r"[^A-Z0-9 ]", " ", re.sub(r"\s+", " ", n)).strip()


def seccion_buscar_en_car(car: str, limpio: pd.DataFrame) -> None:
    """Busca en la página de la CAR los actos de las empresas que más rápido salieron."""
    with st.container(border=True):
        st.subheader("¿Qué hicieron distinto los que salieron más rápido?")
        if car != "CORPOCESAR":
            st.info(f"Por ahora la búsqueda en línea funciona solo con CORPOCESAR. {car} publica un único PDF "
                    "mensual de miles de páginas: para esa CAR usa el script de boletines y sube el resultado al Excel.")
            return
        resueltos = limpio[limpio["dias_radicado_resolucion"].notna() & (limpio["Tipo de persona"] == "Empresa / Entidad")
                           & ~limpio["es_publica"]]
        ranking = (resueltos.groupby("Titular")["dias_radicado_resolucion"].mean().sort_values())
        sugeridas = list(ranking.index[:3])
        unergy = [t for t in ranking.index if "UNERGY" in quitar_acentos(t).upper()]
        st.caption("Elige empresas (por defecto, las 3 más rápidas y Unergy). La app busca sus actos en los "
                   "boletines de corpocesar.gov.co —Valledupar y seccionales— y lee los PDF para comparar "
                   "fechas y si les pidieron información adicional.")
        c1, c2 = st.columns([3, 1])
        empresas = c1.multiselect("Empresas", list(ranking.index) + sorted(set(limpio["Titular"].dropna()) - set(ranking.index)),
                                  default=sugeridas + [u for u in unergy if u not in sugeridas])
        hoy = pd.Timestamp.today().year
        anios = c2.multiselect("Años", list(range(2024, hoy + 1)), default=[hoy - 1, hoy])
        if not empresas or not anios:
            return
        if not st.button("🔎 Buscar en corpocesar.gov.co", type="primary"):
            st.caption("La primera búsqueda tarda 1–2 minutos (lee todos los boletines del año); después queda guardada 24 h.")
            return
        with st.spinner("Leyendo boletines de CORPOCESAR..."):
            actos, errores = _boletines_corpocesar(tuple(sorted(anios)))
        if errores:
            st.caption("No se pudieron leer: " + "; ".join(errores[:6]))
        claves = {e: _nucleo_nombre(e) for e in empresas}
        filas = []
        for a in actos:
            desc = quitar_acentos(a["Descripción"]).upper()
            desc = re.sub(r"[^A-Z0-9 ]", " ", desc)
            for empresa, clave in claves.items():
                if clave and clave in re.sub(r"\s+", " ", desc):
                    filas.append({"Empresa": empresa, **a})
        if not filas:
            st.warning("No encontré actos de esas empresas en los boletines de esos años.")
            return
        hallados = pd.DataFrame(filas).drop_duplicates(subset=["Empresa", "PDF", "Número"]).sort_values(["Empresa", "Fecha"])
        st.success(f"{len(hallados)} actos encontrados de {hallados['Empresa'].nunique()} empresas.")
        st.dataframe(hallados[["Empresa", "Fecha", "Categoría", "Descripción", "PDF"]], hide_index=True, width="stretch",
                     column_config={"PDF": st.column_config.LinkColumn("PDF", display_text="Abrir")})

        resoluciones = hallados[hallados["PDF"].str.contains("RESOL", case=False, na=False)]
        if resoluciones.empty:
            return
        st.markdown("**Lo que dicen las resoluciones**")
        with st.spinner(f"Leyendo {min(len(resoluciones), 20)} resoluciones..."):
            leidas = [{"Empresa": r["Empresa"], "Resolución": r["Fecha"], **_leer_pdf_corpocesar(r["PDF"]), "PDF": r["PDF"]}
                      for _, r in resoluciones.head(20).iterrows()]
        tabla = pd.DataFrame(leidas)
        for c in ["Radicado", "Auto de inicio", "Visita", "Resolución"]:
            if c in tabla:
                tabla[c] = pd.to_datetime(tabla[c], errors="coerce")
        if {"Radicado", "Resolución"} <= set(tabla.columns):
            tabla["Días radicado→resolución"] = (tabla["Resolución"] - tabla["Radicado"]).dt.days
        if {"Auto de inicio", "Visita"} <= set(tabla.columns):
            tabla["Días auto→visita"] = (tabla["Visita"] - tabla["Auto de inicio"]).dt.days
        st.dataframe(tabla, hide_index=True, width="stretch",
                     column_config={"PDF": st.column_config.LinkColumn("PDF", display_text="Abrir")})
        st.caption("Compara sobre todo la columna de información adicional: un requerimiento suele sumar meses. "
                   "Abre los PDF de las más rápidas para ver qué entregaron desde el radicado.")


def main() -> None:
    col_logo, col_titulo = st.columns([1, 6])
    logo, titulo = col_logo.empty(), col_titulo.empty()

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
    cars = sorted(df["CAR"].dropna().unique(), key=lambda c: (c != "CORPOCESAR", c))

    with st.sidebar:
        st.header("Filtros")
        car_sel = selector_car(cars)
        solo_empresas = st.checkbox("Solo empresas / entidades", value=True)
        categorias_validas = ["Aprovechamiento forestal", "Ocupación de cauce"]
        cat_sel = [c for c in categorias_validas if st.checkbox(c, value=True)]

    df = df[df["CAR"] == car_sel]
    titulo.title(f"Tiempos de trámite — {car_sel}")
    ruta_logo = logo_de(car_sel)
    if ruta_logo:
        logo.image(ruta_logo, width=90)

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

    tab_pares, tab_buscar, tab_ranking, tab_tendencia, tab_detalle = st.tabs(
        ["Unergy vs. pares", "Buscar en la CAR", "Ranking de titulares", "Tendencia", "Detalle"]
    )

    with tab_buscar:
        seccion_buscar_en_car(car_sel, limpio)

    with tab_pares:
        # personas naturales como referencia, aunque el filtro "Solo empresas" esté activo
        personas = df[df["Categoría Trámite"].isin(cat_sel) & (df["Tipo de persona"] == "Persona natural")
                      & ~df["_anomalia"]]
        seccion_pares(limpio, unidad, personas)

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
