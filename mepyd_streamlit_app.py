"""
Aplicación Streamlit para analizar el dataset oficial del MEPyD.

Instalación de dependencias:
pip install streamlit pandas plotly requests openpyxl
"""

from io import BytesIO, StringIO
import re

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import requests
import streamlit as st


st.set_page_config(page_title="MEPyD - Ejecución de Proyectos de Inversión", layout="wide")

st.markdown(
    """
    <style>
    .stApp {
        background: radial-gradient(circle at top left, #081f33 0%, #061623 35%, #020710 100%);
        color: #e6f2ff;
    }
    .block-container {
        padding-top: 2rem;
        padding-bottom: 2rem;
        background: rgba(5, 17, 36, 0.85);
        border-radius: 24px;
        box-shadow: 0 20px 50px rgba(0, 0, 0, 0.35);
        border: 1px solid rgba(255, 255, 255, 0.06);
    }
    .st-bf {
        color: #e6f2ff;
    }
    .stTextInput label, .stSelectbox label, .stMultiSelect label, .stButton button {
        color: #b8d4ff !important;
    }
    .stButton>button {
        background-color: #0a4d8c !important;
        color: white !important;
        border: 1px solid rgba(255,255,255,0.12) !important;
    }
    .stDownloadButton>button {
        background-color: #0a4d8c !important;
        color: white !important;
        border: 1px solid rgba(255,255,255,0.12) !important;
    }
    .css-1d391kg .stExpanderHeader {
        background: rgba(255,255,255,0.04);
        border-radius: 16px;
    }
    .css-1d391kg .stExpanderContent {
        background: rgba(255,255,255,0.03);
        border-radius: 16px;
    }
    .stMetric {
        background: rgba(255,255,255,0.05);
        border-radius: 18px;
        padding: 1rem;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

CSV_URL = "https://mepyd.gob.do/download/17164/informe-de-ejecucion-inversion-publica/420911/estadisticas-informe-de-ejecucion-de-proyectos-de-inversion-mepyd-2018-2024-3.csv"
XLSX_URL = "https://mepyd.gob.do/download/17164/informe-de-ejecucion-inversion-publica/420909/estadisticas-informe-de-ejecucion-de-proyectos-de-inversion-mepyd-2018-2024.xlsx"

EXPECTED_COLUMNS = [
    "PERIODO",
    "AMBITO INSTITUCIONAL",
    "SNIP",
    "NOMBRE PROYECTO",
    "INSTITUCION EJECUTORA",
    "SITUACION PRESUPUESTARIA",
    "TIPOLOGIA DE PROYECTO",
    "EJE END",
    "OBJETIVO GENERAL END",
    "OBJETIVO ESPECIFICO END",
    "LINEA DE ACCION END",
    "NOMBRE CORTO ODS",
    "FUNCION",
    "SUB FUNCION",
    "FECHA INICIO PROYECTO",
    "FECHA FIN PROYECTO",
    "REGION O PROVINCIA",
    "PORCENTAJE",
    "PRESUPUESTO VIGENTE FONDO GENERAL",
    "PRESUPUESTO VIGENTE CREDITO EXTERNO",
    "PRESUPUESTO VIGENTE DONACIONES",
    "TOTAL PRESUPUESTO VIGENTE",
    "EJECUCION FONDO GENERAL",
    "EJECUCION CREDITO EXTERNO",
    "EJECUCION DONACIONES",
    "TOTAL EJECUCION",
]

FINANCIAL_COLUMNS = [
    "PRESUPUESTO VIGENTE FONDO GENERAL",
    "PRESUPUESTO VIGENTE CREDITO EXTERNO",
    "PRESUPUESTO VIGENTE DONACIONES",
    "TOTAL PRESUPUESTO VIGENTE",
    "EJECUCION FONDO GENERAL",
    "EJECUCION CREDITO EXTERNO",
    "EJECUCION DONACIONES",
    "TOTAL EJECUCION",
]

PROVINCE_COORDINATES = {
    "santo domingo": (-69.9312, 18.4861),
    "distrito nacional": (-69.9284, 18.4740),
    "santiago": (-70.7079, 19.4550),
    "la vega": (-70.6064, 19.2109),
    "puerto plata": (-70.6878, 19.7928),
    "san cristobal": (-70.4445, 18.4165),
    "la altagracia": (-68.5456, 18.4308),
    "san pedro de macoris": (-69.2963, 18.4531),
    "monte plata": (-69.0167, 18.8133),
    "peravia": (-70.2783, 18.4509),
    "san juan": (-71.2519, 18.4556),
    "azua": (-70.7321, 18.4569),
    "montecristi": (-71.6333, 19.8420),
    "sanchez ramirez": (-70.2436, 19.0412),
    "barahona": (-71.1447, 18.2108),
    "duarte": (-69.3218, 19.1584),
    "maría trinidad sánchez": (-69.4034, 19.3898),
    "maría trinidad sanchez": (-69.4034, 19.3898),
    "samaná": (-69.3306, 19.2226),
    "puerto plata": (-70.6878, 19.7928),
    "peravia": (-70.2783, 18.4509),
    "santiago rodriquez": (-71.1328, 19.3994),
    "santiago rodriguez": (-71.1328, 19.3994),
    "valverde": (-71.2570, 19.6770),
    "hato mayor": (-69.2814, 18.7650),
    "el seibo": (-68.7074, 18.7081),
    "monseñor nouel": (-70.0823, 18.8319),
    "monseñor nouel": (-70.0823, 18.8319),
    "independencia": (-71.8500, 18.4995),
    "pedernales": (-71.7770, 17.9690),
    "bahoruco": (-71.5000, 18.3667),
    "elias piña": (-71.7667, 18.5667),
    "damajabon": (-71.6833, 19.5333),
    "dajaabon": (-71.6833, 19.5333),
    "espallat": (-70.2460, 19.3447),
    "espaillat": (-70.2460, 19.3447),
}


def normalize_column_name(value: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(value).lower())


def normalize_numeric(series: pd.Series) -> pd.Series:
    if pd.api.types.is_numeric_dtype(series):
        return series.astype(float)

    cleaned = (
        series.astype(str)
        .str.replace(r"[^0-9,\.\-]", "", regex=True)
        .str.replace(",", "", regex=False)
        .replace({"": pd.NA})
    )
    return pd.to_numeric(cleaned, errors="coerce").astype(float)


def load_data_from_upload(uploaded_file) -> pd.DataFrame:
    if uploaded_file is None:
        return pd.DataFrame()

    name = uploaded_file.name.lower()
    if name.endswith(".csv"):
        uploaded_file.seek(0)
        raw_bytes = uploaded_file.read()
        text = None

        for encoding in ("utf-8-sig", "utf-8", "latin-1", "iso-8859-1"):
            try:
                text = raw_bytes.decode(encoding)
                break
            except UnicodeDecodeError:
                continue

        if text is None:
            text = raw_bytes.decode("utf-8", errors="ignore")

        for sep in [",", ";", "\t", "|"]:
            try:
                return pd.read_csv(StringIO(text), sep=sep, engine="python")
            except Exception:
                continue
        raise ValueError("No fue posible leer el archivo CSV cargado.")

    if name.endswith((".xlsx", ".xls")):
        uploaded_file.seek(0)
        return pd.read_excel(uploaded_file)

    raise ValueError("Formato de archivo no soportado. Use CSV o Excel.")


@st.cache_data(show_spinner=False)
def load_remote_data(force_refresh: bool = False) -> pd.DataFrame:
    last_error = None
    for url, loader in [(CSV_URL, "csv"), (XLSX_URL, "xlsx")]:
        try:
            response = requests.get(url, timeout=90)
            response.raise_for_status()

            if loader == "csv":
                raw_text = response.content.decode("utf-8", errors="ignore")
                for encoding in ("utf-8-sig", "utf-8", "latin-1", "iso-8859-1"):
                    try:
                        raw_text = response.content.decode(encoding)
                        break
                    except UnicodeDecodeError:
                        continue

                for sep in [",", ";", "\t", "|"]:
                    try:
                        return pd.read_csv(StringIO(raw_text), sep=sep, engine="python")
                    except Exception:
                        continue
            else:
                return pd.read_excel(BytesIO(response.content))
        except Exception as exc:
            last_error = exc

    raise RuntimeError(f"No se pudo descargar el dataset desde MEPyD: {last_error}")


def map_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(col).strip() for col in df.columns]

    normalized_map = {normalize_column_name(col): col for col in df.columns}
    mapped = {}

    for canonical in EXPECTED_COLUMNS:
        key = normalize_column_name(canonical)
        if key in normalized_map:
            mapped[canonical] = normalized_map[key]

    renamed = df.rename(columns=mapped)
    for col in EXPECTED_COLUMNS:
        if col not in renamed.columns:
            renamed[col] = pd.NA

    return renamed[EXPECTED_COLUMNS]


def clean_dataset(df: pd.DataFrame) -> pd.DataFrame:
    cleaned = df.copy()
    cleaned = map_columns(cleaned)

    for column in FINANCIAL_COLUMNS:
        cleaned[column] = normalize_numeric(cleaned[column])

    cleaned["PORCENTAJE"] = pd.to_numeric(
        cleaned["PORCENTAJE"].astype(str).str.replace(",", ".", regex=False).str.replace("%", "", regex=False),
        errors="coerce",
    )

    cleaned["PERIODO"] = pd.to_numeric(cleaned["PERIODO"], errors="coerce").astype("Int64")
    cleaned["FECHA INICIO PROYECTO"] = pd.to_datetime(cleaned["FECHA INICIO PROYECTO"], errors="coerce")
    cleaned["FECHA FIN PROYECTO"] = pd.to_datetime(cleaned["FECHA FIN PROYECTO"], errors="coerce")

    cleaned["% Ejecución Financiera"] = (
        cleaned["TOTAL EJECUCION"] / cleaned["TOTAL PRESUPUESTO VIGENTE"] * 100
    ).replace([float("inf"), float("-inf")], pd.NA)
    cleaned["Subejecución / Brecha"] = (
        cleaned["TOTAL PRESUPUESTO VIGENTE"] - cleaned["TOTAL EJECUCION"]
    )

    return cleaned


def compute_accountability_by_region(df: pd.DataFrame) -> pd.DataFrame:
    region_df = (
        df.groupby("REGION O PROVINCIA", as_index=False)
        .agg(
            Presupuesto=("TOTAL PRESUPUESTO VIGENTE", "sum"),
            Ejecucion=("TOTAL EJECUCION", "sum"),
            Eficiencia=("% Ejecución Financiera", "mean"),
            Brecha=("Subejecución / Brecha", "sum"),
            Proyectos=("SNIP", "nunique"),
        )
    )
    region_df["Brecha"] = region_df["Brecha"].fillna(0)
    region_df["Eficiencia"] = region_df["Eficiencia"].fillna(0)
    region_df["Riesgo de Rendición"] = (
        (100 - region_df["Eficiencia"]) * (region_df["Brecha"] / region_df["Presupuesto"].replace({0: np.nan}))
    ).fillna(0)
    region_df["Índice Ciudadano"] = (100 - region_df["Eficiencia"]) + region_df["Riesgo de Rendición"]
    return region_df.sort_values(["Índice Ciudadano", "Brecha"], ascending=[False, False])


def compute_ods_priorities(df: pd.DataFrame) -> pd.DataFrame:
    ods_df = (
        df.groupby("NOMBRE CORTO ODS", as_index=False)
        .agg(
            Presupuesto=("TOTAL PRESUPUESTO VIGENTE", "sum"),
            Ejecucion=("TOTAL EJECUCION", "sum"),
            Eficiencia=("% Ejecución Financiera", "mean"),
            Proyectos=("SNIP", "nunique"),
        )
    )
    ods_df["Brecha"] = ods_df["Presupuesto"] - ods_df["Ejecucion"]
    ods_df["Urgencia"] = (
        ods_df["Brecha"] * (100 - ods_df["Eficiencia"]) / ods_df["Presupuesto"].replace({0: np.nan})
    ).fillna(0)
    return ods_df.sort_values("Urgencia", ascending=False)


def get_region_coordinates(region: str):
    if not isinstance(region, str) or not region.strip():
        return None
    key = region.lower().strip()
    key = key.replace("á", "a").replace("é", "e").replace("í", "i").replace("ó", "o").replace("ú", "u")
    key = key.replace("ñ", "n").replace("\u00f1", "n").replace("\u00e1", "a").replace("\u00e9", "e")
    key = key.replace("\u00ed", "i").replace("\u00f3", "o").replace("\u00fa", "u")
    key = key.replace("provincia ", "").replace("region ", "").strip()
    if key in PROVINCE_COORDINATES:
        return PROVINCE_COORDINATES[key]
    for name, coords in PROVINCE_COORDINATES.items():
        if name in key or key in name:
            return coords
    return None


def load_dataset(uploaded_file=None, force_refresh: bool = False) -> pd.DataFrame:
    if uploaded_file is not None:
        raw_df = load_data_from_upload(uploaded_file)
    else:
        raw_df = load_remote_data(force_refresh=force_refresh)

    return clean_dataset(raw_df)


st.markdown(
    """
    <div style="background: linear-gradient(135deg, rgba(5, 26, 58, 0.96), rgba(6, 20, 47, 0.96)); padding: 1.6rem 1.8rem; border-radius: 24px; margin-bottom: 1rem; border: 1px solid rgba(255,255,255,0.08);">
    <h1 style="color:#f7fbff; margin:0; font-size:2.2rem; letter-spacing:0.6px; font-weight:800;">MEPyD Ciudadano: Fiscalización y Rendición</h1>
    <p style="color:#c9dbf2; margin:0.65rem 0 0 0; font-size:1rem; line-height:1.5; max-width:780px;">Herramienta para convertir datos de inversión pública en alertas de servicio, prioridades ciudadanas y demandas de transparencia en la República Dominicana.</p>
    </div>
    """,
    unsafe_allow_html=True,
)

with st.sidebar:
    st.header("Datos")
    st.write("Carga el dataset oficial desde el portal de datos abiertos de RD o sube un archivo manualmente.")
    uploaded_file = st.file_uploader("Subir CSV/XLSX manualmente", type=["csv", "xlsx", "xls"])

    if st.button("Recargar datos", use_container_width=True):
        st.cache_data.clear()
        st.session_state["force_refresh"] = True
        st.rerun()

    st.markdown("---")
    st.caption("Datos abiertos RD · Ministerio de Economía, Planificación y Desarrollo")

force_refresh = st.session_state.pop("force_refresh", False)

try:
    df = load_dataset(uploaded_file=uploaded_file, force_refresh=force_refresh)
except Exception as exc:
    st.error(f"No fue posible cargar el dataset: {exc}")
    st.stop()

if df.empty:
    st.warning("El dataset cargado está vacío.")
    st.stop()

# Filtros interactivos ocultos en un expander
with st.expander("Opciones de filtrado avanzadas", expanded=False):
    period_options = sorted([int(x) for x in df["PERIODO"].dropna().unique().tolist()])
    selected_periods = st.multiselect("PERIODO", options=period_options, default=period_options)

    institution_options = sorted([str(x) for x in df["INSTITUCION EJECUTORA"].dropna().unique().tolist()])
    selected_institutions = st.multiselect("INSTITUCION EJECUTORA", options=institution_options, default=institution_options)

    region_options = sorted([str(x) for x in df["REGION O PROVINCIA"].dropna().unique().tolist()])
    selected_regions = st.multiselect("REGION O PROVINCIA", options=region_options, default=region_options)

    ods_options = sorted([str(x) for x in df["NOMBRE CORTO ODS"].dropna().unique().tolist()])
    selected_ods = st.multiselect("NOMBRE CORTO ODS", options=ods_options, default=ods_options)

filtered_df = df[
    df["PERIODO"].isin(selected_periods)
    & df["INSTITUCION EJECUTORA"].isin(selected_institutions)
    & df["REGION O PROVINCIA"].isin(selected_regions)
    & df["NOMBRE CORTO ODS"].isin(selected_ods)
]

filtered_df = filtered_df.dropna(subset=["TOTAL PRESUPUESTO VIGENTE", "TOTAL EJECUCION"], how="all")

# Métricas avanzadas de desempeño
presupuesto_total = float(filtered_df["TOTAL PRESUPUESTO VIGENTE"].sum())
ejecucion_total = float(filtered_df["TOTAL EJECUCION"].sum())
porcentaje_global = (ejecucion_total / presupuesto_total * 100) if presupuesto_total else 0.0
cantidad_proyectos = int(filtered_df.shape[0])
promedio_financiero = float(filtered_df["% Ejecución Financiera"].mean(skipna=True) or 0.0)
promedio_fisico = float(filtered_df["PORCENTAJE"].mean(skipna=True) or 0.0)

situacion_counts = filtered_df["SITUACION PRESUPUESTARIA"].value_counts().head(5)

risk_status = np.select(
    [
        (filtered_df["% Ejecución Financiera"] < 70) & (filtered_df["PORCENTAJE"] < 70),
        (filtered_df["% Ejecución Financiera"] < 80) | (filtered_df["PORCENTAJE"] < 80),
    ],
    ["Alto riesgo", "Riesgo medio"],
    default="Bajo riesgo",
)
filtered_df["Riesgo"] = risk_status
high_risk_projects = int((filtered_df["Riesgo"] == "Alto riesgo").sum())
medium_risk_projects = int((filtered_df["Riesgo"] == "Riesgo medio").sum())
low_risk_projects = int((filtered_df["Riesgo"] == "Bajo riesgo").sum())

st.subheader("Resumen ejecutivo")
metric_col1, metric_col2, metric_col3, metric_col4 = st.columns(4)
metric_col1.metric("Presupuesto Total", f"RD$ {presupuesto_total:,.0f}")
metric_col2.metric("Ejecución Total", f"RD$ {ejecucion_total:,.0f}")
metric_col3.metric("Ejecución Global", f"{porcentaje_global:,.2f}%")
metric_col4.metric("Proyectos", f"{cantidad_proyectos:,}")

risk_col1, risk_col2, risk_col3, risk_col4 = st.columns(4)
risk_col1.metric("Promedio Ejecutado", f"{promedio_financiero:,.2f}%")
risk_col2.metric("Avance Físico Medio", f"{promedio_fisico:,.2f}%")
risk_col3.metric("Alto riesgo", f"{high_risk_projects:,}")
risk_col4.metric("Riesgo medio", f"{medium_risk_projects:,}")

st.markdown("---")

risk_summary = (
    filtered_df["Riesgo"].value_counts(normalize=True).rename_axis("Riesgo").reset_index(name="Porcentaje")
)
compliance_summary = (
    filtered_df.groupby("SITUACION PRESUPUESTARIA", as_index=False)
    .agg(Eficiencia=("% Ejecución Financiera", "mean"))
    .sort_values("Eficiencia", ascending=False)
)

col_pie, col_bar = st.columns(2)
with col_pie:
    if not risk_summary.empty:
        fig_risk = px.pie(
            risk_summary,
            names="Riesgo",
            values="Porcentaje",
            title="Distribución de riesgo de proyectos",
            hole=0.45,
            template="plotly_white",
        )
        fig_risk.update_traces(textposition="inside", textinfo="percent+label")
        st.plotly_chart(fig_risk, use_container_width=True)
    else:
        st.info("No hay datos de riesgo para mostrar.")

with col_bar:
    if not compliance_summary.empty:
        fig_compliance = px.bar(
            compliance_summary,
            x="Eficiencia",
            y="SITUACION PRESUPUESTARIA",
            orientation="h",
            color="Eficiencia",
            title="Eficiencia financiera por situación presupuestaria",
            text="Eficiencia",
            template="plotly_white",
            color_continuous_scale="Viridis",
        )
        fig_compliance.update_traces(texttemplate="%{text:.1f}%")
        fig_compliance.update_layout(xaxis_title="% Ejecución Financiera", showlegend=False)
        st.plotly_chart(fig_compliance, use_container_width=True)
    else:
        st.info("No hay datos de situación presupuestaria para mostrar.")

st.markdown("---")

with st.expander("Visión urgente del desempeño", expanded=True):
    col_a, col_b = st.columns((1, 2))
    with col_a:
        st.markdown("**Situación presupuestaria más común**")
        st.table(situacion_counts.rename_axis("Situación").reset_index(name="Proyectos"))

    with col_b:
        top_underperformers = (
            filtered_df.assign(Brecha=filtered_df["Subejecución / Brecha"])
            .query("`% Ejecución Financiera` < 60")
            .sort_values(["Subejecución / Brecha", "% Ejecución Financiera"], ascending=[False, True])
            .head(5)
            [["NOMBRE PROYECTO", "INSTITUCION EJECUTORA", "REGION O PROVINCIA", "% Ejecución Financiera", "Subejecución / Brecha"]]
        )
        st.markdown("**Proyectos críticos de baja ejecución**")
        st.dataframe(top_underperformers, use_container_width=True, hide_index=True)

st.markdown("---")

# Tablero principal con análisis profundo
tabs = st.tabs(["Análisis Estratégico", "Alertas y Riesgos", "Territorio y Servicios", "Plan de Acción Ciudadana"])

with tabs[0]:
    st.markdown("### Rendimiento por período y financiamiento")
    period_summary = (
        filtered_df.groupby("PERIODO", as_index=False)
        .agg(
            Presupuesto=("TOTAL PRESUPUESTO VIGENTE", "sum"),
            Ejecucion=("TOTAL EJECUCION", "sum"),
            Proyectos=("SNIP", "nunique"),
            EjecucionMedia=("% Ejecución Financiera", "mean"),
        )
    )
    line_fig = px.line(
        period_summary,
        x="PERIODO",
        y=["Presupuesto", "Ejecucion"],
        markers=True,
        title="Tendencia de Presupuesto y Ejecución por Año",
        template="plotly_white",
    )
    line_fig.update_layout(yaxis_title="RD$", legend_title="Serie")
    st.plotly_chart(line_fig, use_container_width=True)

    st.markdown("### Estructura de proyectos por tipología y función")
    category_summary = (
        filtered_df.groupby(["TIPOLOGIA DE PROYECTO", "FUNCION"], as_index=False)
        .agg(
            Proyectos=("SNIP", "nunique"),
            Presupuesto=("TOTAL PRESUPUESTO VIGENTE", "sum"),
        )
        .sort_values("Presupuesto", ascending=False)
    )
    if not category_summary.empty:
        sunburst = px.sunburst(
            category_summary,
            path=["TIPOLOGIA DE PROYECTO", "FUNCION"],
            values="Presupuesto",
            color="Proyectos",
            color_continuous_scale="Blues",
            title="Estructura de inversión por tipología de proyecto y función",
            template="plotly_white",
        )
        st.plotly_chart(sunburst, use_container_width=True)
    else:
        st.info("No hay datos suficientes para generar la visualización de tipologías.")

    st.markdown("### Composición del financiamiento por fuente")
    source_summary = (
        filtered_df[
            [
                "PRESUPUESTO VIGENTE FONDO GENERAL",
                "PRESUPUESTO VIGENTE CREDITO EXTERNO",
                "PRESUPUESTO VIGENTE DONACIONES",
                "EJECUCION FONDO GENERAL",
                "EJECUCION CREDITO EXTERNO",
                "EJECUCION DONACIONES",
            ]
        ]
        .sum()
        .reset_index()
        .rename(columns={"index": "Fuente", 0: "Monto"})
    )
    source_summary["Tipo"] = source_summary["Fuente"].apply(lambda v: "Presupuesto" if "PRESUPUESTO" in v else "Ejecución")
    source_summary["Fuente"] = source_summary["Fuente"].str.replace("PRESUPUESTO VIGENTE ", "", regex=False).str.replace("EJECUCION ", "", regex=False)
    bar_source = px.bar(
        source_summary,
        x="Monto",
        y="Fuente",
        color="Tipo",
        orientation="h",
        barmode="group",
        title="Totales por Fuente de Financiamiento",
        template="plotly_white",
    )
    st.plotly_chart(bar_source, use_container_width=True)

    st.markdown("### Top 10 instituciones por volumen y eficiencia")
    inst_perf = (
        filtered_df.groupby("INSTITUCION EJECUTORA", as_index=False)
        .agg(
            Presupuesto=("TOTAL PRESUPUESTO VIGENTE", "sum"),
            Ejecucion=("TOTAL EJECUCION", "sum"),
            EjecucionMedia=("% Ejecución Financiera", "mean"),
        )
        .sort_values("Presupuesto", ascending=False)
        .head(10)
    )
    inst_perf["Eficiencia"] = inst_perf["EjecucionMedia"]
    bar_inst = px.bar(
        inst_perf,
        x="Presupuesto",
        y="INSTITUCION EJECUTORA",
        orientation="h",
        color="Eficiencia",
        title="Top 10 Instituciones por Presupuesto y Eficiencia Financiera",
        template="plotly_white",
        color_continuous_scale="blues",
    )
    st.plotly_chart(bar_inst, use_container_width=True)

with tabs[1]:
    st.markdown("### Rendimiento proyectado por proyecto")
    scatter = px.scatter(
        filtered_df,
        x="PORCENTAJE",
        y="% Ejecución Financiera",
        size="TOTAL PRESUPUESTO VIGENTE",
        color="Riesgo",
        hover_name="NOMBRE PROYECTO",
        hover_data={
            "SNIP": True,
            "INSTITUCION EJECUTORA": True,
            "TOTAL PRESUPUESTO VIGENTE": ":,.0f",
            "TOTAL EJECUCION": ":,.0f",
        },
        title="Proyectos: avance físico vs ejecución financiera",
        template="plotly_white",
    )
    scatter.update_layout(
        xaxis_title="% Avance Físico",
        yaxis_title="% Ejecución Financiera",
        shapes=[
            dict(type="line", x0=80, x1=80, y0=0, y1=100, line=dict(color="gray", dash="dash")),
            dict(type="line", x0=0, x1=100, y0=80, y1=80, line=dict(color="gray", dash="dash")),
        ],
    )
    st.plotly_chart(scatter, use_container_width=True)

    compliance_score = float(filtered_df["% Ejecución Financiera"].mean(skipna=True) or 0.0)
    fig_gauge = go.Figure(go.Indicator(
        mode="gauge+number+delta",
        value=compliance_score,
        domain={"x": [0, 1], "y": [0, 1]},
        title={"text": "Cumplimiento Financiero Global"},
        delta={"reference": 80, "increasing": {"color": "green"}},
        gauge={
            "axis": {"range": [0, 100]},
            "bar": {"color": "#1f77b4"},
            "steps": [
                {"range": [0, 60], "color": "#ff6961"},
                {"range": [60, 80], "color": "#f8d568"},
                {"range": [80, 100], "color": "#77dd77"},
            ],
            "threshold": {
                "line": {"color": "red", "width": 4},
                "thickness": 0.75,
                "value": 80,
            },
        },
    ))
    st.plotly_chart(fig_gauge, use_container_width=True)

    st.markdown("### Proyectos con mayor brecha y baja ejecución")
    alert_table = (
        filtered_df.assign(
            Brecha=filtered_df["Subejecución / Brecha"],
            EjecucionFinanciera=filtered_df["% Ejecución Financiera"],
        )
        .query("EjecucionFinanciera < 70")
        .sort_values(["Brecha", "EjecucionFinanciera"], ascending=[False, True])
        .head(15)
        [[
            "NOMBRE PROYECTO",
            "SNIP",
            "INSTITUCION EJECUTORA",
            "REGION O PROVINCIA",
            "PRESUPUESTO VIGENTE FONDO GENERAL",
            "PRESUPUESTO VIGENTE CREDITO EXTERNO",
            "PRESUPUESTO VIGENTE DONACIONES",
            "TOTAL PRESUPUESTO VIGENTE",
            "TOTAL EJECUCION",
            "% Ejecución Financiera",
            "PORCENTAJE",
            "Subejecución / Brecha",
        ]]
    )
    st.dataframe(alert_table, use_container_width=True, hide_index=True)

with tabs[2]:
    st.markdown("### Mapa de riesgo ciudadano y financiero por provincia/región")
    region_priority = compute_accountability_by_region(filtered_df)
    region_priority["Clasificación de Riesgo"] = np.select(
        [
            region_priority["Índice Ciudadano"] < 20,
            region_priority["Índice Ciudadano"] < 40,
            region_priority["Índice Ciudadano"] < 60,
        ],
        ["Bajo", "Moderado", "Alto"],
        default="Urgente",
    )
    st.markdown(
        "Este mapa prioriza las regiones donde la baja ejecución financiera y la brecha presupuestaria "
        "se combinan con un alto riesgo de rendición. Las burbujas más grandes muestran el volumen de inversión "
        "y el color refleja el riesgo ciudadano."
    )

    metric_selector = st.selectbox(
        "Métrica del mapa",
        ["Índice Ciudadano", "Brecha", "Eficiencia", "Proyectos"],
        index=0,
        help="Selecciona la métrica que quieres visualizar en color sobre el mapa.",
    )

    map_df = region_priority.copy()
    map_df["coords"] = map_df["REGION O PROVINCIA"].apply(get_region_coordinates)
    map_df = map_df.dropna(subset=["coords"]).copy()
    if not map_df.empty:
        map_df["lon"] = map_df["coords"].apply(lambda c: c[0])
        map_df["lat"] = map_df["coords"].apply(lambda c: c[1])
        size_field = "Brecha" if metric_selector != "Proyectos" else "Proyectos"
        figure_map = px.scatter_mapbox(
            map_df,
            lon="lon",
            lat="lat",
            size=size_field,
            color=metric_selector,
            hover_name="REGION O PROVINCIA",
            hover_data={
                "Presupuesto": ":,.0f",
                "Ejecucion": ":,.0f",
                "Eficiencia": ":.1f",
                "Proyectos": True,
                "Índice Ciudadano": ":.1f",
                "Riesgo de Rendición": ":.1f",
            },
            zoom=6,
            center={"lat": 18.8, "lon": -69.8},
            title="Riesgo ciudadano y financiamiento por provincia/región",
            template="plotly_white",
            mapbox_style="open-street-map",
            color_continuous_scale="thermal",
            size_max=45,
            height=620,
        )
        figure_map.update_layout(
            mapbox=dict(
                center={"lat": 18.8, "lon": -69.8},
                zoom=6,
            ),
            margin=dict(l=0, r=0, t=40, b=0),
        )
        st.plotly_chart(figure_map, use_container_width=True)

        top_metrics, top_table = st.columns((1, 1))
        with top_metrics:
            highest_risk = region_priority.iloc[0]
            lowest_efficiency = region_priority.sort_values("Eficiencia").iloc[0]
            highest_brecha = region_priority.sort_values("Brecha", ascending=False).iloc[0]
            st.metric("Región más urgente", highest_risk["REGION O PROVINCIA"])
            st.metric("Mayor brecha RD$", f"{highest_brecha['Brecha']:,.0f}")
            st.metric("Peor eficiencia", f"{lowest_efficiency['Eficiencia']:.1f}%")
        with top_table:
            st.markdown("#### Regiones en foco")
            st.dataframe(
                region_priority[
                    ["REGION O PROVINCIA", "Clasificación de Riesgo", "Eficiencia", "Brecha", "Índice Ciudadano"]
                ]
                .rename(
                    columns={
                        "REGION O PROVINCIA": "Región/Provincia",
                        "Eficiencia": "Eficiencia (%)",
                        "Brecha": "Brecha RD$",
                        "Índice Ciudadano": "Índice Ciudadano",
                    }
                )
                .head(8),
                use_container_width=True,
                hide_index=True,
            )

        st.markdown("#### Interpretación del riesgo ciudadano")
        st.markdown(
            "- Los valores altos del Índice Ciudadano indican regiones donde la ejecución financiera baja y la brecha presupuestaria generan mayor demanda de transparencia."
            "\n" "- La clasificación de riesgo ayuda a priorizar la presión social y la exigencia de informes públicos."
            "\n" "- Usa este panel para enfocar auditorías, solicitudes de acceso a la información y denuncias ciudadanas."
        )
    else:
        st.info("No se encontraron coordenadas válidas para las regiones disponibles.")

    st.markdown("### Inversión por ODS")
    ods_summary = (
        filtered_df.groupby("NOMBRE CORTO ODS", as_index=False)
        .agg(Presupuesto=("TOTAL PRESUPUESTO VIGENTE", "sum"), Ejecucion=("TOTAL EJECUCION", "sum"))
        .sort_values("Presupuesto", ascending=False)
    )
    if not ods_summary.empty:
        donut = px.pie(
            ods_summary,
            values="Presupuesto",
            names="NOMBRE CORTO ODS",
            title="Distribución de presupuesto por ODS",
            hole=0.45,
            template="plotly_white",
        )
        st.plotly_chart(donut, use_container_width=True)
    else:
        st.info("No hay datos suficientes para generar el donut de ODS.")

    st.markdown("### Alineación estratégica: ODS y Eje END")
    treemap_data = (
        filtered_df.groupby(["NOMBRE CORTO ODS", "EJE END"], as_index=False)
        .agg(Presupuesto=("TOTAL PRESUPUESTO VIGENTE", "sum"), Ejecucion=("TOTAL EJECUCION", "sum"))
    )
    treemap_data["Eficiencia"] = treemap_data["Ejecucion"] / treemap_data["Presupuesto"] * 100
    fig_treemap = px.treemap(
        treemap_data,
        path=[px.Constant("ODS / EJE END"), "NOMBRE CORTO ODS", "EJE END"],
        values="Presupuesto",
        color="Eficiencia",
        color_continuous_scale="plasma",
        title="Distribución del Presupuesto por ODS y Eje END con Eficiencia",
        template="plotly_white",
    )
    st.plotly_chart(fig_treemap, use_container_width=True)

with tabs[3]:
    st.subheader("Plan de Acción Ciudadana")
    action_focus = st.selectbox(
        "Enfoque de acción",
        [
            "Transparencia y rendición de cuentas",
            "Vigilancia de servicios básicos",
            "Seguimiento de ODS prioritarios",
        ],
    )

    st.markdown(
        "Este panel está diseñado para que organizaciones civiles, comunidades y medios identifiquen"
        " proyectos y regiones donde la inversión pública necesita más vigilancia y explicación."
    )

    priority_projects = (
        filtered_df.assign(
            Brecha=filtered_df["Subejecución / Brecha"],
            Eficiencia=filtered_df["% Ejecución Financiera"],
        )
        .query("Eficiencia < 75")
        .sort_values(["Brecha", "Eficiencia"], ascending=[False, True])
        .head(10)
    )

    st.markdown("### Proyectos críticos para seguimiento ciudadano")
    st.dataframe(
        priority_projects[
            [
                "PERIODO",
                "SNIP",
                "NOMBRE PROYECTO",
                "INSTITUCION EJECUTORA",
                "REGION O PROVINCIA",
                "TOTAL PRESUPUESTO VIGENTE",
                "TOTAL EJECUCION",
                "% Ejecución Financiera",
                "Subejecución / Brecha",
            ]
        ],
        use_container_width=True,
        hide_index=True,
    )

    ods_priority = compute_ods_priorities(filtered_df)
    st.markdown("### ODS que requieren mayor atención pública")
    st.dataframe(
        ods_priority[
            ["NOMBRE CORTO ODS", "Presupuesto", "Ejecucion", "Eficiencia", "Brecha", "Urgencia"]
        ]
        .rename(
            columns={
                "NOMBRE CORTO ODS": "ODS",
                "Presupuesto": "Presupuesto RD$",
                "Ejecucion": "Ejecución RD$",
                "Eficiencia": "Eficiencia (%)",
                "Brecha": "Brecha RD$",
                "Urgencia": "Urgencia",
            }
        )
        .head(6),
        use_container_width=True,
        hide_index=True,
    )

    if action_focus == "Transparencia y rendición de cuentas":
        st.markdown(
            "#### Carta de petición ciudadana"
            "\n" "Use esta propuesta como base para solicitar información a instituciones y gobiernos locales."
        )
        message = (
            "Solicito información detallada sobre los proyectos con baja ejecución y alta brecha financiera, "
            "especialmente los listados en esta herramienta para la región seleccionada. "
            "Solicito el acceso a cronogramas, avances físicos, montos ejecutados, contrataciones y uso de recursos."
        )
    elif action_focus == "Vigilancia de servicios básicos":
        st.markdown(
            "#### Mensaje para comunidades"
            "\n" "Identifique los proyectos de su región relacionados con agua, salud o educación y comparta esta información "
            "para exigir informes periódicos sobre resultados."
        )
        message = (
            "Es urgente que las autoridades locales informen el estado real de los proyectos que afectan el servicio "
            "de agua potable, salud y movilidad en nuestra comunidad. Exigimos transparencia en ejecución y entregables."
        )
    else:
        st.markdown(
            "#### Monitoreo de ODS"
            "\n" "Use estos datos para vincular la inversión pública con los Objetivos de Desarrollo Sostenible y monitorear "
            "la entrega de valor social en su región."
        )
        message = (
            "Solicito que se priorice la ejecución de proyectos alineados con los ODS más urgentes, y que se publique "
            "un informe de resultados sobre el impacto social y ambiental esperado y real."
        )

    st.text_area("Texto de acción ciudadana", value=message, height=180)

    csv_buffer = BytesIO()
    explorer_df = filtered_df.copy()
    explorer_df.to_csv(csv_buffer, index=False)
    xlsx_buffer = BytesIO()
    explorer_df.to_excel(xlsx_buffer, index=False, engine="openpyxl")

    col_csv, col_excel = st.columns(2)
    with col_csv:
        st.download_button(
            label="Exportar datos completos",
            data=csv_buffer.getvalue(),
            file_name="mepyd_proyectos_accion_ciudadana.csv",
            mime="text/csv",
        )
    with col_excel:
        st.download_button(
            label="Exportar datos completos a Excel",
            data=xlsx_buffer.getvalue(),
            file_name="mepyd_proyectos_accion_ciudadana.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
