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
        background: linear-gradient(135deg, #f7fbff 0%, #eef4fb 100%);
    }
    .block-container {
        padding-top: 2rem;
        padding-bottom: 2rem;
    }
    .stTitle {
        color: #0f4c81;
        font-weight: 700;
    }
    .stCaption {
        color: #365f7a;
    }
    div[data-testid="stSidebar"] {
        background-color: #0f4c81;
        color: white;
    }
    div[data-testid="stSidebar"] .stTextInput > label,
    div[data-testid="stSidebar"] .stSelectbox > label,
    div[data-testid="stSidebar"] .stMultiSelect > label,
    div[data-testid="stSidebar"] .stButton > button {
        color: white;
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


def load_dataset(uploaded_file=None, force_refresh: bool = False) -> pd.DataFrame:
    if uploaded_file is not None:
        raw_df = load_data_from_upload(uploaded_file)
    else:
        raw_df = load_remote_data(force_refresh=force_refresh)

    return clean_dataset(raw_df)


st.markdown(
    """
    <div style="background-color:#0f4c81; padding: 1.2rem 1.4rem; border-radius: 14px; margin-bottom: 1rem;">
    <h1 style="color:white; margin:0; font-size:2rem;">Análisis de Ejecución de Proyectos de Inversión - MEPyD</h1>
    <p style="color:#dcecfb; margin:0.35rem 0 0 0;">Dataset oficial del Ministerio de Economía, Planificación y Desarrollo, periodo 2018-2024</p>
    </div>
    """,
    unsafe_allow_html=True,
)

with st.sidebar:
    st.header("Fuente de datos")
    st.write("Descarga automática desde la fuente oficial o sube un archivo manualmente.")
    uploaded_file = st.file_uploader("Subir archivo CSV/XLSX manualmente", type=["csv", "xlsx", "xls"])

    if st.button("Recargar / Refrescar Datos desde MEPyD", use_container_width=True):
        st.cache_data.clear()
        st.session_state["force_refresh"] = True
        st.rerun()

    st.markdown("---")
    st.caption("Fuentes oficiales:")
    st.caption(CSV_URL)

force_refresh = st.session_state.pop("force_refresh", False)

try:
    df = load_dataset(uploaded_file=uploaded_file, force_refresh=force_refresh)
except Exception as exc:
    st.error(f"No fue posible cargar el dataset: {exc}")
    st.stop()

if df.empty:
    st.warning("El dataset cargado está vacío.")
    st.stop()

# Filtros interactivos
with st.sidebar:
    st.header("Filtros")
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
tabs = st.tabs(["Dashboard General", "Alertas de Desempeño", "Inversión Territorial", "Explorador y Exportación"])

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

with tab2:
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

with tab3:
    st.markdown("### Inversión por región y por nivel estratégico")
    regional_investment = (
        filtered_df.groupby("REGION O PROVINCIA", as_index=False)
        .agg(
            Presupuesto=("TOTAL PRESUPUESTO VIGENTE", "sum"),
            Ejecucion=("TOTAL EJECUCION", "sum"),
            Eficiencia=("% Ejecución Financiera", "mean"),
        )
        .sort_values("Presupuesto", ascending=False)
    )
    fig_region = px.bar(
        regional_investment,
        x="Presupuesto",
        y="REGION O PROVINCIA",
        orientation="h",
        color="Eficiencia",
        title="Inversión y eficiencia por Región/Provincia",
        template="plotly_white",
        color_continuous_scale="viridis",
    )
    st.plotly_chart(fig_region, use_container_width=True)

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
    st.subheader("Explorador abierto")
    search_text = st.text_input("Buscar por nombre de proyecto o SNIP")
    explorer_df = filtered_df.copy()

    if search_text:
        explorer_df = explorer_df[
            explorer_df["NOMBRE PROYECTO"].astype(str).str.contains(search_text, case=False, na=False)
            | explorer_df["SNIP"].astype(str).str.contains(search_text, case=False, na=False)
        ]

    explorer_df = explorer_df[
        [
            "PERIODO",
            "SNIP",
            "NOMBRE PROYECTO",
            "INSTITUCION EJECUTORA",
            "REGION O PROVINCIA",
            "NOMBRE CORTO ODS",
            "EJE END",
            "TOTAL PRESUPUESTO VIGENTE",
            "TOTAL EJECUCION",
            "% Ejecución Financiera",
            "PORCENTAJE",
            "Subejecución / Brecha",
        ]
    ]

    st.dataframe(explorer_df, use_container_width=True, hide_index=True)

    csv_buffer = BytesIO()
    explorer_df.to_csv(csv_buffer, index=False)
    xlsx_buffer = BytesIO()
    explorer_df.to_excel(xlsx_buffer, index=False, engine="openpyxl")

    col_csv, col_excel = st.columns(2)
    with col_csv:
        st.download_button(
            label="Exportar a CSV",
            data=csv_buffer.getvalue(),
            file_name="mepyd_proyectos_filtrados.csv",
            mime="text/csv",
        )
    with col_excel:
        st.download_button(
            label="Exportar a Excel",
            data=xlsx_buffer.getvalue(),
            file_name="mepyd_proyectos_filtrados.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
