"""
Aplicación Streamlit para analizar el dataset oficial del MEPyD.

Instalación de dependencias:
pip install streamlit pandas plotly requests openpyxl
"""

from io import BytesIO, StringIO
import re

import pandas as pd
import plotly.express as px
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

st.subheader("Resumen ejecutivo")
summary_col1, summary_col2, summary_col3, summary_col4 = st.columns(4)

presupuesto_total = float(filtered_df["TOTAL PRESUPUESTO VIGENTE"].sum())
ejecucion_total = float(filtered_df["TOTAL EJECUCION"].sum())
porcentaje_global = (ejecucion_total / presupuesto_total * 100) if presupuesto_total else 0.0
cantidad_proyectos = int(filtered_df.shape[0])

summary_col1.metric("Monto Total Presupuestado", f"RD$ {presupuesto_total:,.0f}")
summary_col2.metric("Monto Total Ejecutado", f"RD$ {ejecucion_total:,.0f}")
summary_col3.metric("% Global de Ejecución", f"{porcentaje_global:,.2f}%")
summary_col4.metric("Cantidad de Proyectos", f"{cantidad_proyectos:,}")

st.markdown("---")

with st.expander("Insights inteligentes", expanded=True):
    top_gap = filtered_df.sort_values("Subejecución / Brecha", ascending=False).head(3)
    if not top_gap.empty:
        st.write("- Proyectos con mayor brecha presupuestaria: ")
        for _, row in top_gap[["NOMBRE PROYECTO", "Subejecución / Brecha", "% Ejecución Financiera"]].iterrows():
            st.write(f"  - {row['NOMBRE PROYECTO']}: brecha RD$ {row['Subejecución / Brecha']:,.0f} y ejecución de {row['% Ejecución Financiera']:.1f}%")
    else:
        st.write("- No hay datos suficientes para construir insights adicionales.")

    worst_projects = filtered_df.sort_values("% Ejecución Financiera", ascending=True).head(3)
    if not worst_projects.empty:
        st.write("- Proyectos con menor ejecución financiera: ")
        for _, row in worst_projects[["NOMBRE PROYECTO", "% Ejecución Financiera"]].iterrows():
            st.write(f"  - {row['NOMBRE PROYECTO']}: {row['% Ejecución Financiera']:.1f}%")

tab1, tab2, tab3, tab4 = st.tabs(["Dashboard General", "Alertas de Desempeño", "Inversión Territorial", "Explorador y Exportación"])

with tab1:
    by_institution = (
        filtered_df.groupby("INSTITUCION EJECUTORA", as_index=False)
        .agg(Presupuesto=("TOTAL PRESUPUESTO VIGENTE", "sum"), Ejecucion=("TOTAL EJECUCION", "sum"))
    )
    fig_bar = px.bar(
        by_institution,
        x="INSTITUCION EJECUTORA",
        y=["Presupuesto", "Ejecucion"],
        barmode="group",
        title="Presupuesto Vigente vs Ejecución por Institución Ejecutora",
        template="plotly_white",
    )
    fig_bar.update_layout(xaxis_tickangle=-35)
    st.plotly_chart(fig_bar, use_container_width=True)

with tab2:
    scatter = px.scatter(
        filtered_df,
        x="PORCENTAJE",
        y="% Ejecución Financiera",
        color="PERIODO",
        hover_data=["NOMBRE PROYECTO", "INSTITUCION EJECUTORA", "REGION O PROVINCIA"],
        title="Relación entre Avance Físico y Ejecución Financiera",
        template="plotly_white",
    )
    st.plotly_chart(scatter, use_container_width=True)

    st.subheader("Top 15 proyectos con mayor brecha presupuestaria")
    top_gap_table = (
        filtered_df.assign(Brecha=filtered_df["Subejecución / Brecha"])
        .sort_values("Brecha", ascending=False)
        .head(15)[
            [
                "NOMBRE PROYECTO",
                "SNIP",
                "INSTITUCION EJECUTORA",
                "REGION O PROVINCIA",
                "TOTAL PRESUPUESTO VIGENTE",
                "TOTAL EJECUCION",
                "Subejecución / Brecha",
                "% Ejecución Financiera",
            ]
        ]
    )
    st.dataframe(top_gap_table, use_container_width=True, hide_index=True)

with tab3:
    regional_investment = (
        filtered_df.groupby("REGION O PROVINCIA", as_index=False)
        .agg(Inversión=("TOTAL PRESUPUESTO VIGENTE", "sum"))
        .sort_values("Inversión", ascending=False)
    )
    fig_region = px.bar(
        regional_investment,
        x="Inversión",
        y="REGION O PROVINCIA",
        orientation="h",
        color="REGION O PROVINCIA",
        title="Inversión Total por Región o Provincia",
        template="plotly_white",
    )
    fig_region.update_layout(showlegend=False)
    st.plotly_chart(fig_region, use_container_width=True)

    treemap_data = (
        filtered_df.groupby(["NOMBRE CORTO ODS", "EJE END"], as_index=False)
        .agg(Presupuesto=("TOTAL PRESUPUESTO VIGENTE", "sum"))
    )
    fig_treemap = px.treemap(
        treemap_data,
        path=[px.Constant("ODS / EJE END"), "NOMBRE CORTO ODS", "EJE END"],
        values="Presupuesto",
        title="Distribución del Presupuesto por ODS y Eje END",
        template="plotly_white",
    )
    st.plotly_chart(fig_treemap, use_container_width=True)

with tab4:
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
            "NOMBRE PROYECTO",
            "SNIP",
            "INSTITUCION EJECUTORA",
            "REGION O PROVINCIA",
            "NOMBRE CORTO ODS",
            "TOTAL PRESUPUESTO VIGENTE",
            "TOTAL EJECUCION",
            "% Ejecución Financiera",
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
