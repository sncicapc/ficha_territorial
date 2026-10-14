# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
from html import escape as escapar_html
import re
import altair as alt
import plotly.express as px
import json
from io import BytesIO
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable
from reportlab.lib.enums import TA_CENTER

from logica import (
    SIN_DATO, a_fecha, activas_a_fecha, agregar, aplicar_planes_sncic, asignar_sector, buscar_col,
    cargar_config, cargar_sectores, cargar_xlsm, contar_movimiento, fecha_larga, format_cop, format_usd,
    get_col, leer_planes, n_departamentos, n_municipios, n_unicos, norm_text, proyectos_contrapartida,
    registro_publico, titulo_es, top_by_sum, top_comparativo, unificar_colcol, unir, validar_columnas,
    valor_corto,
)

st.set_page_config(
    page_title="Ficha territorial y sectorial de Cooperaci\u00f3n Internacional | APC Colombia",
    layout="wide",
    initial_sidebar_state="collapsed"
)
st.markdown("""
    <style>
    :root { color-scheme: light; }
    </style>
""", unsafe_allow_html=True)

FILE = "Ficha_territorial.xlsm"
FILE_SECTORES = "Ficha sectores.xlsx"
GEO_FILE = "Colombia.geo.json"
LOGO_APC = "logo_apc.png"
LOGO_SNCIC = "logo_sncic.png"

# Cortes, fechas de fuentes, sectores visibles y equivalencias: se editan en este archivo, no en el código
CONFIG_FILE = "configuracion.xlsx"
# Planes de trabajo territoriales del SNCIC (se genera con scripts/actualizar_planes_sncic.py)
PLANES_FILE = "planes_trabajo_sncic.xlsx"
# Columnas auxiliares que no deben salir en las descargas
COLS_INTERNAS = ["DEPT_NORM", "SECTOR_FICHA"]

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@400;500;600;700;800&family=Source+Sans+3:wght@400;600&display=swap');

:root {
    --apc-blue:      #003087;
    --apc-blue-mid:  #004BB4;
    --apc-red:       #C8102E;
    --apc-yellow:    #F5A623;
    --apc-light:     #EEF3FB;
    --apc-gray:      #F7F8FA;
    --apc-gray-mid:  #EAECF0;
    --apc-border:    #D1D9E6;
    --apc-text:      #1C2B4A;
    --apc-muted:     #5A6A85;
    --apc-white:     #FFFFFF;
}

html, body, [class*="css"] {
    font-family: 'Source Sans 3', sans-serif;
    color: var(--apc-text);
    background-color: #F7F8FA;
}

/* \u2500\u2500 Header \u2500\u2500 */
.apc-header {
    background: #003087;
    padding: 1.2rem 2.2rem 1rem 2.2rem;
    margin-bottom: 0;
    display: flex;
    align-items: center;
    justify-content: space-between;
    border-bottom: 4px solid var(--apc-red);
}
.apc-header-title {
    color: white;
    font-family: 'Montserrat', sans-serif;
    font-size: 1.5rem;
    font-weight: 700;
    letter-spacing: 0.2px;
    margin: 0;
}
.apc-header-subtitle {
    color: rgba(255,255,255,0.75);
    font-size: 0.82rem;
    font-weight: 400;
    margin-top: 3px;
    font-family: 'Source Sans 3', sans-serif;
    letter-spacing: 0.3px;
}
.apc-flag-bar {
    height: 5px;
    background: linear-gradient(90deg, #F5A623 33.3%, #003087 33.3% 66.6%, #C8102E 66.6%);
    margin-bottom: 1.5rem;
}

/* \u2500\u2500 Selector card \u2500\u2500 */
.dept-selector-card {
    background: var(--apc-white);
    border-left: 5px solid var(--apc-blue);
    border-radius: 0 8px 8px 0;
    padding: 0.9rem 1.4rem;
    margin-bottom: 1.4rem;
    box-shadow: 0 1px 4px rgba(0,48,135,0.07);
}

/* \u2500\u2500 Banner departamento \u2500\u2500 */
.dept-title-banner {
    background: var(--apc-blue);
    color: white;
    font-family: 'Montserrat', sans-serif;
    font-size: 1.2rem;
    font-weight: 700;
    padding: 0.75rem 1.5rem;
    border-radius: 6px;
    margin-bottom: 1.2rem;
    border-left: 6px solid var(--apc-red);
    letter-spacing: 0.2px;
}

/* \u2500\u2500 Section headers \u2500\u2500 */
.section-header {
    font-family: 'Montserrat', sans-serif;
    font-weight: 700;
    font-size: 0.95rem;
    color: var(--apc-blue);
    text-transform: uppercase;
    letter-spacing: 1px;
    border-bottom: 3px solid var(--apc-yellow);
    padding-bottom: 6px;
    margin: 2rem 0 1rem 0;
}

/* \u2500\u2500 M\u00e9tricas \u2500\u2500 */
div[data-testid="stMetric"] {
    background: var(--apc-white);
    border: 1px solid var(--apc-border);
    border-left: 5px solid var(--apc-blue);
    border-radius: 6px;
    padding: 1rem 1.1rem !important;
    box-shadow: 0 1px 6px rgba(0,48,135,0.06);
    transition: box-shadow 0.2s, border-left-color 0.2s;
}
div[data-testid="stMetric"]:hover {
    box-shadow: 0 4px 14px rgba(0,48,135,0.12);
    border-left-color: var(--apc-red);
}
div[data-testid="stMetricLabel"] {
    font-family: 'Montserrat', sans-serif;
    font-weight: 600;
    font-size: 0.72rem;
    color: var(--apc-muted) !important;
    text-transform: uppercase;
    letter-spacing: 0.7px;
}
div[data-testid="stMetricValue"] {
    font-family: 'Montserrat', sans-serif;
    font-weight: 700;
    font-size: 1.6rem !important;
    color: var(--apc-blue) !important;
}

/* \u2500\u2500 Tabs \u2500\u2500 */
button[data-baseweb="tab"] {
    font-family: 'Montserrat', sans-serif !important;
    font-weight: 600 !important;
    font-size: 0.85rem !important;
    letter-spacing: 0.4px;
    text-transform: uppercase;
    color: var(--apc-muted) !important;
}
button[data-baseweb="tab"][aria-selected="true"] {
    color: var(--apc-blue) !important;
}
div[data-baseweb="tab-highlight"] {
    background-color: var(--apc-red) !important;
    height: 3px !important;
}
div[data-baseweb="tab-border"] {
    background-color: var(--apc-border) !important;
}

/* \u2500\u2500 Dataframes \u2500\u2500 */
div[data-testid="stDataFrame"] {
    border-radius: 6px;
    overflow: hidden;
    border: 1px solid var(--apc-border);
    box-shadow: 0 1px 4px rgba(0,0,0,0.04);
}

/* \u2500\u2500 Botones descarga \u2500\u2500 */
div[data-testid="stDownloadButton"] button {
    background: var(--apc-blue) !important;
    color: white !important;
    border: none !important;
    border-radius: 4px !important;
    font-family: 'Montserrat', sans-serif !important;
    font-weight: 600 !important;
    font-size: 0.82rem !important;
    letter-spacing: 0.5px;
    text-transform: uppercase;
    padding: 0.5rem 1.4rem !important;
    transition: background 0.2s !important;
}
div[data-testid="stDownloadButton"] button:hover {
    background: var(--apc-red) !important;
}

/* \u2500\u2500 Guia card \u2500\u2500 */
.guia-card {
    background: var(--apc-white);
    border: 1px solid var(--apc-border);
    border-radius: 8px;
    padding: 2rem 2.4rem;
    margin-bottom: 1rem;
    line-height: 1.8;
    font-size: 0.97rem;
    box-shadow: 0 1px 6px rgba(0,48,135,0.05);
}
.guia-card p {
    color: var(--apc-text);
    margin-bottom: 1.1rem;
}
.guia-intro {
    font-family: 'Montserrat', sans-serif;
    font-size: 0.95rem;
    font-weight: 700;
    color: var(--apc-blue);
    background: var(--apc-light);
    border-left: 5px solid var(--apc-red);
    border-radius: 0 6px 6px 0;
    padding: 0.8rem 1.2rem;
    margin-bottom: 1.4rem;
    text-transform: uppercase;
    letter-spacing: 0.5px;
}

/* \u2500\u2500 Footer \u2500\u2500 */
.apc-footer {
    text-align: center;
    color: var(--apc-muted);
    font-size: 0.76rem;
    margin-top: 3rem;
    padding-top: 1rem;
    border-top: 1px solid var(--apc-border);
    letter-spacing: 0.3px;
}

div[data-testid="stMetricLabel"] p {
    font-size: 0.65rem !important;
    line-height: 1.3 !important;
    white-space: normal !important;
}

/* Metrica personalizada con texto largo */
.metric-custom {
    background: var(--apc-white);
    border: 1px solid var(--apc-border);
    border-left: 5px solid var(--apc-blue);
    border-radius: 6px;
    padding: 1rem 1.1rem;
    box-shadow: 0 1px 6px rgba(0,48,135,0.06);
    transition: box-shadow 0.2s, border-left-color 0.2s;
    height: 100%;
}
.metric-custom:hover {
    box-shadow: 0 4px 14px rgba(0,48,135,0.12);
    border-left-color: #C8102E;
}
.metric-custom-label {
    font-family: 'Montserrat', sans-serif;
    font-weight: 600;
    font-size: 0.68rem;
    color: #5A6A85;
    text-transform: uppercase;
    letter-spacing: 0.5px;
    line-height: 1.35;
    margin-bottom: 0.4rem;
}
.metric-custom-value {
    font-family: 'Montserrat', sans-serif;
    font-weight: 700;
    font-size: 1.6rem;
    color: #003087;
    line-height: 1.2;
}
.metric-custom-delta {
    font-family: 'Montserrat', sans-serif;
    font-size: 0.72rem;
    font-weight: 600;
    margin-top: 4px;
}
.metric-custom-sub {
    font-family: 'Source Sans 3', sans-serif;
    font-size: 0.72rem;
    color: #5A6A85;
    margin-top: 2px;
    line-height: 1.4;
}

/* Ícono de ayuda (ⓘ) en las tarjetas */
.ayuda {
    cursor: help;
    color: #004BB4;
    font-size: 0.85rem;
    text-transform: none;
    margin-left: 3px;
}

/* Ocultar barra Streamlit */
div[data-testid="stToolbar"],
div[data-testid="stDecoration"],
header[data-testid="stHeader"] {
    display: none !important;
}

/* \u2500\u2500 Logo header \u2500\u2500 */
.app-header {
    background: #003087;
    padding: 0.7rem 2rem;
    display: flex;
    align-items: center;
    justify-content: space-between;
    border-bottom: 4px solid #C8102E;
    margin-bottom: 0;
}
.app-header-center {
    text-align: center;
    flex: 1;
    padding: 0 1rem;
}
.app-header-title {
    color: white;
    font-family: 'Montserrat', sans-serif;
    font-size: 1rem;
    font-weight: 700;
    margin: 0;
    letter-spacing: 0.3px;
}
.app-header-sub {
    color: rgba(255,255,255,0.72);
    font-size: 0.75rem;
    margin-top: 2px;
    font-family: 'Source Sans 3', sans-serif;
}
.logo-side {
    width: 120px;
    display: flex;
    align-items: center;
    justify-content: center;
}

/* \u2500\u2500 Panel general cards \u2500\u2500 */
.panel-stat {
    background: white;
    border: 1px solid #D1D9E6;
    border-left: 5px solid #003087;
    border-radius: 6px;
    padding: 1rem 1.2rem;
    margin-bottom: 0.5rem;
    box-shadow: 0 1px 4px rgba(0,48,135,0.06);
}
.panel-stat-label {
    font-family: 'Montserrat', sans-serif;
    font-size: 0.68rem;
    font-weight: 600;
    color: #5A6A85;
    text-transform: uppercase;
    letter-spacing: 0.6px;
    margin-bottom: 4px;
}
.panel-stat-value {
    font-family: 'Montserrat', sans-serif;
    font-size: 1.6rem;
    font-weight: 700;
    color: #003087;
    line-height: 1.1;
}
.panel-stat-sub {
    font-size: 0.72rem;
    color: #5A6A85;
    margin-top: 3px;
}
.panel-stat-red { border-left-color: #C8102E; }
.panel-stat-yellow { border-left-color: #F5A623; }

/* \u2500\u2500 Section divider with icon \u2500\u2500 */
.section-title {
    font-family: 'Montserrat', sans-serif;
    font-weight: 700;
    font-size: 0.9rem;
    color: #003087;
    text-transform: uppercase;
    letter-spacing: 1px;
    border-bottom: 3px solid #F5A623;
    padding-bottom: 6px;
    margin: 1.8rem 0 1rem 0;
}

/* \u2500\u2500 Selector pill \u2500\u2500 */
div[data-testid="stSelectbox"] > div > div {
    border-radius: 6px !important;
    border-color: #D1D9E6 !important;
}

/* \u2500\u2500 Map container \u2500\u2500 */
.map-title {
    font-family: 'Montserrat', sans-serif;
    font-weight: 600;
    font-size: 0.8rem;
    color: #5A6A85;
    text-transform: uppercase;
    letter-spacing: 0.5px;
    margin-bottom: 6px;
}

/* Fondo blanco forzado */
.stApp {
    background-color: #FFFFFF !important;
}
[data-testid="stAppViewContainer"] {
    background-color: #FFFFFF !important;
}
[data-testid="stMain"] {
    background-color: #FFFFFF !important;
}

/* Forzar modo claro en Mac */
:root {
    color-scheme: light only;
}
body {
    background-color: #FFFFFF !important;
}

/* Textos visibles */
[data-testid="stMarkdownContainer"] p,
[data-testid="stMarkdownContainer"] li,
[data-testid="stCaptionContainer"] p,
[data-testid="stRadio"] label span,
[data-testid="stMetricLabel"] p {
    color: #1C2B4A !important;
}

/* Fondo blanco en gr\u00e1ficas - Mac */
[data-testid="stArrowVegaLiteChart"] > div,
[data-testid="stDataFrame"] > div,
.js-plotly-plot,
.plotly {
    background-color: #FFFFFF !important;
}

/* Texto botones */
div[data-testid="stDownloadButton"] button,
div[data-testid="stDownloadButton"] button p,
div[data-testid="stDownloadButton"] button span {
    color: white !important;
}


</style>
""", unsafe_allow_html=True)

@st.cache_data(show_spinner="Cargando datos de cooperación…")
def load_config():
    return cargar_config(CONFIG_FILE)


@st.cache_data(show_spinner="Cargando datos de cooperación…")
def load_data():
    return cargar_xlsm(FILE)


@st.cache_data(show_spinner="Cargando datos de cooperación…")
def load_sectores(hoja_act, hoja_ant):
    return cargar_sectores(FILE_SECTORES, hoja_act, hoja_ant)


@st.cache_data(show_spinner="Cargando datos de cooperación…")
def load_geo():
    with open(GEO_FILE, encoding="utf-8") as f:
        return json.load(f)


# Nombres del GeoJSON que se escriben distinto en las fuentes
GEO_ALIAS = {"SANTAFE DE BOGOTA D C": "BOGOTA D C"}


def geo_norm(nombre):
    n = norm_text(nombre)
    return GEO_ALIAS.get(n, n)


def make_map(geo, dept_values, etiqueta="Intervenciones", formato=":,.0f", selected_dept=None, alto=520, escala=True):
    """Mapa coroplético de Colombia; dept_values usa como llave el nombre normalizado del departamento."""
    names = [f["properties"]["NOMBRE_DPT"] for f in geo["features"]]
    df_map = pd.DataFrame({"dept_geo": names})
    df_map["value"] = df_map["dept_geo"].map(lambda x: dept_values.get(geo_norm(x), 0))

    fig = px.choropleth(
        df_map,
        geojson=geo,
        locations="dept_geo",
        featureidkey="properties.NOMBRE_DPT",
        color="value",
        color_continuous_scale=[
            [0, "#EEF3FB"],
            [0.2, "#A8C4E8"],
            [0.5, "#4A90D9"],
            [0.8, "#1565C0"],
            [1.0, "#003087"]
        ],
        hover_name="dept_geo",
        hover_data={"value": formato, "dept_geo": False},
        labels={"value": etiqueta},
    )
    # Encuadre fijo de Colombia (incluye San Andrés); fitbounds no ajusta bien el zoom en Plotly reciente
    fig.update_geos(
        visible=False, bgcolor="rgba(0,0,0,0)", projection_type="mercator",
        lonaxis_range=[-82.5, -66.5], lataxis_range=[-4.6, 13.8],
    )
    fig.update_layout(
        margin={"r": 0, "t": 0, "l": 0, "b": 0},
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        height=alto,
        coloraxis=dict(colorbar=dict(
            title=dict(text=etiqueta, font=dict(size=9)), thickness=12, len=0.6, tickfont=dict(size=9)
        )),
    )

    if not escala:
        fig.update_layout(coloraxis_showscale=False)
    if selected_dept:
        selected_features = [
            f for f in geo["features"] if geo_norm(f["properties"]["NOMBRE_DPT"]) == norm_text(selected_dept)
        ]
        if selected_features:
            import plotly.graph_objects as go
            fig.add_trace(go.Choropleth(
                geojson={"type": "FeatureCollection", "features": selected_features},
                locations=[selected_features[0]["properties"]["NOMBRE_DPT"]],
                z=[1],
                featureidkey="properties.NOMBRE_DPT",
                colorscale=[[0, "#C8102E"], [1, "#C8102E"]],
                showscale=False,
                marker_line_color="white",
                marker_line_width=2,
                hoverinfo="skip"
            ))
    return fig


def movimiento_intervenciones(df_act, df_ant):
    """Texto de la tarjeta: nuevas, continúan, terminadas y sin reporte."""
    m = contar_movimiento(df_act, df_ant, CORTE_ACT["vigencia"])
    if m is None:
        return ""
    texto = (f"♥ {m['nuevas']} nuevas &nbsp;|&nbsp; ↺ {m['continuan']} continúan"
             f" &nbsp;|&nbsp; ✓ {m['terminadas']} terminadas")
    if m["sin_reporte"]:
        texto += f" &nbsp;|&nbsp; ? {m['sin_reporte']} sin reporte"
    return texto


ODS_NOMBRES = {
    "ODS 1":  "ODS 1 - Fin de la pobreza",
    "ODS 2":  "ODS 2 - Hambre cero",
    "ODS 3":  "ODS 3 - Salud y bienestar",
    "ODS 4":  "ODS 4 - Educaci\u00f3n de calidad",
    "ODS 5":  "ODS 5 - Igualdad de g\u00e9nero",
    "ODS 6":  "ODS 6 - Agua limpia y saneamiento",
    "ODS 7":  "ODS 7 - Energ\u00eda asequible y no contaminante",
    "ODS 8":  "ODS 8 - Trabajo decente y crecimiento econ\u00f3mico",
    "ODS 9":  "ODS 9 - Industria, innovaci\u00f3n e infraestructura",
    "ODS 10": "ODS 10 - Reducci\u00f3n de las desigualdades",
    "ODS 11": "ODS 11 - Ciudades y comunidades sostenibles",
    "ODS 12": "ODS 12 - Producci\u00f3n y consumo responsables",
    "ODS 13": "ODS 13 - Acci\u00f3n por el clima",
    "ODS 14": "ODS 14 - Vida submarina",
    "ODS 15": "ODS 15 - Vida de ecosistemas terrestres",
    "ODS 16": "ODS 16 - Paz, justicia e instituciones s\u00f3lidas",
    "ODS 17": "ODS 17 - Alianzas para lograr los objetivos",
}



LIMITE_FILAS_PDF = 30


def tabla_top_comparativa(top_act, df_ant, group_col, nombres=None):
    return top_comparativo(top_act, df_ant, group_col, COL_USD_ACT, COL_USD_ANT, nombres=nombres)


def nota_truncado(total, mostrados):
    """Aviso en el PDF cuando una tabla no cabe completa."""
    estilo = ParagraphStyle("nota", fontName="Helvetica-Oblique", fontSize=7,
                            textColor=colors.HexColor("#C8102E"), spaceBefore=3, spaceAfter=4)
    return Paragraph(f"Se muestran {mostrados} de {total} registros. "
                     "Consulte la descarga en Excel para el listado completo.", estilo)


@st.cache_data(show_spinner=False)
def to_excel_ficha(info_row, cic_dept, colcol_dept, contr_dept, css_dept=None, proj_activos=None):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        if not info_row.empty:
            pd.DataFrame(registro_publico(info_row.iloc[0]), columns=["Campo", "Valor"]).to_excel(
                writer, sheet_name="Informacion General", index=False)
        cic_export = cic_dept.drop(columns=COLS_INTERNAS, errors="ignore")
        cic_export.to_excel(writer, sheet_name="AOD - Ciclope", index=False)
        if css_dept is not None and not css_dept.empty:
            COLS_CSS = [
                "C\u00f3digo", "VIA DE COOPERACION", "MODALIDAD", "PAIS SOCIO", "SEGUNDO OFERENTE",
                "REGION", "NOMBRE DE LA INICIATIVA", "TIPO DE INICIATIVA", "FECHA DE APROBACION",
                "OBJETIVO GENERAL/DESCRIPCION DE LA INICIATIVA", "ESTADO",
                "ENTIDAD(ES) NACIONAL(ES)", "ENTIDAD(ES) EXTRANJERA(S)"
            ]
            cols_css = [c for c in COLS_CSS if c in css_dept.columns]
            css_dept[cols_css].to_excel(writer, sheet_name="Coop Sur Sur", index=False)
        colcol_dept.drop(columns=COLS_INTERNAS, errors="ignore").to_excel(writer, sheet_name="ColCol", index=False)
        contr_dept.to_excel(writer, sheet_name="Contrapartidas", index=False)
        if proj_activos is not None and not proj_activos.empty:
            proj_activos.to_excel(writer, sheet_name="Proyectos AOD activos", index=False)
    output.seek(0)
    return output.getvalue()


def _pdf_ficha(tipo, nombre, metricas_info, detalle_info, aod_act, aod_ant, tercera,
               css_df, colcol_df, contr_df=None):
    """PDF común para la ficha territorial y la sectorial.

    tipo: "Territorial" o "Sectorial". metricas_info: [(etiqueta, valor)] para la fila de tarjetas.
    detalle_info: [(campo, valor)] para la tabla de información detallada. tercera: "municipios" o "departamentos".
    contr_df: contrapartidas (solo aplica a la ficha territorial; None la omite).
    """
    output = BytesIO()
    doc = SimpleDocTemplate(output, pagesize=A4, leftMargin=2 * cm, rightMargin=2 * cm,
                            topMargin=2 * cm, bottomMargin=2 * cm,
                            title=f"Ficha {tipo} - {nombre}", author="APC-Colombia")

    AZUL = colors.HexColor("#003087")
    AZUL_CLARO = colors.HexColor("#E8F0FE")
    GRIS = colors.HexColor("#F5F7FA")
    GRIS_BORDE = colors.HexColor("#D0D9EA")
    BLANCO = colors.white
    TEXTO = colors.HexColor("#1A1A2E")
    MUTED = colors.HexColor("#6B7280")

    e_titulo = ParagraphStyle("titulo", fontName="Helvetica-Bold", fontSize=20, textColor=BLANCO, leading=24, spaceAfter=4)
    e_subtitulo = ParagraphStyle("subtitulo", fontName="Helvetica", fontSize=9, textColor=colors.HexColor("#CBD5E1"))
    e_badge = ParagraphStyle("badge", fontName="Helvetica-Bold", fontSize=9, textColor=BLANCO, alignment=TA_CENTER)
    e_banner = ParagraphStyle("banner", fontName="Helvetica-Bold", fontSize=15, textColor=BLANCO, leading=20)
    e_seccion = ParagraphStyle("seccion", fontName="Helvetica-Bold", fontSize=10, textColor=AZUL, spaceBefore=12, spaceAfter=6)
    e_subseccion = ParagraphStyle("subseccion", fontName="Helvetica-Bold", fontSize=8.5, textColor=AZUL, spaceBefore=10, spaceAfter=5)
    e_normal = ParagraphStyle("normal", fontName="Helvetica", fontSize=7.5, textColor=TEXTO, leading=10)
    e_label = ParagraphStyle("label", fontName="Helvetica-Bold", fontSize=7, textColor=MUTED, leading=9)
    e_valor = ParagraphStyle("valor", fontName="Helvetica-Bold", fontSize=13, textColor=AZUL, leading=16)
    e_valor_chico = ParagraphStyle("valor_chico", fontName="Helvetica-Bold", fontSize=9.5, textColor=AZUL, leading=12)
    e_caption = ParagraphStyle("caption", fontName="Helvetica", fontSize=7, textColor=MUTED, spaceAfter=4, alignment=TA_CENTER)
    e_vacio = ParagraphStyle("vacio", fontName="Helvetica-Oblique", fontSize=7.5, textColor=MUTED, spaceAfter=4)

    def esc(v):
        """Texto seguro para Paragraph (escapa &, < y >)."""
        return str(v).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

    def seccion(titulo):
        story.append(Spacer(1, 6))
        story.append(HRFlowable(width="100%", thickness=2, color=AZUL, spaceAfter=4))
        story.append(Paragraph(titulo, e_seccion))

    def fila_metricas(pares):
        """Tarjetas en una fila: etiqueta arriba y valor abajo."""
        if not pares:
            return
        ancho = f"{100 / len(pares):.2f}%"
        valores = [Paragraph(esc(v), e_valor if len(str(v)) <= 12 else e_valor_chico) for _, v in pares]
        t = Table([[Paragraph(esc(k).upper(), e_label) for k, _ in pares], valores], colWidths=[ancho] * len(pares))
        t.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), GRIS), ("LINEABOVE", (0, 0), (-1, 0), 3, AZUL),
            ("GRID", (0, 0), (-1, -1), 0.5, GRIS_BORDE), ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("TOPPADDING", (0, 0), (-1, -1), 6), ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
            ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ]))
        story.append(t)
        story.append(Spacer(1, 6))

    def tabla(df, anchos, total=None, max_chars=90):
        """Tabla con encabezado azul claro; corta a LIMITE_FILAS_PDF filas y lo avisa."""
        mostrar = df.head(LIMITE_FILAS_PDF)
        datos = [[Paragraph(esc(c), e_label) for c in mostrar.columns]]
        for _, r in mostrar.iterrows():
            datos.append([Paragraph(esc("" if pd.isna(v) else str(v)[:max_chars]), e_normal) for v in r.values])
        t = Table(datos, colWidths=anchos, repeatRows=1)
        t.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), AZUL_CLARO), ("ROWBACKGROUNDS", (0, 1), (-1, -1), [BLANCO, GRIS]),
            ("GRID", (0, 0), (-1, -1), 0.5, GRIS_BORDE), ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ]))
        story.append(t)
        total = len(df) if total is None else total
        if total > LIMITE_FILAS_PDF:
            story.append(nota_truncado(total, LIMITE_FILAS_PDF))
        story.append(Spacer(1, 6))

    def vacio(texto):
        story.append(Paragraph(texto, e_vacio))

    story = []
    articulo = "el departamento" if tipo == "Territorial" else "el sector"

    # --- Encabezado ---
    enc = Table([[Paragraph(f"Ficha {tipo}", e_titulo), ""],
                 [Paragraph("Agencia Presidencial de Cooperación Internacional de Colombia", e_subtitulo),
                  Paragraph("APC-Colombia", e_badge)]], colWidths=["75%", "25%"])
    enc.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), AZUL), ("ROUNDEDCORNERS", [8, 8, 8, 8]),
        ("TOPPADDING", (0, 0), (-1, -1), 12), ("BOTTOMPADDING", (0, 0), (-1, -1), 12),
        ("LEFTPADDING", (0, 0), (0, -1), 16), ("RIGHTPADDING", (-1, 0), (-1, -1), 16),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]))
    story += [enc, Spacer(1, 10)]
    banner = Table([[Paragraph(esc(nombre), e_banner)]], colWidths=["100%"])
    banner.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), AZUL), ("ROUNDEDCORNERS", [8, 8, 8, 8]),
        ("TOPPADDING", (0, 0), (-1, -1), 10), ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
        ("LEFTPADDING", (0, 0), (-1, -1), 14),
    ]))
    story += [banner, Spacer(1, 4)]

    # --- Información general ---
    seccion("Información general")
    fila_metricas(metricas_info)
    detalle = [(k, v) for k, v in detalle_info if norm_text(v) not in SIN_DATO]
    if detalle:
        story.append(Paragraph(f"Información detallada {'del departamento' if tipo == 'Territorial' else 'del sector'}", e_subseccion))
        t = Table([[Paragraph(esc(k), e_label), Paragraph(esc(v), e_normal)] for k, v in detalle],
                  colWidths=["35%", "65%"])
        t.setStyle(TableStyle([
            ("ROWBACKGROUNDS", (0, 0), (-1, -1), [BLANCO, GRIS]), ("GRID", (0, 0), (-1, -1), 0.5, GRIS_BORDE),
            ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ("LEFTPADDING", (0, 0), (-1, -1), 8), ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ]))
        story.append(t)

    # --- Resumen de la cooperación ---
    seccion(f"La cooperación en {articulo}")
    resumen = [("Intervenciones AOD", n_unicos(aod_act, "CODIGO INTERVENCION")),
               ("Aporte AOD (USD)", valor_corto(aod_act["VALOR APORTE (USD)"].sum() if not aod_act.empty else 0)),
               ("Proyectos Sur-Sur", n_unicos(css_df, "Código")),
               ("Intercambios ColCol", n_unicos(colcol_df, "CODIGO"))]
    if contr_df is not None:
        resumen.append(("Contrapartidas", len(proyectos_contrapartida(contr_df))))
    fila_metricas(resumen)
    story.append(Paragraph(f"AOD: intervenciones vigentes al {VIGENCIA_TXT}. Sur-Sur: datos a {CSS_ACT}.", e_caption))

    # --- AOD ---
    seccion("Ayuda Oficial al Desarrollo (AOD)")
    story.append(Paragraph(FUENTE_AOD, e_caption))
    if aod_act.empty:
        vacio(f"No hay intervenciones de AOD vigentes al {VIGENCIA_TXT}.")
    else:
        tercera_par = (("Municipios / áreas", n_municipios(aod_act)) if tercera == "municipios"
                       else ("Departamentos", n_departamentos(aod_act)))
        fila_metricas([("Intervenciones", n_unicos(aod_act, "CODIGO INTERVENCION")),
                       ("Cooperantes", n_unicos(aod_act, "NOMBRE ACTOR")), tercera_par,
                       ("Aporte estimado", format_usd(aod_act["VALOR APORTE (USD)"].sum()))])
        anchos_top = ["46%", "18%", "18%", "18%"]
        story.append(Paragraph("Top 5 cooperantes por aporte estimado (USD)", e_subseccion))
        top = top_by_sum(aod_act, "NOMBRE ACTOR", "VALOR APORTE (USD)", 5)
        tabla(tabla_top_comparativa(top, aod_ant, "NOMBRE ACTOR").rename(columns={"NOMBRE ACTOR": "Cooperante"}), anchos_top)
        story.append(Paragraph("Top 5 ODS por aporte estimado (USD)", e_subseccion))
        top = top_by_sum(aod_act, "ODS", "VALOR APORTE (USD)", 5)
        tabla(tabla_top_comparativa(top, aod_ant, "ODS", nombres=ODS_NOMBRES), anchos_top)

    # --- Cooperación Sur-Sur ---
    seccion("Proyectos de Cooperación Sur-Sur aprobados y vigentes")
    story.append(Paragraph(FUENTE_CSS, e_caption))
    t_css = resumen_css(css_df)
    if t_css.empty:
        vacio(f"No se encontraron proyectos de Cooperación Sur-Sur para {articulo}.")
    else:
        tabla(t_css[["Código", "Iniciativa", "País socio", "Región", "Vía de cooperación", "Entidades nacionales"]],
              ["11%", "33%", "14%", "14%", "11%", "17%"])

    # --- Programas APC ---
    seccion("Programas de APC-Colombia")
    story.append(Paragraph("Colombia Enseña Colombia (ColCol)", e_subseccion))
    t_cc = resumen_colcol(colcol_df)
    if t_cc.empty:
        vacio(f"No se encontraron intercambios ColCol para {articulo}.")
    else:
        tabla(t_cc[["Código", "Intercambio", "Etapas", "Año", "Departamentos", "Presupuesto APC (COP)"]],
              ["9%", "35%", "13%", "7%", "20%", "16%"])
    if contr_df is not None:
        story.append(Paragraph("Contrapartidas", e_subseccion))
        if contr_df.empty:
            vacio(f"No se encontraron contrapartidas para {articulo}.")
        else:
            t = contr_df.drop(columns=COLS_INTERNAS + ["Departamento"], errors="ignore").copy()
            t.columns = [str(c).strip() for c in t.columns]
            for col in ("Monto total", "Monto por APC"):
                if col in t.columns:
                    t[col] = pd.to_numeric(t[col], errors="coerce").map(lambda v: format_cop(v) if pd.notna(v) else "")
            t = t.rename(columns={"AÑO": "Año", "Nombre del proyecto": "Proyecto",
                                  "Monto total": "Monto total (COP)", "Monto por APC": "Aporte APC (COP)"})
            orden = [c for c in ("Año", "Proyecto", "Implementador", "Financiador", "Monto total (COP)", "Aporte APC (COP)") if c in t.columns]
            anchos = {"Año": "7%", "Proyecto": "31%", "Implementador": "18%", "Financiador": "18%",
                      "Monto total (COP)": "13%", "Aporte APC (COP)": "13%"}
            tabla(t[orden], [anchos[c] for c in orden])

    # --- Pie ---
    story.append(Spacer(1, 12))
    story.append(HRFlowable(width="100%", thickness=0.5, color=GRIS_BORDE, spaceAfter=4))
    story.append(Paragraph(f"Agencia Presidencial de Cooperación Internacional de Colombia - APC-Colombia · "
                           f"Corte {ACT}: intervenciones AOD vigentes al {VIGENCIA_TXT}", e_caption))
    doc.build(story)
    return output.getvalue()


def _info_territorial(info_row):
    """Tarjetas y detalle de la información general del departamento."""
    if info_row.empty:
        return [], []
    r = info_row.iloc[0]
    pob = get_col(r, "Población", "Poblacion")
    try:
        pob = f"{int(float(pob)):,}".replace(",", ".")
    except (TypeError, ValueError):
        pass
    tiene_plan = str(get_col(r, "¿Cuenta con plan de trabajo en el marco del SNCIC?")).strip().lower() in ("sí", "si")
    anio = str(get_col(r, "Año creación del plan de trabajo"))
    plan = ("Sí" + (f" (desde {anio})" if anio.isdigit() else "")) if tiene_plan else "No"
    metricas = [("Capital", str(get_col(r, "Capital")).title()),
                ("Número de municipios", get_col(r, "Número de Municipios", "Numero de Municipios")),
                ("Población", pob), ("Plan de trabajo SNCIC", plan)]
    detalle = registro_publico(r)
    return metricas, detalle


def _info_sectorial(info_sector):
    """Tarjetas y detalle de la información general del sector."""
    if info_sector.empty:
        return [], []
    r = info_sector.iloc[0]
    con_plan = str(r.get("Cabezas del sector asociadas con planes de trabajo en el marco del SNCIC", "")).strip().lower() in ("sí", "si", "yes")
    metricas = [("Cabeza de sector", str(r.get("Cabeza de sector", "")).strip()),
                ("Plan de trabajo SNCIC", "Sí" if con_plan else "No")]
    detalle = [(str(k), titulo_es(v) if norm_text(k) == "PLANES DE TRABAJO VINCULADOS" else v) for k, v in r.items()]
    return metricas, detalle


@st.cache_data(show_spinner=False)
def to_pdf_ficha(dept, info_row, cic_dept, cic_dept_ant, colcol_dept, contr_dept, css_dept):
    metricas, detalle = _info_territorial(info_row)
    return _pdf_ficha("Territorial", dept, metricas, detalle, cic_dept, cic_dept_ant, "municipios",
                      css_dept, colcol_dept, contr_dept)


@st.cache_data(show_spinner=False)
def to_pdf_sectorial(sector, info_sector, aod_sector, aod_sector_ant, css_sector, colcol_sector):
    metricas, detalle = _info_sectorial(info_sector)
    return _pdf_ficha("Sectorial", sector, metricas, detalle, aod_sector, aod_sector_ant, "departamentos",
                      css_sector, colcol_sector)


# -------------------------
# APP
# -------------------------

# Load data
_cfg = load_config()
CORTES, FUENTES, SECTORES_OCULTOS = _cfg["cortes"], _cfg["fuentes"], _cfg["ocultos"]
EQUIVALENCIAS, REGISTROS_SECTOR = _cfg["equivalencias"], _cfg["registros"]
CORTE_ACT, CORTE_ANT = CORTES["actual"], CORTES["anterior"]
_x = load_data()
_s = load_sectores(CORTE_ACT["hoja"], CORTE_ANT["hoja"])
geo = load_geo()

# Si a una fuente le faltan columnas, se avisa con claridad en lugar de fallar más adelante
_problemas = validar_columnas({
    "infogeneral": ("Ficha_territorial.xlsm · tabla infogeneral", _x["infogeneral"]),
    "css": ("Ficha_territorial.xlsm · hoja CSS", _x["css"]),
    "colcol": ("Ficha_territorial.xlsm · tabla colcol", _x["colcol"]),
    "contrapartidas": ("Ficha_territorial.xlsm · tabla contrapartidas", _x["contrapartidas"]),
    "info_sectores": ("Ficha sectores.xlsx · hoja INFO GENERAL", _s["info"]),
    "aod": (f"Ficha sectores.xlsx · hoja {CORTE_ACT['hoja']} (AOD {CORTE_ACT['etiqueta']})", _s["aod"]),
    "aod_ant": (f"Ficha sectores.xlsx · hoja {CORTE_ANT['hoja']} (AOD {CORTE_ANT['etiqueta']})", _s["aod_ant"]),
    "colcol_sectores": ("Ficha sectores.xlsx · hoja COLCOL", _s["colcol"]),
})
if _problemas:
    st.error("No se pudo cargar la ficha porque los datos no tienen la estructura esperada:\n\n"
             + "\n".join(f"- {p}" for p in _problemas)
             + "\n\nRevise los encabezados de esas tablas (deben llamarse igual que en el corte anterior).")
    st.stop()

infogeneral = _x["infogeneral"]
colcol, contrapartidas, css = _x["colcol"], _x["contrapartidas"], _x["css"]
info_s = _s["info"]
# Sur-Sur y ColCol: una sola fuente para ambas fichas (el .xlsm)
css_s = css.copy()
colcol_s = unificar_colcol(colcol, _s["colcol"])

# Universo único de AOD: solo intervenciones vigentes en la fecha de cada corte
# AOD: una sola fuente para ambas fichas (Ficha sectores.xlsx, con los nombres individuales de los cooperantes)
ciclope = activas_a_fecha(_s["aod"], CORTE_ACT["vigencia"]).copy()
ciclope_ant = activas_a_fecha(_s["aod_ant"], CORTE_ANT["vigencia"]).copy()

# Sector de la ficha para cada registro, según configuracion.xlsx
sectores_visibles = [
    s for s in info_s["Nombre del sector"].dropna().tolist() if norm_text(s) not in SECTORES_OCULTOS
]
sectores_norm = {norm_text(s) for s in sectores_visibles}
_reglas = dict(sectores_norm=sectores_norm, equivalencias=EQUIVALENCIAS, registros=REGISTROS_SECTOR)
for _df in (ciclope, ciclope_ant):
    _df["SECTOR_FICHA"] = asignar_sector(_df, "SECTORES GOB", "AOD", codigo_col="CODIGO INTERVENCION", **_reglas)
css_s["SECTOR_FICHA"] = asignar_sector(css_s, "ESPACIO VINCULADO", "CSS", codigo_col="Código", **_reglas)
colcol_s["SECTOR_FICHA"] = asignar_sector(colcol_s, "SECTOR VINCULADO", "COLCOL", codigo_col="CODIGO", **_reglas)
aod_s, aod_s_ant = ciclope, ciclope_ant

# Textos de fechas y cortes para leyendas y comparativos
ACT, ANT = CORTE_ACT["etiqueta"], CORTE_ANT["etiqueta"]
VIGENCIA_TXT = fecha_larga(CORTE_ACT["vigencia"])
FUENTE_AOD = (f"Fuente: Cíclope, extracción del {fecha_larga(CORTE_ACT['extraccion'])} · "
              f"Intervenciones vigentes al {VIGENCIA_TXT}")
REGLA_VIGENCIA = "Vigente = fecha inicial ≤ fecha de vigencia ≤ fecha final"
VS_ANT = f" vs. {ANT}"
COL_USD_ACT, COL_USD_ANT = f"USD {ACT}", f"USD {ANT}"
CSS_ACT = FUENTES.get("CSS", {}).get("actualizacion", "")
FUENTE_CSS = f"Datos actualizados a {CSS_ACT} · {FUENTES.get('CSS', {}).get('responsable', '')}"

# \u2500\u2500 HEADER WITH LOGOS \u2500\u2500
col_title, col_logos = st.columns([3, 1])
with col_title:
    st.markdown(
        '<div style="padding: 0.6rem 0 0.2rem 0;">'
        '<div style="font-family:Montserrat,sans-serif;font-weight:800;font-size:1.25rem;color:#003087;line-height:1.3;">'
        'Ficha de caracterizaci\u00f3n territorial y sectorial de Cooperaci\u00f3n Internacional'
        '</div>'
        '<div style="font-size:0.82rem;color:#5A6A85;margin-top:4px;font-family:Source Sans 3,sans-serif;">'
        'Sistema Nacional de Cooperaci\u00f3n Internacional de Colombia'
        '</div></div>',
        unsafe_allow_html=True
    )
with col_logos:
    logo_col1, logo_col2 = st.columns(2)
    with logo_col1:
        try:
            st.image(LOGO_APC, width="stretch")
        except:
            st.markdown("APC")
    with logo_col2:
        try:
            st.image(LOGO_SNCIC, width="stretch")
        except:
            st.markdown("SNCIC")

st.markdown('<div style="height:4px;background:linear-gradient(90deg,#F5A623 33%,#003087 33% 66%,#C8102E 66%);margin-bottom:1.5rem;"></div>', unsafe_allow_html=True)

# \u2500\u2500 MAIN NAVIGATION \u2500\u2500
nav_options = ["\U0001f310 Panorama Nacional", "\U0001f5fa\ufe0f Ficha Territorial", "\U0001f3db\ufe0f Ficha Sectorial", "\U0001f4d6 Gu\u00eda de usuario"]
nav = st.radio("", nav_options, horizontal=True, label_visibility="collapsed",
               key="main_nav")

st.markdown("---")


DEPT_COL_INFO = "Departamento"
infogeneral = aplicar_planes_sncic(infogeneral, leer_planes(PLANES_FILE))
infogeneral["DEPT_NORM"] = infogeneral[DEPT_COL_INFO].map(norm_text)
depts = sorted(infogeneral[DEPT_COL_INFO].dropna().unique().tolist())

# Llave normalizada del departamento en cada fuente
for _df in (ciclope, ciclope_ant):
    _df["DEPT_NORM"] = _df["DEPARTAMENTO"].map(norm_text)
css["DEPT_NORM"] = css["ESPACIO VINCULADO"].map(norm_text)
colcol["DEPT_NORM"] = colcol["DEPARTAMENTOS PARTICIPANTES"].astype("string").map(norm_text)
contrapartidas["DEPT_NORM"] = (
    contrapartidas["Departamento"].astype("string").map(norm_text) if "Departamento" in contrapartidas.columns else ""
)
nombre_sector = {norm_text(s): s for s in sectores_visibles}

# Capas del mapa: valores por departamento (llave: nombre normalizado)
_cic_terr = ciclope[ciclope["DEPARTAMENTO"] != "Ámbito Nacional"]
CAPAS_MAPA = {
    "AOD · intervenciones": (
        _cic_terr.groupby("DEPT_NORM")["CODIGO INTERVENCION"].nunique().to_dict(), "Intervenciones",
        "Intervenciones de AOD vigentes en el departamento. No incluye las de ámbito nacional."),
    "AOD · aporte (USD)": (
        _cic_terr.groupby("DEPT_NORM")["VALOR APORTE (USD)"].sum().to_dict(), "USD",
        "Aporte estimado de AOD de las intervenciones vigentes. No incluye las de ámbito nacional."),
    "Sur-Sur · proyectos": (
        css.groupby("DEPT_NORM")["Código"].nunique().to_dict(), "Proyectos",
        "Proyectos Sur-Sur cuyo espacio vinculado es el departamento. Los proyectos de espacios sectoriales no aparecen en el mapa."),
    "ColCol · intercambios": (
        colcol.groupby("DEPT_NORM")["CODIGO"].nunique().to_dict(), "Intercambios",
        "Intercambios de Colombia Enseña Colombia en los que participa el departamento."),
}


def ir_a_departamento():
    """Al hacer clic en el mapa, abre la ficha territorial del departamento."""
    seleccion = st.session_state.get("mapa_nacional")
    puntos = seleccion["selection"]["points"] if seleccion else []
    if not puntos:
        return
    destino = geo_norm(puntos[0].get("location", ""))
    for d in depts:
        if norm_text(d) == destino:
            st.session_state["dept_sel"] = d
            st.session_state["main_nav"] = nav_options[1]
            return


# =============================================================
# COMPONENTES DE PRESENTACIÓN
# =============================================================
COLOR_VARIACION = "#5A6A85"
VIGENCIA_ANT_TXT = fecha_larga(CORTE_ANT["vigencia"])

# Explicación de cada indicador (ícono ⓘ de las tarjetas y desplegable "¿Cómo se calculan?")
AYUDA = {
    "intervenciones": (f"Intervenciones distintas (código de Cíclope) vigentes al {VIGENCIA_TXT}: fecha inicial ≤ "
                       f"{VIGENCIA_TXT} ≤ fecha final. Una intervención con varios municipios o cooperantes se cuenta una vez."),
    "movimiento": (f"Nuevas: vigentes ahora y no en el corte {ANT}. Continúan: vigentes en ambos cortes. "
                   f"Terminadas: vigentes en {ANT} cuya fecha final quedó antes del {VIGENCIA_TXT}. "
                   "Sin reporte: estaban en el corte anterior y ya no aparecen, sin haber terminado."),
    "cooperantes": "Actores distintos que aportan recursos a las intervenciones vigentes.",
    "municipios": ("Municipios o áreas no municipalizadas distintos con al menos una intervención vigente. "
                   "No incluye registros sin municipio reportado."),
    "departamentos": "Departamentos distintos con al menos una intervención vigente. No incluye el ámbito nacional.",
    "aporte": "Suma del valor de aporte en USD reportado en Cíclope para las intervenciones vigentes.",
    "variacion": f"Las flechas comparan con el corte {ANT} (intervenciones vigentes al {VIGENCIA_ANT_TXT}).",
    "plan": ("Según la app del SNCIC (hoja de espacios): si hay plan de trabajo activo, desde qué año "
             "y si hay capítulos especiales con ciudades."),
    "css": (f"Proyectos de Cooperación Sur-Sur aprobados y vigentes, contados una vez aunque participen varias "
            f"entidades. Datos a {CSS_ACT}."),
    "colcol": "Intercambios de Colombia Enseña Colombia, contados una vez aunque tengan varias etapas o departamentos.",
    "contrapartidas": "Proyectos con contrapartida de APC-Colombia; un proyecto en varios departamentos se cuenta una vez.",
}


def explicacion_indicadores():
    with st.expander("¿Cómo se calculan estos indicadores?"):
        st.markdown("\n".join(f"- **{t}:** {AYUDA[k]}" for t, k in (
            ("Intervenciones", "intervenciones"), ("Nuevas, continúan, terminadas", "movimiento"),
            ("Cooperantes", "cooperantes"), ("Municipios", "municipios"), ("Departamentos", "departamentos"),
            ("Aporte estimado", "aporte"), ("Variación", "variacion"))))


def tarjeta(label, valor, delta=None, sub=None, compacto=False, ayuda=None):
    """Tarjeta de indicador. delta = (texto, color); ayuda = explicación que aparece al pasar el cursor por ⓘ."""
    if ayuda:
        label += f' <span class="ayuda" title="{escapar_html(ayuda, quote=True)}">ⓘ</span>'
    estilo = ' style="font-size:1.1rem;"' if compacto else ""
    html = (f'<div class="metric-custom"><div class="metric-custom-label">{label}</div>'
            f'<div class="metric-custom-value"{estilo}>{valor}</div>')
    if delta:
        html += f'<div class="metric-custom-delta" style="color:{delta[1]};">{delta[0]}</div>'
    if sub:
        html += f'<div class="metric-custom-sub">{sub}</div>'
    st.markdown(html + "</div>", unsafe_allow_html=True)


def delta_vs_ant(act, ant, fmt=str):
    d = act - ant
    flecha = "▲ " if d >= 0 else "▼ "
    # Color neutro: que suba o baje no es en sí bueno o malo (p. ej., menos cooperantes)
    return flecha + fmt(abs(d)) + VS_ANT, COLOR_VARIACION


def seccion(titulo):
    st.markdown(f'<div class="section-header">{titulo}</div>', unsafe_allow_html=True)


def barras(data, cat, val, color="#003087", alto=220, titulo_x="", etiquetas=None):
    """Barras horizontales con nombres recortados (nombre completo al pasar el cursor) y valor al final."""
    d = data.copy()
    d["_nombre"] = d[cat].map(lambda x: etiquetas.get(x, x) if etiquetas else x).astype(str)
    d["_etq"] = d["_nombre"].map(lambda s: s if len(s) <= 40 else s[:38] + "…")
    d["_valor"] = d[val].map(valor_corto)
    base = alt.Chart(d).encode(
        y=alt.Y("_etq:N", sort="-x", title=""),
        x=alt.X(f"{val}:Q", title=None, axis=None),
    )
    barras_ = base.mark_bar(color=color, cornerRadiusTopRight=4, cornerRadiusBottomRight=4).encode(
        tooltip=[alt.Tooltip("_nombre:N", title=""), alt.Tooltip(f"{val}:Q", title=titulo_x or val, format=",.0f")])
    textos = base.mark_text(align="left", dx=4, fontSize=11, color="#1C2B4A").encode(text="_valor:N")
    st.altair_chart((barras_ + textos).properties(height=alto), width="stretch")


# ---------- Tablas legibles ----------
def resumen_intervenciones(df):
    """Una fila por intervención AOD (la fuente trae una fila por municipio y cooperante)."""
    if df.empty or "CODIGO INTERVENCION" not in df.columns:
        return pd.DataFrame()
    d = df.copy()
    d["FECHA INICIAL"], d["FECHA FINAL"] = a_fecha(d["FECHA INICIAL"]), a_fecha(d["FECHA FINAL"])
    d["_SECTOR"] = d["SECTOR_FICHA"].map(nombre_sector) if "SECTOR_FICHA" in d.columns else None
    t = agregar(d, "CODIGO INTERVENCION", {
        "Intervencion": ("NOMBRE INTERVENCION", "first"), "Objetivo": ("OBJETIVO GENERAL", "first"),
        "Inicio": ("FECHA INICIAL", "min"), "Fin": ("FECHA FINAL", "max"),
        "Cooperantes": ("NOMBRE ACTOR", unir), "Departamentos": ("DEPARTAMENTO", unir),
        "Municipios": ("MUNICIPIO", unir), "ODS": ("ODS", unir), "ENCI": ("ENCI PRIMER NIVEL", unir),
        "Sector": ("_SECTOR", unir), "USD": ("VALOR APORTE (USD)", "sum"),
    })
    return t.rename(columns={"CODIGO INTERVENCION": "Código", "Intervencion": "Intervención",
                             "ENCI": "Prioridad ENCI", "USD": "Aporte (USD)"}).sort_values("Aporte (USD)", ascending=False)


CONFIG_TABLA = {
    "Código": st.column_config.TextColumn(width="small"),
    "Intervención": st.column_config.TextColumn(width="large"),
    "Iniciativa": st.column_config.TextColumn(width="large"),
    "Intercambio": st.column_config.TextColumn(width="large"),
    "Proyecto": st.column_config.TextColumn(width="large"),
    "Inicio": st.column_config.DateColumn(format="DD/MM/YYYY"),
    "Fin": st.column_config.DateColumn(format="DD/MM/YYYY"),
    "Fecha de aprobación": st.column_config.DateColumn(format="DD/MM/YYYY"),
    "Aporte (USD)": st.column_config.NumberColumn(format="localized"),
}


def tabla_con_detalle(tabla, ocultas, key, detalle):
    """Tabla con selección de una fila; muestra el detalle de la fila elegida debajo."""
    ev = st.dataframe(tabla.drop(columns=ocultas, errors="ignore"), hide_index=True, width="stretch",
                      column_config=CONFIG_TABLA, on_select="rerun", selection_mode="single-row", key=key)
    filas = ev.selection.rows if ev is not None else []
    if filas:
        with st.container(border=True):
            detalle(tabla.iloc[filas[0]])


def detalle_texto(campos):
    for etiqueta, valor in campos:
        if valor is None or (isinstance(valor, float) and pd.isna(valor)) or norm_text(valor) in SIN_DATO:
            continue
        st.markdown(f"**{etiqueta}:** {valor}")


def lista_proyectos_aod(df, key, filtro_geo="MUNICIPIO"):
    """Listado de intervenciones AOD con búsqueda, filtros y detalle."""
    if df.empty:
        st.info(f"No hay intervenciones de AOD vigentes al {VIGENCIA_TXT}.")
        return
    total = df["CODIGO INTERVENCION"].nunique()
    f1, f2, f3, f4 = st.columns([2, 1.4, 1.4, 1.2])
    texto = f1.text_input("Buscar por nombre o código", key=f"{key}_q", placeholder="Ej.: agua, paz, R-AO-24")
    coops = f2.multiselect("Cooperante", sorted(df["NOMBRE ACTOR"].dropna().unique()), key=f"{key}_coop", placeholder="Todos")
    geo_vals = sorted(v for v in df[filtro_geo].dropna().unique() if norm_text(v) not in SIN_DATO)
    geos = f3.multiselect("Municipio" if filtro_geo == "MUNICIPIO" else "Departamento", geo_vals,
                          key=f"{key}_geo", placeholder="Todos")
    ods_vals = sorted(df["ODS"].dropna().unique(), key=lambda x: int(re.sub(r"\D", "", str(x)) or 99))
    ods = f4.multiselect("ODS", ods_vals, key=f"{key}_ods", placeholder="Todos")
    sel = df
    for col, valores in (("NOMBRE ACTOR", coops), (filtro_geo, geos), ("ODS", ods)):
        if valores:
            sel = sel[sel[col].isin(valores)]
    tabla = resumen_intervenciones(df[df["CODIGO INTERVENCION"].isin(sel["CODIGO INTERVENCION"])])
    if texto and not tabla.empty:
        q = norm_text(texto)
        tabla = tabla[(tabla["Intervención"].map(norm_text).str.contains(q, regex=False))
                      | (tabla["Código"].map(norm_text).str.contains(q, regex=False))]
    st.caption(f"Mostrando {len(tabla)} de {total} intervenciones vigentes · "
               "Selecciona una fila para ver el objetivo y el detalle.")
    if tabla.empty:
        st.info("Ninguna intervención coincide con los filtros.")
        return
    columnas_geo = ["Departamentos"] if filtro_geo == "MUNICIPIO" else []
    tabla_con_detalle(
        tabla, ["Objetivo", "Prioridad ENCI", "Sector"] + columnas_geo, f"{key}_tabla",
        lambda r: (st.markdown(f"#### {r['Intervención']}"), detalle_texto([
            ("Código", r["Código"]), ("Objetivo", r["Objetivo"]), ("Cooperantes", r["Cooperantes"]),
            ("Departamentos", r["Departamentos"]), ("Municipios", r["Municipios"]),
            ("Vigencia", f"{r['Inicio']:%d/%m/%Y} a {r['Fin']:%d/%m/%Y}" if pd.notna(r["Inicio"]) and pd.notna(r["Fin"]) else None),
            ("Aporte estimado", format_usd(r["Aporte (USD)"])), ("ODS", r["ODS"]),
            ("Prioridad ENCI", r["Prioridad ENCI"]), ("Sector", r["Sector"])])))


def resumen_css(df):
    """Una fila por proyecto Sur-Sur (la fuente repite el proyecto por cada entidad participante)."""
    if df.empty:
        return pd.DataFrame()
    d = df.copy()
    t = agregar(d, "Código", {
        "Iniciativa": ("NOMBRE DE LA INICIATIVA", "first"),
        "Objetivo": ("OBJETIVO GENERAL/DESCRIPCION DE LA INICIATIVA", "first"),
        "Pais": ("PAIS SOCIO", "first"), "Region": ("REGION", "first"), "Via": ("VIA DE COOPERACION", "first"),
        "Modalidad": ("MODALIDAD", "first"), "Aprobacion": ("FECHA DE APROBACION", "first"),
        "Nacionales": ("ENTIDAD(ES) NACIONAL(ES)", "first"), "Extranjeras": ("ENTIDAD(ES) EXTRANJERA(S)", "first"),
    })
    for col in ("Iniciativa", "Objetivo", "Pais", "Region", "Via", "Modalidad", "Nacionales", "Extranjeras"):
        if col not in t.columns:
            t[col] = ""
    t["Aprobacion"] = a_fecha(t["Aprobacion"]) if "Aprobacion" in t.columns else pd.NaT
    return t.rename(columns={"Pais": "País socio", "Region": "Región", "Via": "Vía de cooperación",
                             "Aprobacion": "Fecha de aprobación", "Nacionales": "Entidades nacionales",
                             "Extranjeras": "Entidades extranjeras"})


def tabla_css(df, key):
    tabla = resumen_css(df)
    if tabla.empty:
        st.info("No se encontraron proyectos de Cooperación Sur-Sur.")
        return
    st.caption(f"{len(tabla)} proyecto(s) · Selecciona una fila para ver el objetivo y las entidades.")
    tabla_con_detalle(
        tabla, ["Objetivo", "Entidades extranjeras", "Modalidad"], key,
        lambda r: (st.markdown(f"#### {r['Iniciativa']}"), detalle_texto([
            ("Código", r["Código"]), ("Objetivo", r["Objetivo"]), ("País socio", r["País socio"]),
            ("Vía de cooperación", r["Vía de cooperación"]), ("Modalidad", r["Modalidad"]),
            ("Entidades nacionales", r["Entidades nacionales"]), ("Entidades extranjeras", r["Entidades extranjeras"])])))


def resumen_colcol(df):
    """Una fila por intercambio ColCol (la fuente trae una fila por etapa y departamento)."""
    if df.empty:
        return pd.DataFrame()
    c = lambda *n: buscar_col(df, *n)
    presupuesto = c("PRESUPUESTO ESTIMADO APC COLOMBIA")
    d = df.copy()
    d["_ppto"] = pd.to_numeric(d[presupuesto], errors="coerce") if presupuesto else None
    t = agregar(d, c("CODIGO"), {
        "Intercambio": (c("NOMBRE DEL INTERCAMBIO"), "first"), "Objetivo": (c("OBJETIVO DEL INTERCAMBIO"), "first"),
        "Etapas": (c("ETAPA"), unir), "Linea": (c("LÍNEA TEMÁTICA"), unir), "Anio": (c("AÑO DE REALIZACIÓN"), unir),
        "Entidad": (c("ENTIDAD SOCIA NACIONAL"), unir), "Departamentos": (c("DEPARTAMENTOS PARTICIPANTES"), unir),
        "Practica": (c("BUENA PRÁCTICA"), "first"), "Presupuesto": ("_ppto", "max"),
    })
    for col in ("Intercambio", "Objetivo", "Etapas", "Departamentos", "Practica"):
        if col not in t.columns:
            t[col] = ""
    if "Presupuesto" not in t.columns:
        t["Presupuesto"] = float("nan")
    t["Presupuesto"] = t["Presupuesto"].map(lambda v: format_cop(v) if pd.notna(v) else "")
    return t.rename(columns={c("CODIGO"): "Código", "Linea": "Línea temática", "Anio": "Año",
                             "Entidad": "Entidad socia", "Practica": "Buena práctica",
                             "Presupuesto": "Presupuesto APC (COP)"})


def tabla_colcol(df, key):
    tabla = resumen_colcol(df)
    if tabla.empty:
        st.info("No se encontraron intercambios de Colombia Enseña Colombia.")
        return
    st.caption(f"{len(tabla)} intercambio(s) · Selecciona una fila para ver el objetivo y la buena práctica.")
    tabla_con_detalle(
        tabla, ["Objetivo", "Buena práctica"], key,
        lambda r: (st.markdown(f"#### {r['Intercambio']}"), detalle_texto([
            ("Código", r["Código"]), ("Objetivo", r["Objetivo"]), ("Buena práctica", r["Buena práctica"]),
            ("Departamentos participantes", r["Departamentos"]), ("Etapas", r["Etapas"])])))


def tabla_contrapartidas(df):
    if df.empty:
        st.info("No se encontraron contrapartidas.")
        return
    t = df.drop(columns=COLS_INTERNAS, errors="ignore").copy()
    t.columns = [str(c).strip() for c in t.columns]
    for col in ("Monto total", "Monto por APC"):
        if col in t.columns:
            t[col] = pd.to_numeric(t[col], errors="coerce").map(lambda v: format_cop(v) if pd.notna(v) else "")
    t = t.rename(columns={"AÑO": "Año", "Nombre del proyecto": "Proyecto",
                          "Monto total": "Monto total (COP)", "Monto por APC": "Aporte APC (COP)"})
    st.dataframe(t, hide_index=True, width="stretch", column_config=CONFIG_TABLA)


# ---------- Bloques reutilizables ----------
def bloque_aod(act, ant, tercera="municipios"):
    """Tarjetas y top 5 de AOD con comparativo frente al corte anterior."""
    st.caption(FUENTE_AOD)
    if act.empty:
        st.info(f"No hay intervenciones de AOD vigentes al {VIGENCIA_TXT}.")
        return
    c1, c2, c3, c4 = st.columns(4)
    n_act, n_ant = n_unicos(act, "CODIGO INTERVENCION"), n_unicos(ant, "CODIGO INTERVENCION")
    with c1:
        tarjeta("Intervenciones (únicas)", n_act, delta_vs_ant(n_act, n_ant), movimiento_intervenciones(act, ant),
                ayuda=f"{AYUDA['intervenciones']} {AYUDA['movimiento']}")
    with c2:
        tarjeta("Cooperantes", n_unicos(act, "NOMBRE ACTOR"),
                delta_vs_ant(n_unicos(act, "NOMBRE ACTOR"), n_unicos(ant, "NOMBRE ACTOR")), ayuda=AYUDA["cooperantes"])
    with c3:
        if tercera == "municipios":
            tarjeta("Municipios o áreas no municipalizadas", n_municipios(act),
                    delta_vs_ant(n_municipios(act), n_municipios(ant)), ayuda=AYUDA["municipios"])
        else:
            tarjeta("Departamentos", n_departamentos(act), delta_vs_ant(n_departamentos(act), n_departamentos(ant)),
                    ayuda=AYUDA["departamentos"])
    with c4:
        usd, usd_ant = act["VALOR APORTE (USD)"].sum(), ant["VALOR APORTE (USD)"].sum() if not ant.empty else 0
        tarjeta("Aporte estimado (USD)", format_usd(usd), delta_vs_ant(usd, usd_ant, format_usd), compacto=True,
                ayuda=AYUDA["aporte"])
    explicacion_indicadores()
    k1, k2 = st.columns(2)
    with k1:
        st.markdown("**Top 5 cooperantes por aporte (USD)**")
        top = top_by_sum(act, "NOMBRE ACTOR", "VALOR APORTE (USD)", 5)
        barras(top, "NOMBRE ACTOR", "VALOR APORTE (USD)", "#003087", 200, "USD")
        st.dataframe(tabla_top_comparativa(top, ant, "NOMBRE ACTOR").rename(columns={"NOMBRE ACTOR": "Cooperante"}),
                     hide_index=True, width="stretch")
    with k2:
        st.markdown("**Top 5 ODS por aporte (USD)**")
        top = top_by_sum(act, "ODS", "VALOR APORTE (USD)", 5)
        barras(top, "ODS", "VALOR APORTE (USD)", "#1565C0", 200, "USD", etiquetas=ODS_NOMBRES)
        st.dataframe(tabla_top_comparativa(top, ant, "ODS", nombres=ODS_NOMBRES), hide_index=True, width="stretch")


def barra_descargas(nombre_archivo, excel, pdf):
    _, b1, b2 = st.columns([3, 1, 1])
    with b1:
        st.download_button("\U0001f4e5 Excel", data=excel, file_name=f"{nombre_archivo}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", width="stretch")
    with b2:
        st.download_button("\U0001f4e5 PDF", data=pdf, file_name=f"{nombre_archivo}.pdf",
                           mime="application/pdf", width="stretch")


@st.cache_data(show_spinner=False)
def to_excel_sector(info_sector, aod_sector, css_sector, colcol_sector, proyectos):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        if not info_sector.empty:
            info_sector.T.reset_index().set_axis(["Campo", "Valor"], axis=1).to_excel(writer, sheet_name="Info General", index=False)
        for df, hoja in ((aod_sector, "AOD"), (css_sector, "Coop Sur Sur"), (colcol_sector, "ColCol"),
                         (proyectos, "Proyectos AOD activos")):
            if not df.empty:
                df.drop(columns=COLS_INTERNAS, errors="ignore").to_excel(writer, sheet_name=hoja, index=False)
    return output.getvalue()


FOOTER = '<div class="apc-footer">Agencia Presidencial de Cooperación Internacional de Colombia · APC-Colombia</div>'


# =============================================================
# FICHA TERRITORIAL
# =============================================================
if nav == nav_options[1]:

    dept = st.selectbox("\U0001f5fa️ Selecciona un departamento", depts, key="dept_sel")
    dn = norm_text(dept)
    info = infogeneral[infogeneral["DEPT_NORM"] == dn].head(1)
    cic_dept = ciclope[ciclope["DEPT_NORM"] == dn]
    cic_dept_ant = ciclope_ant[ciclope_ant["DEPT_NORM"] == dn]
    css_dept = css[css["DEPT_NORM"] == dn]
    colcol_dept = colcol[colcol["DEPT_NORM"] == dn]
    contr_dept = contrapartidas[contrapartidas["DEPT_NORM"] == dn]

    st.markdown(f'<div class="dept-title-banner">\U0001f4cd {dept}</div>', unsafe_allow_html=True)
    barra_descargas(
        f"Ficha_Territorial_{dept}",
        to_excel_ficha(info.drop(columns=COLS_INTERNAS, errors="ignore"), cic_dept, colcol_dept,
                       contr_dept.drop(columns=COLS_INTERNAS, errors="ignore"), css_dept, resumen_intervenciones(cic_dept)),
        to_pdf_ficha(dept, info.drop(columns=COLS_INTERNAS, errors="ignore"), cic_dept, cic_dept_ant, colcol_dept,
                     contr_dept.drop(columns=COLS_INTERNAS, errors="ignore"), css_dept),
    )

    t_res, t_aod, t_css, t_apc, t_proy = st.tabs(["Resumen", "AOD", "Sur-Sur", "Programas APC", "Proyectos AOD"])

    with t_res:
        col_info, col_mapa = st.columns([2, 1])
        with col_mapa:
            st.plotly_chart(
                make_map(geo, CAPAS_MAPA["AOD · intervenciones"][0], "Intervenciones", ":,.0f",
                         selected_dept=dept, alto=340, escala=False),
                width="stretch", key=f"mini_mapa_{dn}", config={"displayModeBar": False})
            st.caption("Departamento resaltado en rojo. El azul indica cuántas intervenciones AOD vigentes "
                       "tiene cada departamento.")
        with col_info:
            seccion("Información general")
        if info.empty:
            col_info.warning("No se encontró el departamento en la tabla de información general.")
        else:
            r = info.iloc[0]
            pob = get_col(r, "Población", "Poblacion")
            try:
                pob = f"{int(float(pob)):,}".replace(",", ".")
            except (TypeError, ValueError):
                pass
            tiene_plan = str(get_col(r, "¿Cuenta con plan de trabajo en el marco del SNCIC?")).strip().lower() in ("sí", "si")
            anio_plan = get_col(r, "Año creación del plan de trabajo")
            cap = get_col(r, "Cuál")
            sub_plan = []
            if tiene_plan and str(anio_plan).isdigit():
                sub_plan.append(f"Desde {anio_plan}")
            if cap and norm_text(cap) not in SIN_DATO:
                sub_plan.append(f"Capítulo especial: {cap}")
            with col_info:
                g1, g2 = st.columns(2)
                with g1:
                    tarjeta("Capital", str(get_col(r, "Capital")).title())
                with g2:
                    tarjeta("Número de municipios", get_col(r, "Número de Municipios", "Numero de Municipios"))
                st.markdown('<div style="height:0.6rem"></div>', unsafe_allow_html=True)
                g3, g4 = st.columns(2)
                with g3:
                    tarjeta("Población", pob)
                with g4:
                    tarjeta("Plan de trabajo SNCIC", "Sí" if tiene_plan else "No", sub=" · ".join(sub_plan) or None,
                            ayuda=AYUDA["plan"])

        seccion("La cooperación en el departamento")
        s1, s2, s3, s4, s5 = st.columns(5)
        with s1:
            tarjeta("Intervenciones AOD", n_unicos(cic_dept, "CODIGO INTERVENCION"), sub="Pestaña AOD",
                    ayuda=AYUDA["intervenciones"])
        with s2:
            tarjeta("Aporte AOD (USD)", valor_corto(cic_dept["VALOR APORTE (USD)"].sum()), sub="Pestaña AOD", compacto=True,
                    ayuda=AYUDA["aporte"])
        with s3:
            tarjeta("Proyectos Sur-Sur", n_unicos(css_dept, "Código"), sub="Pestaña Sur-Sur", ayuda=AYUDA["css"])
        with s4:
            tarjeta("Intercambios ColCol", n_unicos(colcol_dept, "CODIGO"), sub="Pestaña Programas APC",
                    ayuda=AYUDA["colcol"])
        with s5:
            tarjeta("Contrapartidas", len(proyectos_contrapartida(contr_dept)), sub="Pestaña Programas APC",
                    ayuda=AYUDA["contrapartidas"])
        st.caption(f"AOD: intervenciones vigentes al {VIGENCIA_TXT}. Sur-Sur: {CSS_ACT}.")

        if not info.empty:
            with st.expander("Ver registro completo del departamento"):
                det = pd.DataFrame(registro_publico(info.iloc[0]), columns=["Campo", "Valor"])
                st.dataframe(det.astype(str), hide_index=True, width="stretch")

    with t_aod:
        bloque_aod(cic_dept, cic_dept_ant, "municipios")

    with t_css:
        seccion("Proyectos de Cooperación Sur-Sur aprobados y vigentes")
        st.caption(FUENTE_CSS)
        tabla_css(css_dept, f"css_{dn}")

    with t_apc:
        seccion("Colombia Enseña Colombia (ColCol)")
        tabla_colcol(colcol_dept, f"colcol_{dn}")
        seccion("Contrapartidas")
        tabla_contrapartidas(contr_dept)

    with t_proy:
        seccion("Proyectos AOD vigentes")
        st.caption(f"{FUENTE_AOD} · {REGLA_VIGENCIA}")
        lista_proyectos_aod(cic_dept, f"aod_{dn}", "MUNICIPIO")

    st.markdown(FOOTER, unsafe_allow_html=True)


# =============================================================
# FICHA SECTORIAL
# =============================================================
elif nav == nav_options[2]:

    sector = st.selectbox("\U0001f3db️ Selecciona un sector", sectores_visibles, key="sector_sel")
    sn = norm_text(sector)
    info_sector = info_s[info_s["Nombre del sector"].map(norm_text) == sn]
    aod_sector = aod_s[aod_s["SECTOR_FICHA"] == sn].copy()
    aod_sector_ant = aod_s_ant[aod_s_ant["SECTOR_FICHA"] == sn].copy()
    css_sector = css_s[css_s["SECTOR_FICHA"] == sn]
    colcol_sector = colcol_s[colcol_s["SECTOR_FICHA"] == sn]

    st.markdown(f'<div class="dept-title-banner">\U0001f3db️ {sector}</div>', unsafe_allow_html=True)
    barra_descargas(
        f"Ficha_Sectorial_{sector}",
        to_excel_sector(info_sector, aod_sector, css_sector, colcol_sector, resumen_intervenciones(aod_sector)),
        to_pdf_sectorial(sector, info_sector, aod_sector, aod_sector_ant, css_sector, colcol_sector),
    )

    t_res, t_aod, t_css, t_cc, t_proy = st.tabs(["Resumen", "AOD", "Sur-Sur", "ColCol", "Proyectos AOD"])

    with t_res:
        seccion("Información general del sector")
        if not info_sector.empty:
            r = info_sector.iloc[0]
            con_plan = str(r.get("Cabezas del sector asociadas con planes de trabajo en el marco del SNCIC", "")).strip().lower() in ("sí", "si", "yes")
            planes = str(r.get("Planes de trabajo vinculados", ""))
            g1, g2 = st.columns([2, 1])
            with g1:
                tarjeta("Cabeza de sector", str(r.get("Cabeza de sector", "")).strip(), compacto=True)
            with g2:
                tarjeta("Plan de trabajo SNCIC", "Sí" if con_plan else "No",
                        sub=titulo_es(planes) if norm_text(planes) not in SIN_DATO else None)
        seccion("La cooperación en el sector")
        s1, s2, s3, s4 = st.columns(4)
        with s1:
            tarjeta("Intervenciones AOD", n_unicos(aod_sector, "CODIGO INTERVENCION"), sub="Pestaña AOD",
                    ayuda=AYUDA["intervenciones"])
        with s2:
            tarjeta("Aporte AOD (USD)", valor_corto(aod_sector["VALOR APORTE (USD)"].sum()), sub="Pestaña AOD", compacto=True,
                    ayuda=AYUDA["aporte"])
        with s3:
            tarjeta("Proyectos Sur-Sur", n_unicos(css_sector, "Código"), sub="Pestaña Sur-Sur", ayuda=AYUDA["css"])
        with s4:
            tarjeta("Intercambios ColCol", n_unicos(colcol_sector, "CODIGO"), sub="Pestaña ColCol", ayuda=AYUDA["colcol"])
        st.caption(f"AOD: intervenciones vigentes al {VIGENCIA_TXT}. Sur-Sur: {CSS_ACT}.")

    with t_aod:
        bloque_aod(aod_sector, aod_sector_ant, "departamentos")

    with t_css:
        seccion("Proyectos de Cooperación Sur-Sur")
        st.caption(FUENTE_CSS)
        tabla_css(css_sector, f"css_s_{sn}")

    with t_cc:
        seccion("Colombia Enseña Colombia (ColCol)")
        tabla_colcol(colcol_sector, f"colcol_s_{sn}")

    with t_proy:
        seccion("Proyectos AOD vigentes del sector")
        st.caption(f"{FUENTE_AOD} · {REGLA_VIGENCIA}")
        lista_proyectos_aod(aod_sector, f"aod_s_{sn}", "DEPARTAMENTO")

    st.markdown(FOOTER, unsafe_allow_html=True)


# =============================================================
# PANORAMA NACIONAL
# =============================================================
elif nav == nav_options[0]:

    st.markdown('<div class="dept-title-banner">\U0001f310 Panorama Nacional de la Cooperación Internacional</div>',
                unsafe_allow_html=True)

    css_unicos = css.drop_duplicates(subset=["Código"])
    contr_unicas = proyectos_contrapartida(contrapartidas)

    seccion("Ayuda Oficial al Desarrollo (AOD)")
    st.caption(FUENTE_AOD + ". Incluye ámbito nacional y territorial.")
    n1, n2, n3, n4 = st.columns(4)
    n_act, n_ant = n_unicos(ciclope, "CODIGO INTERVENCION"), n_unicos(ciclope_ant, "CODIGO INTERVENCION")
    with n1:
        tarjeta("Intervenciones (únicas)", n_act, delta_vs_ant(n_act, n_ant), movimiento_intervenciones(ciclope, ciclope_ant),
                ayuda=f"{AYUDA['intervenciones']} {AYUDA['movimiento']}")
    with n2:
        tarjeta("Cooperantes", n_unicos(ciclope, "NOMBRE ACTOR"),
                delta_vs_ant(n_unicos(ciclope, "NOMBRE ACTOR"), n_unicos(ciclope_ant, "NOMBRE ACTOR")),
                ayuda=AYUDA["cooperantes"])
    with n3:
        tarjeta("Departamentos con AOD", n_departamentos(ciclope),
                delta_vs_ant(n_departamentos(ciclope), n_departamentos(ciclope_ant)), ayuda=AYUDA["departamentos"])
    with n4:
        usd, usd_ant = ciclope["VALOR APORTE (USD)"].sum(), ciclope_ant["VALOR APORTE (USD)"].sum()
        tarjeta("Aporte estimado (USD)", "USD " + valor_corto(usd), delta_vs_ant(usd, usd_ant, format_usd), compacto=True,
                ayuda=AYUDA["aporte"])
    explicacion_indicadores()

    seccion("Cooperación Sur-Sur y programas de APC-Colombia")
    o1, o2, o3, o4 = st.columns(4)
    with o1:
        tarjeta("Proyectos Sur-Sur", len(css_unicos), sub=f"Datos a {CSS_ACT}", ayuda=AYUDA["css"])
    with o2:
        paises = {p.strip() for v in css_unicos["PAIS SOCIO"].dropna() for p in re.split(r"[,;/]", str(v))
                  if norm_text(p) not in SIN_DATO}
        tarjeta("Países socios Sur-Sur", len(paises))
    with o3:
        tarjeta("Intercambios ColCol", n_unicos(colcol, "CODIGO"), ayuda=AYUDA["colcol"])
    with o4:
        tarjeta("Contrapartidas", len(contr_unicas), ayuda=AYUDA["contrapartidas"])

    seccion("Mapa de la cooperación por departamento")
    capa = st.radio("Capa del mapa", list(CAPAS_MAPA), horizontal=True, key="capa_mapa")
    valores_mapa, etiqueta_mapa, nota_mapa = CAPAS_MAPA[capa]
    st.plotly_chart(
        make_map(geo, valores_mapa, etiqueta_mapa, ":,.0f"),
        width="stretch", key="mapa_nacional",
        on_select=ir_a_departamento, selection_mode="points",
        config={"displayModeBar": False},
    )
    st.caption(f"{nota_mapa} Haz clic en un departamento para abrir su ficha territorial.")

    p_aod, p_css, p_apc = st.tabs(["AOD", "Sur-Sur", "Programas APC"])

    with p_aod:
        seccion("Cooperantes")
        a1, a2 = st.columns(2)
        with a1:
            st.markdown("**Top 10 cooperantes por aporte (USD)**")
            top = top_by_sum(ciclope, "NOMBRE ACTOR", "VALOR APORTE (USD)", 10)
            barras(top, "NOMBRE ACTOR", "VALOR APORTE (USD)", "#003087", 300, "USD")
        with a2:
            st.markdown("**Top 10 cooperantes por número de intervenciones**")
            top = (ciclope.groupby("NOMBRE ACTOR")["CODIGO INTERVENCION"].nunique()
                   .sort_values(ascending=False).head(10).reset_index(name="Intervenciones"))
            barras(top, "NOMBRE ACTOR", "Intervenciones", "#1565C0", 300, "Intervenciones")
        seccion("ODS y sectores")
        a3, a4 = st.columns(2)
        with a3:
            st.markdown("**Top 10 ODS por aporte (USD)**")
            barras(top_by_sum(ciclope, "ODS", "VALOR APORTE (USD)", 10), "ODS", "VALOR APORTE (USD)",
                   "#003087", 300, "USD", etiquetas=ODS_NOMBRES)
        with a4:
            st.markdown("**Top 10 sectores por aporte (USD)**")
            top = top_by_sum(ciclope.assign(Sector=ciclope["SECTOR_FICHA"].map(nombre_sector)).dropna(subset=["Sector"]),
                             "Sector", "VALOR APORTE (USD)", 10)
            barras(top, "Sector", "VALOR APORTE (USD)", "#1565C0", 300, "USD")
        seccion("Departamentos")
        a5, a6 = st.columns(2)
        with a5:
            st.markdown("**Top 10 departamentos por aporte (USD)**")
            barras(top_by_sum(_cic_terr, "DEPARTAMENTO", "VALOR APORTE (USD)", 10), "DEPARTAMENTO",
                   "VALOR APORTE (USD)", "#003087", 300, "USD")
        with a6:
            st.markdown("**Top 10 departamentos por número de intervenciones**")
            top = (_cic_terr.groupby("DEPARTAMENTO")["CODIGO INTERVENCION"].nunique()
                   .sort_values(ascending=False).head(10).reset_index(name="Intervenciones"))
            barras(top, "DEPARTAMENTO", "Intervenciones", "#1565C0", 300, "Intervenciones")

    with p_css:
        st.caption(FUENTE_CSS)
        b1, b2 = st.columns(2)
        with b1:
            st.markdown("**Proyectos por región**")
            barras(css_unicos["REGION"].value_counts().reset_index(name="Proyectos"), "REGION", "Proyectos",
                   "#003087", 220, "Proyectos")
        with b2:
            st.markdown("**Proyectos por vía de cooperación**")
            barras(css_unicos["VIA DE COOPERACION"].value_counts().reset_index(name="Proyectos"),
                   "VIA DE COOPERACION", "Proyectos", "#1565C0", 220, "Proyectos")
        b3, b4 = st.columns(2)
        with b3:
            st.markdown("**Top 10 países socios**")
            barras(css_unicos["PAIS SOCIO"].value_counts().head(10).reset_index(name="Proyectos"), "PAIS SOCIO",
                   "Proyectos", "#003087", 300, "Proyectos")
        with b4:
            st.markdown("**Proyectos por sector**")
            por_sector = (css_s.assign(Sector=css_s["SECTOR_FICHA"].map(nombre_sector)).dropna(subset=["Sector"])
                          .groupby("Sector")["Código"].nunique().sort_values(ascending=False).head(10)
                          .reset_index(name="Proyectos"))
            barras(por_sector, "Sector", "Proyectos", "#1565C0", 300, "Proyectos")
        seccion("Proyectos de Cooperación Sur-Sur")
        tabla_css(css, "css_nacional")

    with p_apc:
        seccion("Colombia Enseña Colombia (ColCol)")
        c_dep = (colcol[~colcol["DEPT_NORM"].isin(SIN_DATO)].groupby("DEPARTAMENTOS PARTICIPANTES")["CODIGO"]
                 .nunique().sort_values(ascending=False).head(10).reset_index(name="Intercambios"))
        e1, e2 = st.columns([1, 2])
        with e1:
            tarjeta("Intercambios", n_unicos(colcol, "CODIGO"))
            st.markdown("<br>", unsafe_allow_html=True)
            tarjeta("Departamentos participantes", colcol.loc[~colcol["DEPT_NORM"].isin(SIN_DATO), "DEPT_NORM"].nunique())
        with e2:
            st.markdown("**Departamentos con más intercambios**")
            barras(c_dep, "DEPARTAMENTOS PARTICIPANTES", "Intercambios", "#003087", 260, "Intercambios")
        tabla_colcol(colcol, "colcol_nacional")
        seccion("Contrapartidas")
        aporte_apc = pd.to_numeric(contr_unicas.get("Monto por APC"), errors="coerce").sum() if not contr_unicas.empty else 0
        f1, f2, f3 = st.columns(3)
        with f1:
            tarjeta("Proyectos", len(contr_unicas))
        with f2:
            tarjeta("Departamentos", contrapartidas.loc[~contrapartidas["DEPT_NORM"].isin(SIN_DATO), "DEPT_NORM"].nunique())
        with f3:
            tarjeta("Aporte APC (COP)", format_cop(aporte_apc), compacto=True)
        tabla_contrapartidas(contrapartidas)

    st.markdown(FOOTER, unsafe_allow_html=True)


# =============================================================
# GUIA DE USUARIO
# =============================================================
elif nav == nav_options[3]:

    st.markdown('<div class="dept-title-banner">Guía de usuario</div>', unsafe_allow_html=True)

    guia_html = (
        '<div class="guia-card">'
        '<div class="guia-intro">¿Qué es esta herramienta?</div>'
        '<p>Esta herramienta permite conocer cómo se está moviendo la cooperación internacional en Colombia, '
        'tanto desde una perspectiva territorial como sectorial. Integra información de la Ayuda Oficial al '
        'Desarrollo (AOD), la cooperación Sur-Sur y los programas de APC-Colombia (Colombia Enseña Colombia y '
        'Contrapartidas).</p>'
        '<p><strong>Panorama Nacional</strong> presenta los totales nacionales de AOD, Sur-Sur y programas de '
        'APC-Colombia, y un mapa por departamento con varias capas (intervenciones y aporte de AOD, proyectos '
        'Sur-Sur e intercambios ColCol). Al hacer clic en un departamento del mapa se abre su ficha territorial. '
        'Debajo del mapa, las pestañas AOD, Sur-Sur y Programas APC muestran los principales cooperantes, ODS, '
        f'sectores, países socios y regiones. Los indicadores de AOD se comparan con el corte anterior ({ANT}).</p>'
        '<p><strong>Ficha Territorial</strong> reúne la cooperación de cada uno de los 33 departamentos '
        '(incluida Bogotá, D.C.) en pestañas: <em>Resumen</em> (información general, plan de trabajo SNCIC y '
        'totales), <em>AOD</em>, <em>Sur-Sur</em>, <em>Programas APC</em> y <em>Proyectos AOD</em>. '
        'Las descargas en Excel y PDF están arriba, junto al nombre del departamento.</p>'
        f'<p><strong>Ficha Sectorial</strong> ofrece lo mismo para los {len(sectores_visibles)} sectores de '
        'gobierno, con las pestañas Resumen, AOD, Sur-Sur, ColCol y Proyectos AOD.</p>'
        '<p>En las tablas puede <strong>buscar y filtrar</strong> (por cooperante, municipio o departamento y ODS) '
        'y <strong>seleccionar una fila</strong> para ver el objetivo y el detalle del proyecto. '
        'Las flechas ▲ (subió) y ▼ (bajó) indican la variación frente al corte anterior; el color es neutro '
        'porque subir o bajar no es en sí bueno o malo. Pase el cursor por el ícono ⓘ de cada tarjeta para ver cómo se calcula. '
        'En la tarjeta de intervenciones: ♥ nuevas &nbsp;|&nbsp; ↺ continúan &nbsp;|&nbsp; ✓ terminadas.</p>'
        '</div>'
    )
    st.markdown(guia_html, unsafe_allow_html=True)

    g1, g2, g3, g4 = st.columns(4)
    with g1:
        st.info("**\U0001f310 Panorama Nacional**\n\nVisión nacional de AOD, Sur-Sur y programas de APC, con mapa.")
    with g2:
        st.info("**\U0001f5fa️ Ficha Territorial**\n\nCooperación por departamento, organizada en pestañas.")
    with g3:
        st.info("**\U0001f3db️ Ficha Sectorial**\n\nCooperación por sector de gobierno, organizada en pestañas.")
    with g4:
        st.info(f"**Fuentes**\n\nAOD: Cíclope, extracción del {fecha_larga(CORTE_ACT['extraccion'])}; "
                f"intervenciones vigentes al {VIGENCIA_TXT}. Sur-Sur: APC-Colombia ({CSS_ACT}). "
                "ColCol y Contrapartidas: APC-Colombia, DCI. Planes de trabajo: SNCIC.")

    st.markdown(FOOTER, unsafe_allow_html=True)
