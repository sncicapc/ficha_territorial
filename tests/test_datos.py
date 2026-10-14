"""Integridad de los datos reales y de la configuración (leen los Excel de la carpeta app)."""
import pandas as pd
import pytest

import logica as L


@pytest.fixture(scope="module")
def datos():
    cfg = L.cargar_config("configuracion.xlsx")
    act, ant = cfg["cortes"]["actual"], cfg["cortes"]["anterior"]
    x = L.cargar_xlsm("Ficha_territorial.xlsm")
    s = L.cargar_sectores("Ficha sectores.xlsx", act["hoja"], ant["hoja"])
    return cfg, x, s


def test_configuracion_tiene_los_dos_cortes(datos):
    cfg, _, _ = datos
    act, ant = cfg["cortes"]["actual"], cfg["cortes"]["anterior"]
    assert ant["vigencia"] < act["vigencia"], "La vigencia del corte anterior debe ser anterior a la actual"
    assert act["etiqueta"] and ant["etiqueta"]


def test_fuentes_tienen_las_columnas_que_usa_la_app(datos):
    _, x, s = datos
    problemas = L.validar_columnas({
        "infogeneral": ("infogeneral", x["infogeneral"]), "aod": ("AOD actual", s["aod"]),
        "aod_ant": ("AOD anterior", s["aod_ant"]), "css": ("CSS", x["css"]), "colcol": ("colcol", x["colcol"]),
        "contrapartidas": ("contrapartidas", x["contrapartidas"]), "info_sectores": ("INFO GENERAL", s["info"]),
        "colcol_sectores": ("COLCOL", s["colcol"]),
    })
    assert problemas == []


def test_no_se_cargan_datos_personales_de_enlaces(datos):
    _, x, _ = datos
    columnas = {L.norm_text(c) for c in x["infogeneral"].columns}
    assert not columnas & {L.norm_text(c) for c in L.COLUMNAS_NO_PUBLICAS}


def test_hay_intervenciones_vigentes_en_ambos_cortes(datos):
    cfg, _, s = datos
    act = L.activas_a_fecha(s["aod"], cfg["cortes"]["actual"]["vigencia"])
    ant = L.activas_a_fecha(s["aod_ant"], cfg["cortes"]["anterior"]["vigencia"])
    assert act["CODIGO INTERVENCION"].nunique() > 0 and ant["CODIGO INTERVENCION"].nunique() > 0


def test_cooperantes_con_nombre_individual(datos):
    """Criterio de APC (13/10/2026): los cooperantes se muestran con su nombre, sin agrupar como 'Filantropía'."""
    _, _, s = datos
    for corte in ("aod", "aod_ant"):
        agrupados = set(s[corte]["NOMBRE ACTOR"].map(L.norm_text)) & {"FILANTROPIA", "BID FCS"}
        assert not agrupados, f"{corte}: hay cooperantes agrupados {agrupados}"


def test_departamentos_de_las_fuentes_existen_en_infogeneral(datos):
    _, x, s = datos
    deptos = set(x["infogeneral"]["Departamento"].map(L.norm_text)) | {L.norm_text(L.AMBITO_NACIONAL)}
    assert set(s["aod"]["DEPARTAMENTO"].map(L.norm_text)) <= deptos
    assert set(x["contrapartidas"]["Departamento"].map(L.norm_text)) <= deptos
    colcol = set(x["colcol"]["DEPARTAMENTOS PARTICIPANTES"].map(L.norm_text)) - L.SIN_DATO
    assert colcol <= deptos


def test_todos_los_intercambios_colcol_tienen_sector(datos):
    cfg, x, s = datos
    sectores = {L.norm_text(v) for v in s["info"]["Nombre del sector"]} - cfg["ocultos"]
    cc = L.unificar_colcol(x["colcol"], s["colcol"])
    asignado = L.asignar_sector(cc, "SECTOR VINCULADO", "COLCOL", sectores, "CODIGO",
                                cfg["equivalencias"], cfg["registros"])
    sin_sector = sorted(cc.loc[asignado.isna(), "CODIGO"].astype(str).unique())
    assert sin_sector == [], f"Intercambios sin sector: {sin_sector}. Asígnelos en la hoja REGISTROS."


def test_igualdad_no_aparece_como_sector_y_sus_registros_se_reubican(datos):
    cfg, _, s = datos
    assert L.norm_text("Igualdad y Equidad") in cfg["ocultos"]
    sectores = {L.norm_text(v) for v in s["info"]["Nombre del sector"]} - cfg["ocultos"]
    aod = s["aod"][s["aod"]["SECTORES GOB"].map(L.norm_text) == "IGUALDAD Y EQUIDAD"]
    asignado = L.asignar_sector(aod, "SECTORES GOB", "AOD", sectores, "CODIGO INTERVENCION",
                                cfg["equivalencias"], cfg["registros"])
    assert asignado.notna().all(), "Hay intervenciones de Igualdad sin reubicar en REGISTROS"
    assert L.norm_text("Igualdad y Equidad") not in set(asignado)


def test_planes_sncic_corresponden_a_departamentos(datos):
    _, x, _ = datos
    planes = L.leer_planes("planes_trabajo_sncic.xlsx")
    assert planes is not None
    assert set(planes.index) <= set(pd.to_numeric(x["infogeneral"]["Divipola"]).astype(int))
