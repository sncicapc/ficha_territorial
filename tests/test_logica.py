"""Reglas de negocio con datos pequeños de ejemplo (rápidas, no leen los Excel)."""
import pandas as pd

import logica as L


# ---------- Texto y formatos ----------
def test_norm_text_ignora_tildes_signos_y_mayusculas():
    assert L.norm_text("Bogotá, D.C.") == "BOGOTA D C"
    assert L.norm_text("  Nariño ") == "NARINO"
    assert L.norm_text(None) == ""


def test_santander_no_coincide_con_norte_de_santander():
    """Hallazgo de la Fase 1: con 'contiene', SANTANDER incluía NORTE DE SANTANDER."""
    cc = pd.DataFrame({"DEPARTAMENTOS PARTICIPANTES": ["SANTANDER", "NORTE DE SANTANDER", "CAUCA", "VALLE DEL CAUCA"]})
    llave = cc["DEPARTAMENTOS PARTICIPANTES"].map(L.norm_text)
    assert (llave == L.norm_text("Santander")).sum() == 1
    assert (llave == L.norm_text("Cauca")).sum() == 1


def test_formatos_colombianos():
    assert L.format_usd(1234567.4) == "USD 1.234.567"
    assert L.format_cop(2500) == "$ 2.500"
    assert L.format_usd("no es número") == ""
    assert L.valor_corto(1_772_000_000) == "1.772 M"
    assert L.valor_corto(72_500_000) == "72,5 M"
    assert L.valor_corto(350_000) == "350 mil"
    assert L.valor_corto(52) == "52"


def test_variacion():
    assert L.variacion_txt(150, 100) == "▲ 50%"
    assert L.variacion_txt(97, 100) == "▼ 3%"
    assert L.variacion_txt(100, 100) == "Sin cambio"
    assert L.variacion_txt(10, 0) == "Nuevo"


def test_titulo_es():
    assert L.titulo_es("AVANCE SOCIAL, IGUALDAD Y EQUIDAD") == "Avance Social, Igualdad y Equidad"


# ---------- Vigencia y movimiento entre cortes ----------
def _aod(filas):
    return pd.DataFrame(filas, columns=["CODIGO INTERVENCION", "FECHA INICIAL", "FECHA FINAL"])


def test_vigentes_incluye_los_limites_y_excluye_sin_fecha():
    corte = pd.Timestamp("2026-06-30")
    df = _aod([
        ("A", "2026-06-30", "2027-01-01"),   # empieza el día del corte -> vigente
        ("B", "2020-01-01", "2026-06-30"),   # termina el día del corte -> vigente
        ("C", "2020-01-01", "2026-06-29"),   # terminó antes
        ("D", "2026-07-01", "2027-01-01"),   # empieza después
        ("E", None, "2027-01-01"),           # sin fecha inicial
    ])
    assert set(L.activas_a_fecha(df, corte)["CODIGO INTERVENCION"]) == {"A", "B"}


def test_movimiento_entre_cortes():
    ant = _aod([("A", "2025-01-01", "2026-12-31"), ("B", "2025-01-01", "2026-05-01"),
                ("C", "2025-01-01", "2027-12-31")])
    act = _aod([("A", "2025-01-01", "2026-12-31"), ("N", "2026-04-01", "2027-12-31")])
    m = L.contar_movimiento(act, ant, pd.Timestamp("2026-06-30"))
    # A continúa, N es nueva, B terminó (fecha final antes del corte), C salió sin terminar
    assert m == {"nuevas": 1, "continuan": 1, "terminadas": 1, "sin_reporte": 1}


# ---------- Asignación de sectores ----------
SECTORES = {"CULTURAS ARTES Y SABERES", "INTERIOR", "AMBIENTE Y DESARROLLO SOSTENIBLE", "SALUD Y PROTECCION SOCIAL"}
EQUIV = {("AOD", "CULTURA"): "CULTURAS ARTES Y SABERES",
         ("AOD", "IGUALDAD Y EQUIDAD"): "POR PROYECTO",
         ("AOD", "RAMA JUDICIAL"): "SIN SECTOR"}
REGISTROS = {("AOD", "R-1"): "INTERIOR", ("COLCOL", "117"): "CULTURAS ARTES Y SABERES"}


def test_asignar_sector_aplica_las_reglas_en_orden():
    df = pd.DataFrame({
        "SECTORES GOB": ["Cultura", "Ambiente y Desarrollo Sostenible", "Igualdad y Equidad", "Rama Judicial", "Otro"],
        "CODIGO INTERVENCION": ["X", "Y", "R-1", "Z", "W"],
    })
    s = L.asignar_sector(df, "SECTORES GOB", "AOD", SECTORES, "CODIGO INTERVENCION", EQUIV, REGISTROS)
    assert s.tolist() == ["CULTURAS ARTES Y SABERES",          # equivalencia
                          "AMBIENTE Y DESARROLLO SOSTENIBLE",  # coincide exacto con un sector
                          "INTERIOR",                          # "Por proyecto" -> REGISTROS
                          None,                                # "Sin sector"
                          None]                                # valor desconocido sin registro


def test_registros_completan_los_que_no_traen_sector():
    """ColCol 117-119 no tienen sector vinculado en la fuente."""
    cc = pd.DataFrame({"CODIGO": [117, 105], "SECTOR VINCULADO": ["nan", "Interior"]})
    s = L.asignar_sector(cc, "SECTOR VINCULADO", "COLCOL", SECTORES, "CODIGO", EQUIV, REGISTROS)
    assert s.tolist() == ["CULTURAS ARTES Y SABERES", "INTERIOR"]


def test_unificar_colcol_toma_el_sector_por_codigo():
    xlsm = pd.DataFrame({"CODIGO": [105, 105, 117], "ETAPA": ["Apertura", "Cierre", "Apertura"]})
    sect = pd.DataFrame({"CODIGO": [105], "SECTOR VINCULADO": ["Ambiente y Desarrollo Sostenible"]})
    out = L.unificar_colcol(xlsm, sect)
    assert out["SECTOR VINCULADO"].tolist()[:2] == ["Ambiente y Desarrollo Sostenible"] * 2
    assert pd.isna(out["SECTOR VINCULADO"].iloc[2])


# ---------- Información general ----------
def test_planes_sncic():
    info = pd.DataFrame({
        "Divipola": [8, 17, 91, 50],
        "¿Cuenta con plan de trabajo en el marco del SNCIC?": ["sí", "no", "sí", "sí"],
        "Capítulo especial (Plan de trabajo con ciudades dentro del SNCIC)": ["No"] * 4,
        "Cuál": ["n/a"] * 4,
        "Año creación del plan de trabajo": ["2025", "NA", "2024", "2025"],
    })
    planes = pd.DataFrame({
        "Divipola": [8, 17, 50],
        "Tiene plan de trabajo": ["Sí", "No", "Sí"],
        "Año de inicio del plan": [None, None, None],
        "Capítulo especial": ["Sí", "Sí", "No"],
        "Ciudades con capítulo especial": ["Barranquilla", "Manizales", None],
        "Año de inicio del capítulo": [2025, 2025, None],
    }).set_index("Divipola")
    out = L.aplicar_planes_sncic(info, planes).set_index("Divipola")
    plan, cual, anio = out.columns[0], "Cuál", out.columns[3]
    assert out.loc[8, plan] == "Sí" and out.loc[8, cual] == "Barranquilla" and out.loc[8, anio] == "2025"
    assert out.loc[17, plan] == "No" and out.loc[17, cual] == "Manizales"      # capítulo sin plan
    assert out.loc[91, plan] == "No" and out.loc[91, anio] == "No aplica"      # ya no está en la app SNCIC
    assert out.loc[50, anio] == "2025"                                          # sin fecha: conserva el año


def test_registro_publico_no_muestra_divipola_ni_columnas_internas():
    fila = pd.Series({"Divipola": 8, "Departamento": "ATLÁNTICO", "DEPT_NORM": "ATLANTICO", "Capital": "Barranquilla"})
    assert [k for k, _ in L.registro_publico(fila)] == ["Departamento", "Capital"]


def test_columnas_no_publicas_definidas():
    assert {L.norm_text(c) for c in L.COLUMNAS_NO_PUBLICAS} == {"NOMBRE ENLACE COOPERACION", "CORREO DEL ENLACE DE COOPERACION"}


# ---------- Conteos y comparativos ----------
def test_top_comparativo_usa_los_mismos_grupos_del_corte_actual():
    act = pd.DataFrame({"NOMBRE ACTOR": ["A", "B"], "VALOR APORTE (USD)": [300.0, 100.0]})
    ant = pd.DataFrame({"NOMBRE ACTOR": ["A", "C"], "VALOR APORTE (USD)": [200.0, 999.0]})
    t = L.top_comparativo(L.top_by_sum(act, "NOMBRE ACTOR", "VALOR APORTE (USD)"), ant, "NOMBRE ACTOR", "Act", "Ant")
    assert t["Ant"].tolist() == ["USD 200", "USD 0"]       # B no tenía recursos: USD 0, no "-"
    assert t["Variación"].tolist() == ["▲ 50%", "Nuevo"]


def test_conteos():
    df = pd.DataFrame({"MUNICIPIO": ["Leticia", "No Reporta", "LETICIA", "Puerto Nariño"],
                       "DEPARTAMENTO": ["Amazonas", "Amazonas", "Amazonas", L.AMBITO_NACIONAL]})
    assert L.n_municipios(df) == 2
    assert L.n_departamentos(df) == 1
    assert L.unir(df["MUNICIPIO"]) == "LETICIA; Leticia; Puerto Nariño"


def test_contrapartidas_unicas():
    df = pd.DataFrame({"Departamento": ["CAUCA", "CHOCÓ"], "Nombre del proyecto": ["P"] * 2,
                       "Implementador": ["I"] * 2, "Financiador": ["F"] * 2, "AÑO": [2025] * 2})
    assert len(L.proyectos_contrapartida(df)) == 1


def test_validar_columnas_detecta_faltantes():
    df = pd.DataFrame({"CODIGO": [1]})
    problemas = L.validar_columnas({"colcol_sectores": ("Hoja COLCOL", df)})
    assert problemas and "SECTOR VINCULADO" in problemas[0]
    ok = L.validar_columnas({"colcol_sectores": ("Hoja COLCOL", pd.DataFrame({"codigo": [1], "Sector  vinculado": ["x"]}))})
    assert ok == []
