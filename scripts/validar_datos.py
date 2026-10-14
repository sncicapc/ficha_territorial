"""Revisa la calidad de los datos de la ficha y genera un reporte para corregir en las fuentes.

Uso (desde la carpeta "Versión 2.0"):
    python scripts/validar_datos.py

Crea revision/Reporte_calidad_datos.xlsx. No modifica ningún archivo de datos.
Correrlo cada vez que llegue un corte nuevo, antes de publicar.
"""
import sys
from pathlib import Path

import pandas as pd

RAIZ = Path(__file__).resolve().parents[1]
# Sirve en la carpeta de trabajo (Versión 2.0/app) y en el repositorio (la app está en la raíz)
APP = RAIZ / "app" if (RAIZ / "app" / "app.py").exists() else RAIZ
sys.path.insert(0, str(APP))
import logica as L  # noqa: E402

hallazgos = []


def anotar(gravedad, fuente, tema, detalle, cantidad, accion):
    hallazgos.append({"Gravedad": gravedad, "Fuente": fuente, "Tema": tema, "Detalle": detalle,
                      "Cantidad": cantidad, "Qué hacer": accion})


cfg = L.cargar_config(APP / "configuracion.xlsx")
act, ant = cfg["cortes"]["actual"], cfg["cortes"]["anterior"]
x = L.cargar_xlsm(APP / "Ficha_territorial.xlsm")
s = L.cargar_sectores(APP / "Ficha sectores.xlsx", act["hoja"], ant["hoja"])

# 1. Estructura
for p in L.validar_columnas({
    "infogeneral": ("Ficha_territorial.xlsm · infogeneral", x["infogeneral"]),
    "aod": (f"Ficha sectores.xlsx · {act['hoja']}", s["aod"]),
    "aod_ant": (f"Ficha sectores.xlsx · {ant['hoja']}", s["aod_ant"]),
    "css": ("Ficha_territorial.xlsm · CSS", x["css"]),
    "colcol": ("Ficha_territorial.xlsm · colcol", x["colcol"]),
    "contrapartidas": ("Ficha_territorial.xlsm · contrapartidas", x["contrapartidas"]),
    "info_sectores": ("Ficha sectores.xlsx · INFO GENERAL", s["info"]),
    "colcol_sectores": ("Ficha sectores.xlsx · COLCOL", s["colcol"]),
}):
    anotar("Bloquea la app", "Estructura", "Columnas faltantes", p, 1,
           "Corregir los encabezados para que coincidan con los del corte anterior.")

deptos = set(x["infogeneral"]["Departamento"].map(L.norm_text)) | {L.norm_text(L.AMBITO_NACIONAL)}
sectores = {L.norm_text(v) for v in s["info"]["Nombre del sector"]} - cfg["ocultos"]
reglas = dict(sectores_norm=sectores, equivalencias=cfg["equivalencias"], registros=cfg["registros"])

# 2. AOD
aod = L.activas_a_fecha(s["aod"], act["vigencia"])
fi, ff = L.a_fecha(s["aod"]["FECHA INICIAL"]), L.a_fecha(s["aod"]["FECHA FINAL"])
sin_fecha = s["aod"].loc[fi.isna() | ff.isna(), "CODIGO INTERVENCION"].nunique()
if sin_fecha:
    anotar("Revisar", "AOD (Cíclope)", "Intervenciones sin fecha inicial o final", "Quedan fuera del universo de vigentes.",
           sin_fecha, "Completar las fechas en Cíclope o en la tabla del corte.")
sin_usd = aod.groupby("CODIGO INTERVENCION")["VALOR APORTE (USD)"].sum()
if (sin_usd == 0).any():
    anotar("Revisar", "AOD (Cíclope)", "Intervenciones vigentes sin aporte en USD",
           ", ".join(sin_usd[sin_usd == 0].index[:15]), int((sin_usd == 0).sum()),
           "Verificar si el valor está pendiente de registro.")
otros = sorted(set(aod["DEPARTAMENTO"].map(L.norm_text)) - deptos)
if otros:
    anotar("Revisar", "AOD (Cíclope)", "Departamentos no reconocidos", ", ".join(otros), len(otros),
           "No aparecerán en ninguna ficha territorial: corregir el nombre.")
aod_s = aod.copy()
aod_s["S"] = L.asignar_sector(aod_s, "SECTORES GOB", "AOD", codigo_col="CODIGO INTERVENCION", **reglas)
sin_sector = aod_s[aod_s["S"].isna()]
if not sin_sector.empty:
    valores = sin_sector.groupby("SECTORES GOB")["CODIGO INTERVENCION"].nunique().sort_values(ascending=False)
    anotar("Informativo", "AOD (Cíclope)", "Intervenciones vigentes sin sector de la ficha",
           "; ".join(f"{k} ({v})" for k, v in valores.items()), int(sin_sector["CODIGO INTERVENCION"].nunique()),
           "Si deben ir a un sector, agregar el valor en EQUIVALENCIAS o el código en REGISTROS (configuracion.xlsx).")
agrupados = aod[aod["NOMBRE ACTOR"].map(L.norm_text).isin({"FILANTROPIA", "BID FCS"})]
if not agrupados.empty:
    anotar("Revisar", "AOD (Cíclope)", "Cooperantes agrupados",
           "El criterio de APC es mostrar cada cooperante con su nombre, no agrupado como Filantropía o BID (FCS).",
           int(agrupados["CODIGO INTERVENCION"].nunique()), "Pegar la extracción de Cíclope sin agrupar los actores.")

# 3. Sur-Sur
css = x["css"].copy()
css["S"] = L.asignar_sector(css, "ESPACIO VINCULADO", "CSS", codigo_col="Código", **reglas)
css["T"] = css["ESPACIO VINCULADO"].map(L.norm_text).isin(deptos)
por_proyecto = css.groupby("Código").agg(sector=("S", lambda v: v.notna().any()), territorio=("T", "any"))
huerfanos = por_proyecto[~por_proyecto["sector"] & ~por_proyecto["territorio"]].index.tolist()
if huerfanos:
    anotar("Revisar", "Sur-Sur", "Proyectos que no aparecen en ninguna ficha", ", ".join(huerfanos), len(huerfanos),
           "Asignar sector en REGISTROS o corregir el espacio vinculado en la fuente.")

# 4. ColCol
cc = L.unificar_colcol(x["colcol"], s["colcol"])
cc["S"] = L.asignar_sector(cc, "SECTOR VINCULADO", "COLCOL", codigo_col="CODIGO", **reglas)
sin = sorted(cc.loc[cc["S"].isna(), "CODIGO"].astype(str).unique())
if sin:
    anotar("Revisar", "ColCol", "Intercambios sin sector", ", ".join(sin), len(sin),
           "Asignar el sector en la hoja REGISTROS de configuracion.xlsx (Fuente = ColCol).")
nuevos = sorted(set(x["colcol"]["CODIGO"].astype(str)) - set(s["colcol"]["CODIGO"].astype(str)))
if nuevos:
    anotar("Informativo", "ColCol", "Intercambios que no están en Ficha sectores.xlsx", ", ".join(nuevos), len(nuevos),
           "Su sector se toma de REGISTROS; si se agregan a Ficha sectores, se usará el sector vinculado de allá.")
for fuente, serie in (("ColCol", x["colcol"]["DEPARTAMENTOS PARTICIPANTES"]),
                      ("Contrapartidas", x["contrapartidas"]["Departamento"])):
    otros = sorted(set(serie.map(L.norm_text)) - deptos - L.SIN_DATO)
    if otros:
        anotar("Revisar", fuente, "Departamentos no reconocidos", ", ".join(otros), len(otros),
               "No aparecerán en la ficha territorial: corregir el nombre.")

# 5. Planes SNCIC
planes = L.leer_planes(APP / "planes_trabajo_sncic.xlsx")
if planes is None:
    anotar("Revisar", "Planes SNCIC", "Falta planes_trabajo_sncic.xlsx", "", 1,
           "Correr scripts/actualizar_planes_sncic.py con el archivo _SNCIC_App.xlsx.")
else:
    con_plan = planes[planes["Tiene plan de trabajo"] == "Sí"]
    sin_anio = con_plan[con_plan["Año de inicio del plan"].isna()]
    if not sin_anio.empty:
        nombres = x["infogeneral"].assign(D=pd.to_numeric(x["infogeneral"]["Divipola"]).astype(int)).set_index("D")["Departamento"]
        anotar("Informativo", "Planes SNCIC", "Planes sin fecha de inicio en la app SNCIC",
               ", ".join(nombres.reindex(sin_anio.index).fillna("?")), len(sin_anio),
               "Completar la fecha de inicio en la app SNCIC (hoja espacios); mientras tanto se muestra el año anterior de la ficha.")

# Reporte
orden = {"Bloquea la app": 0, "Revisar": 1, "Informativo": 2}
reporte = pd.DataFrame(hallazgos, columns=["Gravedad", "Fuente", "Tema", "Detalle", "Cantidad", "Qué hacer"])
reporte = reporte.sort_values("Gravedad", key=lambda c: c.map(orden))
(RAIZ / "revision").mkdir(exist_ok=True)
salida = RAIZ / "revision" / "Reporte_calidad_datos.xlsx"
with pd.ExcelWriter(salida, engine="openpyxl") as w:
    reporte.to_excel(w, index=False, sheet_name="Hallazgos")
    hoja = w.sheets["Hallazgos"]
    for col, ancho in zip("ABCDEF", (16, 18, 40, 80, 10, 70)):
        hoja.column_dimensions[col].width = ancho

print(f"Corte {act['etiqueta']} (vigentes al {L.fecha_larga(act['vigencia'])}): {len(reporte)} hallazgos")
for _, r in reporte.iterrows():
    print(f"  [{r['Gravedad']}] {r['Fuente']} · {r['Tema']}: {r['Cantidad']}")
print(f"Reporte: {salida}")
sys.exit(1 if (reporte["Gravedad"] == "Bloquea la app").any() else 0)
