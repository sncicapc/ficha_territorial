"""Extrae de _SNCIC_App.xlsx (hoja "espacios") los planes de trabajo territoriales del SNCIC
y genera app/planes_trabajo_sncic.xlsx, que la ficha territorial usa para la información general.

Solo se copian campos públicos: no se copian responsables de APC, usuarios ni las demás hojas del archivo.

Uso (desde la carpeta "Versión 2.0"):
    python scripts/actualizar_planes_sncic.py "../_SNCIC_App.xlsx"
"""
import sys
from pathlib import Path

import pandas as pd

ORIGEN = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(__file__).resolve().parents[2] / "_SNCIC_App.xlsx"
RAIZ = Path(__file__).resolve().parents[1]
# Sirve en la carpeta de trabajo (Versión 2.0/app) y en el repositorio (la app está en la raíz)
APP = RAIZ / "app" if (RAIZ / "app" / "app.py").exists() else RAIZ
DESTINO = APP / "planes_trabajo_sncic.xlsx"


def nombre_propio(texto):
    t = str(texto).strip().title()
    for palabra in (" Y ", " De ", " Del ", " La ", " Las ", " Los "):
        t = t.replace(palabra, palabra.lower())
    return t


def anio(valor):
    """Año de la fecha de inicio; None si falta o es el valor de relleno 01/01/1970."""
    fecha = pd.to_datetime(valor, errors="coerce", dayfirst=True)
    if pd.isna(fecha) or fecha.year < 2000:
        return None
    return int(fecha.year)


esp = pd.read_excel(ORIGEN, sheet_name="espacios")
esp = esp[(esp["tipo_espacio"].astype(str).str.strip() == "Territorial")
          & (esp["estado"].astype(str).str.strip() == "Activo")].copy()
esp["codigo"] = esp["codigo_espacio"].astype(str).str.strip()
esp["divipola"] = esp["codigo"].str.replace("CE", "", regex=False).astype(int)
esp["es_capitulo"] = esp["codigo"].str.endswith("CE")
esp["anio"] = esp["fecha_inicio"].map(anio)

filas = []
for divipola, grupo in esp.groupby("divipola"):
    plan = grupo[~grupo["es_capitulo"]]
    capitulos = grupo[grupo["es_capitulo"]]
    filas.append({
        "Divipola": divipola,
        "Tiene plan de trabajo": "Sí" if not plan.empty else "No",
        "Año de inicio del plan": plan["anio"].dropna().min() if not plan.empty else None,
        "Capítulo especial": "Sí" if not capitulos.empty else "No",
        "Ciudades con capítulo especial": "; ".join(capitulos["nombre_espacio"].map(nombre_propio)),
        "Año de inicio del capítulo": capitulos["anio"].dropna().min() if not capitulos.empty else None,
    })

salida = pd.DataFrame(filas).sort_values("Divipola")
salida.to_excel(DESTINO, index=False, sheet_name="PLANES")
print(f"{len(salida)} departamentos con plan o capítulo especial -> {DESTINO}")
