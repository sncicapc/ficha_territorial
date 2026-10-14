# Ficha de caracterización territorial y sectorial de Cooperación Internacional · Versión 2.0

Herramienta pública (Streamlit) de APC-Colombia que consolida la cooperación internacional por
departamento y por sector: AOD (Cíclope), Cooperación Sur-Sur, Colombia Enseña Colombia (ColCol)
y Contrapartidas.

## Estructura

En la carpeta de trabajo la app está en `app/`; **en el repositorio de GitHub (rama V2) el contenido de `app/`
está en la raíz**, junto a `scripts/` y este README. Las carpetas `revision/` y `respaldos/` son internas y no se suben.

```
Versión 2.0/
├── app/                          ← lo que se publica (raíz del repositorio)
│   ├── app.py                    ← presentación (páginas, pestañas, tablas, PDF y Excel)
│   ├── logica.py                 ← lectura de datos y reglas de negocio (sin Streamlit)
│   ├── configuracion.xlsx        ← cortes, fechas, sectores visibles y equivalencias (se edita aquí, no en el código)
│   ├── planes_trabajo_sncic.xlsx ← planes de trabajo territoriales (lo genera un script)
│   ├── Ficha sectores.xlsx       ← AOD de Cíclope (fuente única), información de sectores y sector vinculado de ColCol
│   ├── Ficha_territorial.xlsm    ← información general, Sur-Sur, ColCol y Contrapartidas
│   ├── tests/                    ← pruebas automáticas
│   ├── requirements.txt          ← versiones fijas para el despliegue
│   └── requirements-dev.txt      ← agrega pytest para correr las pruebas
├── scripts/
│   ├── actualizar_planes_sncic.py ← _SNCIC_App.xlsx (hoja espacios) → app/planes_trabajo_sncic.xlsx
│   └── validar_datos.py           ← reporte de calidad de datos (revision/Reporte_calidad_datos.xlsx)
├── revision/                     ← tablas de decisiones, justificaciones y reportes
└── respaldos/                    ← copias de app.py antes de cambios grandes
```

## Reglas principales

- **AOD: una sola fuente** (`Ficha sectores.xlsx`, una hoja por corte) para las dos fichas, con cada cooperante
  con su nombre individual (criterio de APC: no se agrupan actores como "Filantropía").
- **AOD: un solo universo.** Solo cuentan las intervenciones *vigentes* en la fecha de cada corte
  (fecha inicial ≤ fecha de vigencia ≤ fecha final). No se usa el campo de estado de Cíclope.
- **Comparativos** contra el corte anterior definido en `configuracion.xlsx` (hoja CORTES).
- **Sectores:** el sector de cada registro sale de la hoja EQUIVALENCIAS; los casos puntuales
  (valores "Por proyecto" o registros sin sector en la fuente) se asignan por código en REGISTROS.
  Igualdad y Equidad está oculto (ministerio liquidado, Decreto 0626 de 2026); sus registros se reubicaron por tema.
- **Datos personales:** no se publican nombre ni correo de los enlaces de cooperación
  (ver `revision/Justificacion_datos_enlaces_cooperacion.md`). Tampoco se muestra el Divipola en el registro.
- **Departamentos:** coincidencia exacta del nombre (así Santander no incluye a Norte de Santander).

## Cómo actualizar con un corte nuevo

1. **Datos de Cíclope:** pegar la nueva extracción en una hoja nueva de `Ficha sectores.xlsx` (por ejemplo, AOD202603),
   con los mismos encabezados y **sin agrupar cooperantes**. Ya no hace falta copiarla en `Ficha_territorial.xlsm`.
2. **Configuración** (`app/configuracion.xlsx`, hoja CORTES): la fila *actual* pasa a *anterior*; en la
   nueva fila *actual* escribir la etiqueta (p. ej. 2026-3), la fecha de vigencia, la fecha de extracción
   y el nombre de la hoja de AOD. Actualizar la hoja FUENTES si cambió la fecha de Sur-Sur.
3. **Planes SNCIC:** si se actualizó `_SNCIC_App.xlsx`, correr
   `python scripts/actualizar_planes_sncic.py "../_SNCIC_App.xlsx"`.
4. **Validar:** `python scripts/validar_datos.py` y revisar `revision/Reporte_calidad_datos.xlsx`.
   Si aparece "Bloquea la app", corregir antes de publicar. Los valores nuevos sin sector se resuelven
   en EQUIVALENCIAS o REGISTROS.
5. **Probar:** `python -m pytest app/tests -q` (unos 2 minutos).
6. **Revisar en local:** `cd app` y `python -m streamlit run app.py`.
7. **Publicar** (ver abajo).

## Correr en local y probar

```
cd app
python -m pip install -r requirements-dev.txt
python -m streamlit run app.py
python -m pytest tests -q
```

## Publicación

- Streamlit Community Cloud, desde un repositorio de GitHub **privado** cuya raíz sea la carpeta `app`,
  archivo principal `app.py`, Python 3.12.
- El repositorio debe ser privado: los Excel incluyen datos de contacto que la app no muestra
  pero que se podrían descargar desde un repositorio público.
- Las versiones de `requirements.txt` están fijas a las probadas; actualizarlas solo después de correr las pruebas.

## Pendientes conocidos

- Las tablas de AOD que siguen dentro de `Ficha_territorial.xlsm` (Tabla7, ciclope20261, ciclope2025) ya no se usan;
  se pueden borrar del archivo cuando se quiera aligerarlo (la app no las lee).
- Sur-Sur DOCI028 (EDUEVOLUCION URUCOL) sin sector; DOCI074 tiene como espacio vinculado "Manizales" (ciudad),
  por lo que no aparece en la ficha de Caldas.
- Planes SNCIC sin fecha de inicio en la app SNCIC (ver el reporte de calidad).
