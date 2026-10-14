"""Prueba de humo: la app abre todas las páginas, departamentos y sectores sin errores."""
from pathlib import Path

import pytest
from streamlit.testing.v1 import AppTest

APP_PY = str(Path(__file__).resolve().parents[1] / "app.py")

PANORAMA, TERRITORIAL, SECTORIAL, GUIA = (
    "\U0001f310 Panorama Nacional", "\U0001f5fa️ Ficha Territorial",
    "\U0001f3db️ Ficha Sectorial", "\U0001f4d6 Guía de usuario")


@pytest.fixture(scope="module")
def app():
    at = AppTest.from_file(APP_PY, default_timeout=900)
    at.run()
    assert not at.exception, at.exception
    return at


def errores(at):
    return [str(e.value) for e in at.exception]


def test_panorama_y_capas_del_mapa(app):
    app.radio(key="main_nav").set_value(PANORAMA).run()
    assert [t.label for t in app.tabs] == ["AOD", "Sur-Sur", "Programas APC"]
    for capa in app.radio(key="capa_mapa").options:
        app.radio(key="capa_mapa").set_value(capa).run()
        assert errores(app) == [], capa


def test_guia(app):
    app.radio(key="main_nav").set_value(GUIA).run()
    assert errores(app) == []


def test_todos_los_departamentos(app):
    app.radio(key="main_nav").set_value(TERRITORIAL).run()
    for d in app.selectbox(key="dept_sel").options:
        app.selectbox(key="dept_sel").set_value(d).run()
        assert errores(app) == [], d
        assert len(app.get("download_button")) == 2, d
        registro = [x.value for x in app.dataframe if list(x.value.columns) == ["Campo", "Valor"]]
        assert registro and "Divipola" not in set(registro[0]["Campo"]), d


def test_todos_los_sectores(app):
    app.radio(key="main_nav").set_value(SECTORIAL).run()
    opciones = app.selectbox(key="sector_sel").options
    assert not any("IGUALDAD" in o.upper() for o in opciones)
    for s in opciones:
        app.selectbox(key="sector_sel").set_value(s).run()
        assert errores(app) == [], s
        assert len(app.get("download_button")) == 2, s
