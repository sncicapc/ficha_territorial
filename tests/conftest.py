"""Configuración común de las pruebas: se ejecutan desde la carpeta de la app."""
import sys
from pathlib import Path

import pytest

APP_DIR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(APP_DIR))


@pytest.fixture(autouse=True)
def en_carpeta_app(monkeypatch):
    """Las rutas de los archivos de datos son relativas a la carpeta de la app."""
    monkeypatch.chdir(APP_DIR)
