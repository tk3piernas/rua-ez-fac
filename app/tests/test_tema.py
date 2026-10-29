"""Colours of the app (Configuración → Ventana y colores): saved in config.toml, claro by default."""
from appfacturas import config


def test_tema_por_defecto_y_guardado(tmp_path):
    archivo = tmp_path / "config.toml"
    assert config.cargar(archivo).tema == "claro"                  # a new config.toml starts light
    config.guardar_opcion(archivo, "app", "tema", "oscuro")
    assert config.cargar(archivo).tema == "oscuro"
    assert config.cargar(archivo).modo_real is False               # the other options stay as they were
    config.guardar_opcion(archivo, "app", "tema", "violeta")       # anything else falls back to light
    assert config.cargar(archivo).tema == "claro"
