"""Taskbar identity: the shortcut and the window carry the same AppUserModelID (pinnable with its own icon)."""
import sys

import pytest

from appfacturas import ventana

pytestmark = pytest.mark.skipif(sys.platform != "win32", reason="solo Windows")


def test_acceso_directo_con_identidad(tmp_path):
    lnk = tmp_path / "RUA EZ-Fac.lnk"
    ventana.acceso_directo(lnk, tmp_path)
    assert lnk.is_file() and ventana.id_de_acceso(lnk) == ventana.AUMID == "RUA.EZFac"
    assert ventana.comando().endswith('-m appfacturas ui') and ventana.ICONO.is_file()


def test_un_clic_con_la_app_abierta_trae_la_ventana(monkeypatch):
    """By default the icon brings the open window to the front; with [app] varias_ventanas it opens another."""
    from appfacturas import ui
    from appfacturas.config import Config
    abiertas = []
    monkeypatch.setattr(ui, "abrir_ventana", lambda url: abiertas.append(url))
    monkeypatch.setattr(ventana, "mostrar_la_abierta", lambda: True)
    ui.mostrar_app("http://x", Config())
    assert abiertas == []
    ui.mostrar_app("http://x", Config(varias_ventanas=True))
    assert abiertas == ["http://x"]
    monkeypatch.setattr(ventana, "mostrar_la_abierta", lambda: False)      # none open: opens one
    ui.mostrar_app("http://x", Config())
    assert abiertas == ["http://x", "http://x"]


def test_opcion_varias_ventanas_en_config(tmp_path):
    from appfacturas import config
    toml = tmp_path / "config.toml"
    assert config.cargar(toml).varias_ventanas is False                    # off by default (new file)
    config.guardar_opcion(toml, "app", "varias_ventanas", True)
    assert config.cargar(toml).varias_ventanas is True
