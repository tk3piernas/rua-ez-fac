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
