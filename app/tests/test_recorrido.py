"""Guided tour: every step points at an element that is marked in the screen, and it is shown only once."""
import re
from pathlib import Path

from appfacturas import recorrido

UI = (Path(recorrido.__file__).parent / "ui.py").read_text(encoding="utf-8")
MARCAS = set(re.findall(r"data-tour=([\w-]+)", UI))


def test_cada_paso_apunta_a_una_marca():
    for p in recorrido.PASOS:
        assert p["titulo"] and p["texto"]
        for clave in ("sel", "tab"):
            if p.get(clave):
                assert p[clave] in MARCAS, f"el paso '{p['titulo']}' apunta a {p[clave]}, que no está en ui.py"


def test_los_pasos_que_pueden_faltar_lo_explican():
    # elements that only exist with a batch read or files waiting
    for p in recorrido.PASOS:
        if p.get("sel") in ("para-leer", "avisos", "grilla", "confirmar"):
            assert p.get("falta"), p["titulo"]


def test_se_muestra_una_vez(tmp_path):
    assert not recorrido.visto(tmp_path)
    recorrido.marcar_visto(tmp_path, terminado=False)
    assert recorrido.visto(tmp_path)
    assert "ezRecorrido" in recorrido.js() and "Bienvenido" in recorrido.js()
