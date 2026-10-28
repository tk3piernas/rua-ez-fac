"""A2: confirming the same batch twice (double click) never runs it again."""
import json
import threading

import pytest

from appfacturas import ejecutor
from appfacturas.config import Config
from appfacturas.model import Lote


def _ejecutor(tmp_path):
    cfg = Config(base=tmp_path / "Facturas pendientes", planilla=tmp_path / "DEUDA A PAGAR ARG.xlsx",
                 excel_motor="openpyxl")
    carpeta = tmp_path / "lotes" / "2026-01-01_000000"
    carpeta.mkdir(parents=True)
    return ejecutor.Ejecutor(cfg, Lote("x"), carpeta, log=lambda s: None, con_excel=False), carpeta


@pytest.mark.parametrize("ya", ["en curso", "confirmado", "confirmado con errores", "deshecho"])
def test_no_confirma_dos_veces(tmp_path, ya):
    ej, carpeta = _ejecutor(tmp_path)
    (carpeta / "estado.json").write_text(json.dumps({"estado": ya}), encoding="utf-8")
    with pytest.raises(ejecutor.YaConfirmado):
        ej.ejecutar()
    assert json.loads((carpeta / "estado.json").read_text(encoding="utf-8")) == {"estado": ya}
    assert not (carpeta / "diario.json").exists()


def test_segunda_confirmacion_seguida(tmp_path):
    ej, carpeta = _ejecutor(tmp_path)
    ej.ejecutar()
    assert ejecutor.estado(carpeta)["estado"] == "confirmado"
    with pytest.raises(ejecutor.YaConfirmado):
        ejecutor.Ejecutor(ej.cfg, Lote("x"), carpeta, log=lambda s: None, con_excel=False).ejecutar()


def test_mientras_corre_se_niega(tmp_path, monkeypatch):
    ej, carpeta = _ejecutor(tmp_path)
    adentro, seguir = threading.Event(), threading.Event()

    def lento(self):
        adentro.set()
        seguir.wait(5)
        return {}
    monkeypatch.setattr(ejecutor.Ejecutor, "_ejecutar", lento)
    t = threading.Thread(target=ej.ejecutar)
    t.start()
    assert adentro.wait(5)
    try:
        with pytest.raises(ejecutor.YaConfirmado):
            ejecutor.Ejecutor(ej.cfg, Lote("x"), carpeta, log=lambda s: None, con_excel=False).ejecutar()
        with pytest.raises(ejecutor.YaDeshecho):
            ejecutor.deshacer(carpeta, log=lambda s: None, cfg=ej.cfg)
    finally:
        seguir.set()
        t.join()
