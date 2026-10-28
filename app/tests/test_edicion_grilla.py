"""E5: editing the grid. Provider name (new: renamed everywhere; registered: not) and amount (column H)."""
from decimal import Decimal as D
from pathlib import Path
from urllib.parse import quote

from appfacturas import ui
from test_lote_sintetico import CUIT, armar, factura, sandbox  # noqa: F401  (fixture)


def _editar(lote, i, col, v):
    return ui.aplicar_edicion(lote, {"colId": col, "data": {"idx": i, col: v}})


def test_proveedor_nuevo_se_renombra_en_todo(sandbox, monkeypatch):
    monkeypatch.setattr(ui.S, "cfg", sandbox)
    factura(sandbox.entrada / "Vehicular" / "uno.pdf", 91)
    factura(sandbox.entrada / "Vehicular" / "dos.pdf", 92)
    lote = armar(sandbox)
    assert lote.facturas[0].proveedor_nuevo
    assert _editar(lote, 0, "proveedor", "Grúas del Plata") is None
    assert {r.proveedor for r in lote.filas} == {"Grúas del Plata"}               # both invoices
    (alta,) = [p for p in lote.plan if p["op"] == "proveedor_nuevo"]
    assert alta["nombre"] == "Grúas del Plata"
    movs = [p["a"] for p in lote.plan if p["op"] == "mover"]
    assert len(movs) == 2 and all(Path(a).parent.name == alta["carpeta"] for a in movs)
    assert all("/" + quote(alta["carpeta"]) + "/" in r.url for r in lote.filas)


def test_proveedor_registrado_no_se_renombra(sandbox, monkeypatch):
    monkeypatch.setattr(ui.S, "cfg", sandbox)
    (sandbox.base / "_Script" / "proveedores.csv").write_text(
        f"CUIT;Area;Proveedor;Carpeta\r\n{CUIT};Vehicular;Gruas Inventadas;Gruas Inventadas\r\n", encoding="utf-8-sig")
    factura(sandbox.entrada / "Vehicular" / "uno.pdf", 93)
    lote = armar(sandbox)
    assert "ya está registrado" in _editar(lote, 0, "proveedor", "Otro Nombre")
    assert lote.filas[0].proveedor == "Gruas Inventadas"


def test_editar_monto_recalcula_h(sandbox, monkeypatch):
    monkeypatch.setattr(ui.S, "cfg", sandbox)
    factura(sandbox.entrada / "Vehicular" / "a.pdf", 94, neto="100.000,00", iva="10.500,00", total="110.500,00")
    factura(sandbox.entrada / "Vehicular" / "b.pdf", 95, letra="B", neto=None, iva=None, total="121.000,00")
    lote = armar(sandbox)
    i_a = next(i for i, r in enumerate(lote.filas) if r.iva == D("10500.00"))
    i_b = next(i for i, r in enumerate(lote.filas) if r.iva == 0)
    assert "fórmula" in _editar(lote, i_a, "monto", "90.000,00")
    assert (lote.filas[i_a].monto, lote.filas[i_a].iva) == (D("90000.00"), None)
    assert _editar(lote, i_b, "monto", "100.000,00") is None
    assert lote.filas[i_b].iva == 0
