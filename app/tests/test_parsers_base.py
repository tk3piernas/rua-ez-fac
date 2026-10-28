"""Task numbers vs prices (C4) and receipt / sale-note types (C6)."""
import pytest

from appfacturas.carpetas import Indice, clave
from appfacturas.parsers.base import tareas_en, tipo_de_codigo


@pytest.mark.parametrize("texto,esperado", [
    ("Auxilio mecánico tarea 412345", ["412345"]),
    ("Precio unitario 350000", []),
    ("Precio unitario: $ 350000", []),
    ("Importe 412345", []),
    ("$412345", []),
    ("Importe tarea 412345", ["412345"]),
    ("Exp. 412345 / 412346 Total 350000", ["412345", "412346"]),
])
def test_tarea_no_es_un_precio(texto, esperado):
    assert tareas_en(texto, 300000, 999999) == esperado


@pytest.mark.parametrize("cod,esperado", [
    (1, ("A", "Factura")), (4, ("A", "Recibo")), (9, ("B", "Recibo")), (15, ("C", "Recibo")),
    (5, ("A", "Nota de Venta")), (10, ("B", "Nota de Venta")), (3, ("A", "Nota de Crédito")),
    (7, ("B", "Nota de Débito")), (201, ("A", "Factura")), (203, ("A", "Nota de Crédito")),
])
def test_tipo_de_codigo(cod, esperado):
    assert tipo_de_codigo(cod) == esperado


def test_recibo_no_choca_con_la_factura_del_mismo_numero(tmp_path):
    d = tmp_path / "Vehicular" / "Prov Inventado"
    d.mkdir(parents=True)
    (d / "Factura 0001-00000017.pdf").write_bytes(b"f")
    (d / "Recibo 0001-00000017.pdf").write_bytes(b"r")
    ix = Indice(tmp_path, ["Vehicular"], {"prov inventado": "30700001118"})
    assert clave("30700001118", "Factura", 1, 17) in ix.por_clave
    assert clave("30700001118", "Recibo", 1, 17) in ix.por_clave
    assert len(ix.por_clave) == 2
