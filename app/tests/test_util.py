from decimal import Decimal as D

import pytest

from appfacturas.util import cbu_valido, cuit_valido, norm, parse_monto, url_sharepoint


@pytest.mark.parametrize("txt,esperado", [
    ("130000,00", D("130000.00")), ("1.234.567,89", D("1234567.89")), ("$95,000.00", D("95000.00")),
    ("$25.600,00", D("25600.00")), ("1,5", D("1.5")), ("130.000", D("130000")), ("-260000,00", D("-260000.00")),
    ("0,00", D("0.00")), ("", None), ("abc", None), ("21%", None),
])
def test_parse_monto(txt, esperado):
    assert parse_monto(txt) == esperado


def test_cuit():
    assert cuit_valido("20-11111111-2") and cuit_valido("23123456785") and cuit_valido("30717419428")
    assert not cuit_valido("20010203047") and not cuit_valido("27050607089") and not cuit_valido("123")


def test_cbu():
    for c in ("0110000612345678901233", "0000003110000000000014", "0720000798765432109877", "3840000955555555555555"):
        assert cbu_valido(c)
    assert not cbu_valido("0110000612345678901234")


def test_url_como_el_ps1():
    u = url_sharepoint("https://x/Documents/Facturas pendientes", "Hogar", "Sebastián Cabral", "Factura 0001-00000461.pdf")
    assert u == "https://x/Documents/Facturas%20pendientes/Hogar/Sebasti%C3%A1n%20Cabral/Factura%200001-00000461.pdf"


def test_norm():
    assert norm("Roldán  Juan-Ramón") == "roldan juan ramon"
