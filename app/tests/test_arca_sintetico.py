"""ARCA layout, synthetic: products invoice without "Período Facturado" (C7) and "No Gravado" (C5)."""
from decimal import Decimal as D

import pymupdf

from appfacturas.config import Config
from appfacturas.parsers.arca import PARSER
from appfacturas.pdf import leer
from test_lote_sintetico import CUIT, _fmt

RUA = "30717419428"


def arca(ruta, periodo=False, no_gravado=None, dy=0):
    """`dy` moves the header down (a taller logo): the old fixed cut at y=175 then missed the issuer."""
    doc = pymupdf.open()
    pg = doc.new_page(width=595, height=842)

    def t(x, y, s, size=8):
        pg.insert_text((x, y + (dy if y < 600 else 0)), s, fontsize=size)
    t(250, 40, "ORIGINAL", 10)
    t(280, 70, "A", 24); t(275, 85, "COD. 01")
    t(350, 70, "FACTURA", 14)
    t(30, 90, "GRUAS INVENTADAS DEL PLATA", 11)
    t(350, 100, "Punto de Venta:"); t(420, 100, "00003"); t(460, 100, "Comp. Nro:"); t(510, 100, "00000071")
    t(350, 115, "Fecha de Emisión:"); t(430, 115, "15/10/2026")
    t(30, 115, "Razón Social:"); t(90, 115, "GRUAS INVENTADAS SRL")
    t(350, 130, "CUIT:"); t(380, 130, CUIT)
    t(30, 130, "Domicilio Comercial:"); t(120, 130, "Calle Falsa 123")
    t(30, 145, "Condición frente al IVA:"); t(140, 145, "IVA Responsable Inscripto")
    y = 175
    if periodo:
        t(30, y, "Período Facturado Desde:"); t(140, y, "01/10/2026")
        t(350, y, "Fecha de Vto. para el pago:"); t(470, y, "30/11/2026")
        y += 20
    t(30, y + 10, "CUIT:"); t(60, y + 10, RUA)
    t(200, y + 10, "Apellido y Nombre / Razón Social:"); t(340, y + 10, "RUA ASISTENCIA S.A.")
    t(30, y + 25, "Condición frente al IVA:"); t(140, y + 25, "IVA Responsable Inscripto")
    yi = y + 60
    t(30, yi, "Código"); t(80, yi, "Producto / Servicio"); t(300, yi, "Cantidad"); t(360, yi, "Precio Unit.")
    t(440, yi, "Subtotal"); t(480, yi, "Alicuota IVA"); t(530, yi, "Subtotal c/IVA")
    t(30, yi + 15, "P1"); t(80, yi + 15, "Cubierta 175/70"); t(300, yi + 15, "1,00")
    t(360, yi + 15, "100000,00"); t(440, yi + 15, "100000,00"); t(490, yi + 15, "21%"); t(530, yi + 15, "121000,00")
    total = D("121000.00") + (no_gravado or 0)
    yt = 650
    t(380, yt, "Importe Neto Gravado: $"); t(520, yt, "100000,00")
    if no_gravado:
        t(380, yt + 12, "Importe Neto No Gravado: $"); t(520, yt + 12, f"{no_gravado:.2f}".replace(".", ","))
    t(380, yt + 24, "IVA 21%: $"); t(520, yt + 24, "21000,00")
    t(380, yt + 36, "Importe Total: $"); t(520, yt + 36, f"{total:.2f}".replace(".", ","))
    t(30, 780, "CAE N°: 76123456789012")
    doc.save(ruta)


def test_arca_de_productos_sin_periodo(tmp_path):
    ruta = tmp_path / "productos.pdf"
    arca(ruta, dy=80)
    doc = leer(ruta)
    assert PARSER.detectar(doc, ruta.name) > 0
    e = PARSER.extraer(doc, ruta.name, Config())
    assert (e.cuit, e.cuit_receptor) == (CUIT, RUA)
    assert (e.letra, e.pv, e.nro) == ("A", 3, 71)
    assert (e.neto, e.iva, e.total) == (D("100000.00"), D("21000.00"), D("121000.00"))


def test_arca_no_gravado(tmp_path):
    ruta = tmp_path / "ng.pdf"
    arca(ruta, periodo=True, no_gravado=D("5000"))
    e = PARSER.extraer(leer(ruta), ruta.name, Config())
    assert (e.neto, e.no_gravado, e.total) == (D("100000.00"), D("5000.00"), D("126000.00"))
    assert e.cuit == CUIT and _fmt(CUIT)
