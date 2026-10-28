"""Batch building with synthetic invoices (invented provider, invented CUIT with a valid check digit)."""
import datetime
from decimal import Decimal as D

import pymupdf
import pytest
from openpyxl import load_workbook

from appfacturas import preparar
from appfacturas.config import Config
from appfacturas.lote import Armador, entradas_de_carpeta
from appfacturas.model import Nivel
from appfacturas.util import cuit_valido

HOY = datetime.date(2026, 10, 27)


def _cuit(base10: str) -> str:
    return next(base10 + str(d) for d in range(10) if cuit_valido(base10 + str(d)))


CUIT = _cuit("3070000111")


def _fmt(c):
    return f"{c[:2]}-{c[2:10]}-{c[10]}"


@pytest.fixture
def sandbox(tmp_path):
    base = tmp_path / "OneDrive - Empresa" / "Facturas pendientes"
    for a in ("Vehicular", "Hogar", "Persona", "ServiciosYa"):
        (base / "_Entrada" / a).mkdir(parents=True)
    (base / "_Script").mkdir()
    (base / "_Script" / "proveedores.csv").write_text("CUIT;Area;Proveedor;Carpeta\r\n", encoding="utf-8-sig")
    cfg = Config(base=base, planilla=base.parent / "DEUDA A PAGAR ARG.xlsx", excel_motor="openpyxl",
                 url_base="https://empresa.example/Documents/Facturas pendientes")
    preparar.crear_planilla_vacia(cfg.planilla)
    return cfg


def factura(ruta, nro, letra="A", neto="100.000,00", iva="21.000,00", total="121.000,00", items=None,
            vto_pago=True, extras=(), cabecera=""):
    """Synthetic invoice for the generic reader. Lines with None are left out."""
    doc = pymupdf.open()
    pg = doc.new_page(width=595, height=842)

    def t(x, y, s, size=9):
        pg.insert_text((x, y), s, fontsize=size)
    t(30, 40, "GRUAS INVENTADAS DEL PLATA", 13)
    if cabecera:
        t(30, 22, cabecera)
    t(290, 45, letra, 24)
    t(285, 60, "Cod. 01" if letra == "A" else "Cod. 11", 8)
    t(360, 40, "FACTURA", 14)
    t(360, 60, f"Nº: 0003-{nro:08d}")
    t(360, 74, "Fecha: 15/10/2026")
    t(360, 88, f"CUIT: {_fmt(CUIT)}")
    t(30, 120, "Cliente: RUA ASISTENCIA S.A.   CUIT: 30-71741942-8   I.V.A.: Responsable Inscripto")
    t(30, 160, "Código"); t(80, 160, "Descripción"); t(400, 160, "Cantidad"); t(520, 160, "Importe")
    for k, (desc, monto) in enumerate(items or [("Auxilio mecánico tarea 412345", neto or total)]):
        t(30, 180 + 15 * k, "AUX"); t(80, 180 + 15 * k, desc); t(410, 180 + 15 * k, "1")
        if monto:
            t(520, 180 + 15 * k, monto)
    if neto:
        t(380, 600, "Neto:"); t(520, 600, neto)
    if iva:
        t(380, 615, "IVA 21%:"); t(520, 615, iva)
    t(380, 630, "Total:"); t(520, 630, total)
    for k, (label, monto) in enumerate(extras):
        t(380, 645 + 12 * k, label); t(520, 645 + 12 * k, monto)
    if vto_pago:
        t(30, 660, "Fecha de venc. de pago: 30/11/2026")
    t(30, 700, "CAE: 76123456789012")
    t(30, 715, "Fecha de Vencimiento CAE: 25/10/2026")
    doc.save(ruta)


def armar(cfg):
    ents, txts = entradas_de_carpeta(cfg)
    return Armador(cfg, simular=True, hoy=HOY).procesar(ents, txts)


def test_hoja_faltante_no_frena_el_resto(sandbox):
    """A4: an invoice of an area without its sheet is shown but not loaded; the others still are."""
    wb = load_workbook(sandbox.planilla)
    del wb["Persona"]
    wb.save(sandbox.planilla)
    factura(sandbox.entrada / "Persona" / "p.pdf", 11)
    factura(sandbox.entrada / "Vehicular" / "v.pdf", 12)
    lote = armar(sandbox)
    assert any(a.codigo == "sin_hoja" and a.nivel == Nivel.BLOQUEA for a in lote.avisos)
    por_hoja = {r.hoja: r for r in lote.filas}
    assert por_hoja["Persona"].excluida and por_hoja["Persona"].bloqueada
    assert not por_hoja["Vehicular"].excluida
    from appfacturas import excel
    from types import SimpleNamespace as NS
    res = excel.escribir(sandbox, lote, NS(anotar=lambda p: None), sandbox.base.parent / "backups", log=lambda s: None)
    assert res["filas"] == 1


def test_factura_a_sin_neto_no_se_carga(sandbox):
    """A5: Factura A without a net amount would get G = total and H = formula (21 % paid twice)."""
    factura(sandbox.entrada / "Vehicular" / "a.pdf", 21, neto=None, iva=None, total="121.000,00")
    lote = armar(sandbox)
    assert [a.codigo for a in lote.avisos if a.nivel == Nivel.BLOQUEA] == ["a_sin_neto"]
    assert all(r.excluida for r in lote.filas)
    assert not [p for p in lote.plan if p["op"] == "mover"]


def test_factura_a_exenta_h_en_cero(sandbox):
    """A5: IVA 0 read and total = net: H is 0, not the 21 % formula."""
    factura(sandbox.entrada / "Vehicular" / "e.pdf", 22, neto="100.000,00", iva="0,00", total="100.000,00")
    lote = armar(sandbox)
    (r,) = lote.filas
    assert (r.monto, r.iva) == (D("100000.00"), D("0"))
    assert not r.excluida


def test_factura_a_normal_sigue_con_formula(sandbox):
    factura(sandbox.entrada / "Vehicular" / "n.pdf", 23)
    (r,) = armar(sandbox).filas
    assert (r.monto, r.iva) == (D("100000.00"), None)


def test_item_sin_monto_en_factura_a(sandbox, monkeypatch):
    """C8: an item without amount in a Factura A with several items does not break the reading."""
    from appfacturas import lote as lote_mod
    elegir = lote_mod.parsers.elegir

    def elegir_con_item_vacio(doc, nombre):
        parser, puntaje = elegir(doc, nombre)

        class SinMonto:
            def extraer(self, *a):
                e = parser.extraer(*a)
                e.items[1].monto = None          # some readers return an item without its amount
                return e
        return SinMonto(), puntaje
    monkeypatch.setattr(lote_mod.parsers, "elegir", elegir_con_item_vacio)
    factura(sandbox.entrada / "Vehicular" / "i.pdf", 24, items=[("Auxilio tarea 412345", "60.000,00"),
                                                                 ("Peaje tarea 412346", "1.000,00"),
                                                                 ("Remolque tarea 412347", "39.000,00")])
    lote = armar(sandbox)
    assert [i.monto for i in lote.facturas[0].ext.items] == [D("60000.00"), None, D("39000.00")]
    assert len(lote.filas) == 3 and not lote.pendientes


def test_vencimiento_no_es_el_del_cae(sandbox):
    """C1: without a payment due date, F gets the issue date (with a warning), never the CAE's due date."""
    factura(sandbox.entrada / "Vehicular" / "c.pdf", 25, vto_pago=False)
    f = armar(sandbox).facturas[0]
    assert f.ext.vto == datetime.date(2026, 10, 15)
    factura(sandbox.entrada / "Vehicular" / "c.pdf", 25)
    assert armar(sandbox).facturas[0].ext.vto == datetime.date(2026, 11, 30)


def test_no_gravado_va_en_g_y_h_lleva_el_iva_real(sandbox):
    """C5: taxed 100.000 + not taxed 5.000 (tolls) + IVA 21.000 = 126.000: G 105.000, H 21.000, I 126.000."""
    factura(sandbox.entrada / "Vehicular" / "ng.pdf", 41, neto="100.000,00", iva="21.000,00", total="126.000,00",
            extras=[("No Gravado:", "5.000,00")])
    lote = armar(sandbox)
    f = lote.facturas[0]
    assert (f.ext.neto, f.ext.no_gravado) == (D("100000.00"), D("5000.00"))
    (r,) = lote.filas
    assert (r.monto, r.iva) == (D("105000.00"), D("21000.00"))
    assert not [a for a in lote.avisos if a.codigo == "total"]


OTRO_CUIT = _cuit("3069999888")


def test_cuit_del_emisor_es_el_de_la_etiqueta_arriba(sandbox):
    """C3: another valid CUIT printed before the issuer's (a carrier, a bank) is not taken as the provider."""
    factura(sandbox.entrada / "Vehicular" / "c3.pdf", 51, cabecera=f"Transporte asociado {_fmt(OTRO_CUIT)}")
    lote = armar(sandbox)
    assert lote.facturas[0].ext.cuit == CUIT
    assert any("varios CUIT" in a.texto and a.nivel == Nivel.REVISAR for a in lote.avisos)
