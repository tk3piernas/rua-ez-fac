"""G3 modo B: unread invoices handed to Claude (Cowork) through a folder, and its answer used when reading again."""
import json
from decimal import Decimal as D

import pymupdf

from appfacturas import claude_cowork
from appfacturas.model import Estado, Nivel
from test_lote_sintetico import CUIT, armar, factura, sandbox  # noqa: F401  (fixture)
from test_casos_lectura import _cbu


def _escaneada(ruta):
    doc = pymupdf.open()
    doc.new_page().draw_rect(pymupdf.Rect(50, 50, 300, 300), color=(0, 0, 0), fill=(0.8, 0.8, 0.8))
    doc.save(ruta)


def test_pedido_respuesta_y_relectura(sandbox):
    _escaneada(sandbox.entrada / "Vehicular" / "scan.pdf")
    lote = armar(sandbox)
    (f,) = claude_cowork.candidatas(lote)
    assert f.estado == Estado.NO_LEIDA and not claude_cowork.tiene_cbu(f)

    d = claude_cowork.preparar(sandbox, lote, [f])
    assert d == sandbox.base / "_Revisar con Claude" / lote.id
    assert sorted(p.name for p in d.iterdir()) == ["pedido.json", "pedido.md", "scan.pdf"]
    pedido = (d / "pedido.md").read_text(encoding="utf-8")
    assert "No copies CBU" in pedido and "scan.pdf" in pedido
    assert (sandbox.entrada / "Vehicular" / "scan.pdf").exists()                 # the original stays
    assert claude_cowork.pedidos(sandbox)[0]["respondidos"] == 0

    (d / "respuesta.json").write_text(json.dumps({"facturas": [{
        "archivo": "scan.pdf", "razon_social": "Gruas Inventadas SRL", "cuit": CUIT, "cuit_receptor": "30717419428",
        "letra": "A", "tipo": "Factura", "punto_de_venta": 3, "numero": 301, "fecha_emision": "15/10/2026",
        "vencimiento": "30/11/2026", "neto": "100.000,00", "iva": "21.000,00", "total": "121.000,00",
        "items": [{"descripcion": "Auxilio mecánico", "monto": "100.000,00", "tareas": ["412345"]}]}]}),
        encoding="utf-8")
    assert claude_cowork.pedidos(sandbox)[0]["respondidos"] == 1

    lote = armar(sandbox)
    f = lote.facturas[0]
    assert f.ext.fuente == "claude" and (f.ext.cuit, f.ext.pv, f.ext.nro) == (CUIT, 3, 301)
    assert f.texto_link == "Factura 0003-00000301"
    (r,) = lote.filas
    assert (r.monto, r.tarea) == (D("100000.00"), "412345")
    assert any("Claude" in a.texto and a.nivel == Nivel.REVISAR for a in lote.avisos)
    assert not claude_cowork.candidatas(lote)

    claude_cowork.borrar(d)
    assert not d.exists()


def test_no_pude_no_se_usa(sandbox):
    _escaneada(sandbox.entrada / "Vehicular" / "scan.pdf")
    lote = armar(sandbox)
    d = claude_cowork.preparar(sandbox, lote, claude_cowork.candidatas(lote))
    (d / "respuesta.json").write_text(json.dumps({"facturas": [{"archivo": "scan.pdf", "no_pude": "no es una factura"}]}),
                                      encoding="utf-8")
    assert armar(sandbox).facturas[0].estado == Estado.NO_LEIDA


def test_pdf_con_cbu_se_marca(sandbox, tmp_path):
    ruta = sandbox.entrada / "Vehicular" / "recibo.pdf"
    doc = pymupdf.open()
    doc.new_page().insert_text((50, 80), f"Recibo de pago  CBU: {_cbu('0170099', '2200001234567')}", fontsize=10)
    doc.save(ruta)
    lote = armar(sandbox)
    (f,) = claude_cowork.candidatas(lote)
    assert claude_cowork.tiene_cbu(f)


def test_respuesta_con_cuit_invalido_avisa_por_que(sandbox):
    """The sample ARCA invoice has an invented CUIT: Claude reads it, the app explains why it can't load it."""
    _escaneada(sandbox.entrada / "Vehicular" / "scan.pdf")
    lote = armar(sandbox)
    d = claude_cowork.preparar(sandbox, lote, claude_cowork.candidatas(lote))
    (d / "respuesta.json").write_text(json.dumps({"facturas": [{
        "archivo": "scan.pdf", "cuit": "20010203047", "letra": "B", "punto_de_venta": 2, "numero": 1,
        "total": "8500.00"}]}), encoding="utf-8")
    lote = armar(sandbox)
    f = lote.facturas[0]
    assert f.estado == Estado.NO_LEIDA and f.ext.fuente == "claude"
    assert any("Claude la leyó, pero el CUIT 20010203047 no es válido" in p for p in lote.pendientes)
    assert any(a.codigo == "claude_no_carga" for a in lote.avisos)
    assert claude_cowork.candidatas(lote) == [] and claude_cowork.respondidas_sin_cargar(lote) == [f]
