"""C7 / C9: reading cases that had no test. Synthetic invoices, invented provider and numbers."""
import zipfile
from decimal import Decimal as D

import pymupdf

from appfacturas import txt_proveedor
from appfacturas.model import Estado, Nivel
from appfacturas.util import cbu_valido
from test_lote_sintetico import CUIT, _fmt, armar, factura, sandbox  # noqa: F401  (fixture)


def _cbu(banco7: str, cuenta13: str) -> str:
    b = next(banco7 + str(d) for d in range(10) if cbu_valido(banco7 + str(d) + "0" * 14))
    return next(b + cuenta13 + str(d) for d in range(10) if cbu_valido(b + cuenta13 + str(d)))


def test_nota_de_debito(sandbox):
    factura(sandbox.entrada / "Vehicular" / "nd.pdf", 61)
    doc = pymupdf.open(sandbox.entrada / "Vehicular" / "nd.pdf")
    pg = doc[0]
    pg.insert_text((360, 25), "NOTA DE DEBITO", fontsize=10)
    doc.saveIncr()
    doc.close()
    lote = armar(sandbox)
    f = lote.facturas[0]
    assert f.ext.tipo == "Nota de Débito" and f.texto_link.startswith("Nota de Débito 0003-")
    assert lote.filas[0].monto == D("100000.00")             # a debit note adds, it is not negative


def test_iva_10_5(sandbox):
    """Factura A at 10,5 %: H gets the real IVA (it differs from the sheet's 21 % formula)."""
    factura(sandbox.entrada / "Vehicular" / "r.pdf", 62, neto="100.000,00", iva="10.500,00", total="110.500,00")
    (r,) = armar(sandbox).filas
    assert (r.monto, r.iva) == (D("100000.00"), D("10500.00"))


def test_pdf_escaneado_queda_sin_leer(sandbox):
    doc = pymupdf.open()
    pg = doc.new_page()
    pg.draw_rect(pymupdf.Rect(50, 50, 300, 300), color=(0, 0, 0), fill=(0.8, 0.8, 0.8))   # only a picture
    doc.save(sandbox.entrada / "Vehicular" / "scan.pdf")
    lote = armar(sandbox)
    assert lote.facturas[0].estado == Estado.NO_LEIDA and not lote.filas
    assert any("escaneada" in p for p in lote.pendientes)
    assert not [p for p in lote.plan if p["op"] == "mover"]


def test_pdf_con_contrasena_queda_sin_leer(sandbox, tmp_path):
    origen = tmp_path / "abierta.pdf"
    factura(origen, 63)
    doc = pymupdf.open(origen)
    doc.save(sandbox.entrada / "Vehicular" / "clave.pdf", encryption=pymupdf.PDF_ENCRYPT_AES_256,
             user_pw="secreta", owner_pw="secreta")
    lote = armar(sandbox)
    assert lote.facturas[0].estado == Estado.NO_LEIDA and not lote.filas


def test_multipagina_lee_solo_el_original(sandbox, tmp_path):
    """ORIGINAL and DUPLICADO in the same PDF: one invoice, amounts not doubled."""
    una = tmp_path / "una.pdf"
    factura(una, 64)
    doc = pymupdf.open()
    for marca in ("ORIGINAL", "DUPLICADO"):
        doc.insert_pdf(pymupdf.open(una))
        doc[-1].insert_text((250, 15), marca, fontsize=10)
    doc.save(sandbox.entrada / "Vehicular" / "dos_copias.pdf")
    lote = armar(sandbox)
    assert len(lote.facturas) == 1
    (r,) = lote.filas
    assert r.monto == D("100000.00")


def test_zip_con_subcarpetas_y_nombres_repetidos(sandbox, tmp_path):
    a, b = tmp_path / "a.pdf", tmp_path / "b.pdf"
    factura(a, 65)
    factura(b, 66)
    with zipfile.ZipFile(sandbox.entrada / "Vehicular" / "lote.zip", "w") as z:
        z.write(a, "enero/factura.pdf")
        z.write(b, "febrero/factura.pdf")
    lote = armar(sandbox)
    nros = sorted(f.ext.nro for f in lote.facturas if f.ext)
    assert nros == [65, 66] and len(lote.filas) == 2
    destinos = [p["a"] for p in lote.plan if p["op"] == "copiar"]
    assert len(set(destinos)) == 2


def test_txt_con_dos_cbu():
    uno, dos = _cbu("0170099", "2200001234567"), _cbu("0070123", "4000009876543")
    d = txt_proveedor.desde_texto(f"Proveedor: Grúas Inventadas\nCUIT: {_fmt(CUIT)}\n"
                                  f"CBU: {uno}\nAlias: GRUAS.INV\n\nCBU: {dos}\n")
    assert [c.cbu for c in d.cuentas] == [uno, dos]


def test_mismo_numero_con_otro_importe_contra_las_carpetas(sandbox, tmp_path):
    """Already saved in the provider's folder with the same number but another amount: stopped."""
    guardada = sandbox.base / "Vehicular" / "Gruas Inventadas" / "Factura 0003-00000067.pdf"
    guardada.parent.mkdir(parents=True)
    factura(guardada, 67, neto="90.000,00", iva="18.900,00", total="108.900,00")
    csv = sandbox.base / "_Script" / "proveedores.csv"
    csv.write_text(f"CUIT;Area;Proveedor;Carpeta\r\n{CUIT};Vehicular;Gruas Inventadas;Gruas Inventadas\r\n",
                   encoding="utf-8-sig")
    factura(sandbox.entrada / "Vehicular" / "nueva.pdf", 67)
    lote = armar(sandbox)
    assert [a.codigo for a in lote.avisos if a.nivel == Nivel.BLOQUEA] == ["mismo_numero"]
    assert all(r.excluida for r in lote.filas)
    # same number and same amount (another copy of the same invoice): already saved, not stopped
    factura(guardada, 67)
    lote = armar(sandbox)
    assert not [a for a in lote.avisos if a.nivel == Nivel.BLOQUEA]
