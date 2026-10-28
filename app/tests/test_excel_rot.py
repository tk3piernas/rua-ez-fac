"""Finding the open workbook in Excel (Running Object Table) without taking a different one with the same name."""
import sys
import types

import pytest

from appfacturas import excel
from appfacturas.config import Config

URL_BASE = "https://empresa-my.sharepoint.com/personal/usuario_empresa_com/Documents/Facturas pendientes"
URL_REAL = "https://empresa-my.sharepoint.com/personal/usuario_empresa_com/Documents/DEUDA%20A%20PAGAR%20ARG.xlsx"
URL_OTRA = "https://otra-my.sharepoint.com/personal/x/Documents/Exportadas/DEUDA%20A%20PAGAR%20ARG.xlsx"


def _config_real(tmp_path):
    raiz = tmp_path / "OneDrive - Empresa"
    (raiz / "Facturas pendientes").mkdir(parents=True)
    toml = tmp_path / "config.toml"
    toml.write_text(f'[rutas]\nfacturas_pendientes = "{(raiz / "Facturas pendientes").as_posix()}"\n'
                    f'url_base = "{URL_BASE}"\n', encoding="utf-8")
    return toml, raiz


def test_url_solo_para_la_planilla_real(tmp_path):
    toml, raiz = _config_real(tmp_path)
    real = Config(base=raiz / "Facturas pendientes", planilla=raiz / "DEUDA A PAGAR ARG.xlsx", archivo=toml)
    assert excel.url_onedrive_de(real, real.planilla) == URL_REAL
    # test mode: same config.toml, but the sheet is the copy in the test folder
    prueba = tmp_path / "prueba" / "OneDrive - Empresa"
    cfg_prueba = Config(base=prueba / "Facturas pendientes", planilla=prueba / "DEUDA A PAGAR ARG.xlsx",
                        url_base=URL_BASE, archivo=toml)
    assert excel.url_onedrive_de(cfg_prueba, cfg_prueba.planilla) == ""
    # an exported copy somewhere else
    assert excel.url_onedrive_de(real, tmp_path / "Exportadas" / "DEUDA A PAGAR ARG.xlsx") == ""


def test_es_el_libro(tmp_path):
    ruta = tmp_path / "DEUDA A PAGAR ARG.xlsx"
    assert excel.es_el_libro(str(ruta), ruta, "")
    assert excel.es_el_libro(URL_REAL, ruta, URL_REAL)
    assert excel.es_el_libro(URL_REAL.replace("%20", " "), ruta, URL_REAL)     # FullName comes decoded
    assert not excel.es_el_libro(URL_OTRA, ruta, URL_REAL)                      # same name, other URL
    assert not excel.es_el_libro(URL_REAL, ruta, "")                            # test mode: no URL at all
    assert not excel.es_el_libro(str(tmp_path / "otra" / ruta.name), ruta, URL_REAL)


class _Moniker:
    def __init__(self, nombre):
        self.nombre = nombre

    def GetDisplayName(self, ctx, x):
        return self.nombre


class _Libro:
    def __init__(self, full):
        self.FullName = full

    def QueryInterface(self, iid):
        return self


def _rot_simulada(monkeypatch, nombres):
    libros = {n: _Libro(n) for n in nombres}

    class Rot:
        def __iter__(self):
            return iter([_Moniker(n) for n in nombres])

        def GetObject(self, m):
            return libros[m.nombre]

    monkeypatch.setitem(sys.modules, "pythoncom", types.SimpleNamespace(
        GetRunningObjectTable=lambda: Rot(), CreateBindCtx=lambda n: None, IID_IDispatch=None))
    client = types.SimpleNamespace(Dispatch=lambda o: o)
    monkeypatch.setitem(sys.modules, "win32com", types.SimpleNamespace(client=client))
    monkeypatch.setitem(sys.modules, "win32com.client", client)


def test_buscar_abierto_no_toma_otro_libro_con_el_mismo_nombre(tmp_path, monkeypatch):
    ruta = tmp_path / "prueba" / "DEUDA A PAGAR ARG.xlsx"
    _rot_simulada(monkeypatch, [URL_REAL, URL_OTRA])
    # test mode: the real sheet open in Excel must never be taken for the test copy
    assert excel.LibroCOM._buscar_abierto(ruta, "") is None
    assert excel.LibroCOM._buscar_abierto(ruta, URL_REAL).FullName == URL_REAL


def test_buscar_abierto_por_ruta_local(tmp_path, monkeypatch):
    ruta = tmp_path / "DEUDA A PAGAR ARG.xlsx"
    _rot_simulada(monkeypatch, [URL_OTRA, str(ruta)])
    assert excel.LibroCOM._buscar_abierto(ruta, "").FullName == str(ruta)


def test_revertir_deja_el_libro_abierto_como_guardado(tmp_path, monkeypatch):
    """A6: a failure halfway with the sheet open in Excel reverts and leaves it 'saved' (no false LibroConCambios)."""
    from types import SimpleNamespace as NS

    from appfacturas import preparar
    cfg = Config(base=tmp_path / "Facturas pendientes", planilla=tmp_path / "DEUDA A PAGAR ARG.xlsx",
                 excel_motor="openpyxl")
    preparar.crear_planilla_vacia(cfg.planilla)
    abiertos = []

    def abrir(c, ruta=None):
        libro = excel.LibroOpenpyxl(ruta or c.planilla)
        libro.ya_abierto = True
        hoja = libro.hoja

        def hoja_que_ensucia(nombre):
            libro.saved = False          # like Excel: any write marks the open workbook dirty
            return hoja(nombre)
        libro.hoja = hoja_que_ensucia
        abiertos.append(libro)
        return libro
    monkeypatch.setattr(excel, "abrir", abrir)
    fila = dict(proveedor="Prov Inventado", tarea="", servicio="", empresa="", fecha=None, vto=None,
                monto=None, iva=None, url="", texto_link="F 1", resaltar=False, comentario="",
                factura_id="f1", excluida=False)
    lote = NS(filas=[NS(hoja="Vehicular", **fila), NS(hoja="NoExiste", **fila)], proveedores=[])
    diario = NS(anotar=lambda paso: None)
    with pytest.raises(excel.ExcelError):
        excel.escribir(cfg, lote, diario, tmp_path / "backups", log=lambda s: None)
    assert abiertos[0].saved is True
    assert abiertos[0].hoja("Vehicular").formula(excel.PRIMERA, "A") in ("", None)
