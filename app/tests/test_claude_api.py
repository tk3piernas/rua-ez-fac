"""G3 modo A: unread invoices read through the Claude API. The API and the key store are simulated (no money,
no real Windows Credential Manager)."""
import json
from types import SimpleNamespace as NS

import keyring
import pytest
from keyring.backend import KeyringBackend

from appfacturas import claude_api, claude_cowork
from appfacturas.model import Nivel
from test_claude_cowork import _escaneada
from test_lote_sintetico import CUIT, armar, sandbox  # noqa: F401  (fixture)


class Llavero(KeyringBackend):
    priority = 1

    def __init__(self):
        self.d = {}

    def get_password(self, s, u):
        return self.d.get((s, u))

    def set_password(self, s, u, p):
        self.d[(s, u)] = p

    def delete_password(self, s, u):
        if (s, u) not in self.d:
            raise keyring.errors.PasswordDeleteError(u)
        del self.d[(s, u)]


@pytest.fixture
def llavero(monkeypatch):
    previo = keyring.get_keyring()
    keyring.set_keyring(Llavero())
    monkeypatch.delenv("ANTHROPIC_API_KEY", raising=False)
    yield
    keyring.set_keyring(previo)


def test_clave_en_el_llavero(llavero, monkeypatch):
    assert claude_api.clave() == "" and claude_api.origen_clave() == ""
    with pytest.raises(ValueError):
        claude_api.guardar_clave("no es una clave")
    claude_api.guardar_clave("sk-ant-prueba-123")
    assert claude_api.clave() == "sk-ant-prueba-123"
    assert claude_api.origen_clave() == "Administrador de credenciales de Windows"
    claude_api.borrar_clave()
    claude_api.borrar_clave()                                     # twice: nothing to delete, no error
    monkeypatch.setenv("ANTHROPIC_API_KEY", "sk-ant-env")
    assert claude_api.clave() == "sk-ant-env" and claude_api.origen_clave().startswith("variable")


def _cliente(texto="{}", stop="end_turn"):
    pedidos = []

    def create(**kw):
        pedidos.append(kw)
        return NS(stop_reason=stop, content=[NS(type="text", text=texto)])
    return NS(beta=NS(messages=NS(create=create))), pedidos


def test_leer_pdf_manda_el_pdf_con_esquema(sandbox, tmp_path):
    pdf = tmp_path / "f.pdf"
    _escaneada(pdf)
    cliente, pedidos = _cliente(json.dumps({"cuit": CUIT, "letra": "C", "total": "1.000,00"}))
    x = claude_api.leer_pdf(cliente, pdf, sandbox)
    assert x["cuit"] == CUIT and x["archivo"] == "f.pdf"
    kw = pedidos[0]
    assert kw["model"] == "claude-opus-5-5" and kw["fallbacks"] == "default"
    assert kw["output_config"]["format"]["type"] == "json_schema"
    assert "temperature" not in kw
    contenido = kw["messages"][0]["content"]
    assert contenido[0]["type"] == "document" and contenido[0]["source"]["media_type"] == "application/pdf"
    assert "No copies CBU" in contenido[1]["text"]
    assert "cbu" not in json.dumps(claude_api.ESQUEMA).lower()           # no field for bank details


def test_leer_pdf_rechazo_y_formato(sandbox, tmp_path):
    pdf = tmp_path / "f.pdf"
    _escaneada(pdf)
    with pytest.raises(claude_api.ErrorClaude):
        claude_api.leer_pdf(_cliente(stop="refusal")[0], pdf, sandbox)
    with pytest.raises(claude_api.ErrorClaude):
        claude_api.leer_pdf(_cliente(texto="no es json")[0], pdf, sandbox)


def test_leer_guarda_y_la_relectura_la_usa(sandbox, llavero, monkeypatch):
    _escaneada(sandbox.entrada / "Vehicular" / "scan.pdf")
    lote = armar(sandbox)
    with pytest.raises(claude_api.SinClave):
        claude_api.leer(sandbox, lote, claude_cowork.candidatas(lote))
    claude_api.guardar_clave("sk-ant-prueba-123")
    respuesta = {"razon_social": "Gruas API SRL", "cuit": CUIT, "cuit_receptor": "30717419428", "letra": "C",
                 "tipo": "Factura", "punto_de_venta": "4", "numero": "88", "numero_no_fiscal": "",
                 "fecha_emision": "20/10/2026", "fecha_desde": "", "vencimiento": "", "neto": "", "iva": "",
                 "no_gravado": "", "total": "50.000,00",
                 "items": [{"descripcion": "Remolque", "monto": "50.000,00", "tareas": ["412399"]}], "no_pude": ""}
    monkeypatch.setattr(claude_api, "leer_pdf", lambda cliente, ruta, cfg: dict(respuesta, archivo=ruta.name))
    r = claude_api.leer(sandbox, lote, claude_cowork.candidatas(lote), log=lambda s: None)
    assert r["leidas"] == 1 and r["errores"] == [] and r["carpeta"].name.endswith(" API")
    lote = armar(sandbox)
    f = lote.facturas[0]
    assert f.ext.fuente == "claude" and f.texto_link == "Factura 0004-00000088"
    assert lote.filas and lote.filas[0].tarea == "412399"
    assert any("Claude" in a.texto and a.nivel == Nivel.REVISAR for a in lote.avisos)


def test_un_error_no_frena_las_demas(sandbox, llavero, monkeypatch):
    _escaneada(sandbox.entrada / "Vehicular" / "a.pdf")
    _escaneada(sandbox.entrada / "Vehicular" / "b.pdf")
    lote = armar(sandbox)
    claude_api.guardar_clave("sk-ant-prueba-123")

    def leer_pdf(cliente, ruta, cfg):
        if ruta.name == "a.pdf":
            raise claude_api.ErrorClaude("no hay conexión con la API de Claude")
        return {"cuit": CUIT, "letra": "C", "punto_de_venta": "1", "numero": "9", "total": "10,00", "items": []}
    monkeypatch.setattr(claude_api, "leer_pdf", leer_pdf)
    r = claude_api.leer(sandbox, lote, claude_cowork.candidatas(lote), log=lambda s: None)
    assert r["leidas"] == 1 and [n for n, _ in r["errores"]] == ["a.pdf"]
