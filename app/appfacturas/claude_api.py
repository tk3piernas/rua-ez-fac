"""Modo A: the invoices the app could not read are read through the Claude API.

The API key lives in the Windows Credential Manager (keyring) or in ANTHROPIC_API_KEY, never in a file.
What Claude reads is saved like a Cowork answer (modo B, claude_cowork.py): a request folder in
'_Revisar con Claude' with pedido.json and respuesta.json, so reading the batch again uses it the same way
(violet row, warning to review) and the request can be looked at or deleted afterwards.

Only the PDF is sent, never the .txt files, and PDFs with a CBU printed on them only if the user ticks them.
The answer schema has no field for bank details, and the prompt says not to copy them."""
from __future__ import annotations

import base64
import json
import os
from datetime import datetime
from pathlib import Path

from . import claude_cowork

SERVICIO = "RUA EZ-Fac"              # name of the entry in the Windows Credential Manager
USUARIO = "anthropic-api-key"
MODELO = "claude-opus-5-5"


class SinClave(Exception):
    """No API key configured."""


class ErrorClaude(Exception):
    """The API could not read one invoice (shown to the user; the rest go on)."""


# ------------------------------------------------------------------ the key

def clave() -> str:
    """The API key: Windows Credential Manager first, then the ANTHROPIC_API_KEY environment variable."""
    try:
        import keyring
        k = keyring.get_password(SERVICIO, USUARIO)
    except Exception:
        k = None
    return (k or os.environ.get("ANTHROPIC_API_KEY", "")).strip()


def origen_clave() -> str:
    """Where the key comes from, for the screen ("" if there is none). Never shows the key."""
    try:
        import keyring
        if keyring.get_password(SERVICIO, USUARIO):
            return "Administrador de credenciales de Windows"
    except Exception:
        pass
    return "variable de entorno ANTHROPIC_API_KEY" if os.environ.get("ANTHROPIC_API_KEY") else ""


def guardar_clave(k: str) -> None:
    import keyring
    k = (k or "").strip()
    if not k.startswith("sk-ant-"):
        raise ValueError("no parece una clave de la API de Claude (empieza con sk-ant-)")
    keyring.set_password(SERVICIO, USUARIO, k)


def borrar_clave() -> None:
    import keyring
    try:
        keyring.delete_password(SERVICIO, USUARIO)
    except keyring.errors.PasswordDeleteError:
        pass


# ------------------------------------------------------------------ reading

_TEXTO = {"type": "string"}
ESQUEMA = {
    "type": "object",
    "properties": {
        "razon_social": _TEXTO, "cuit": _TEXTO, "cuit_receptor": _TEXTO, "letra": _TEXTO, "tipo": _TEXTO,
        "punto_de_venta": _TEXTO, "numero": _TEXTO, "numero_no_fiscal": _TEXTO,
        "fecha_emision": _TEXTO, "fecha_desde": _TEXTO, "vencimiento": _TEXTO,
        "neto": _TEXTO, "iva": _TEXTO, "no_gravado": _TEXTO, "total": _TEXTO,
        "items": {"type": "array", "items": {
            "type": "object",
            "properties": {"descripcion": _TEXTO, "monto": _TEXTO, "tareas": {"type": "array", "items": _TEXTO}},
            "required": ["descripcion", "monto", "tareas"], "additionalProperties": False}},
        "no_pude": _TEXTO,
    },
    "required": ["razon_social", "cuit", "cuit_receptor", "letra", "tipo", "punto_de_venta", "numero",
                 "numero_no_fiscal", "fecha_emision", "fecha_desde", "vencimiento", "neto", "iva", "no_gravado",
                 "total", "items", "no_pude"],
    "additionalProperties": False,
}

INSTRUCCIONES = """Leé esta factura de un proveedor de RUA Asistencia (CUIT {cuit_rua}) y completá los campos.
- letra: A, B, C o M (el recuadro grande de arriba). tipo: Factura, Nota de Crédito, Nota de Débito, Recibo o Nota de Venta.
- cuit: el del proveedor que emite, solo números (nunca el de RUA). cuit_receptor: el del cliente.
- punto_de_venta y numero: los dos números de "0002-00000189", solo dígitos. Si no es fiscal (sin CUIT ni CAE), dejalos vacíos y poné su número propio en numero_no_fiscal (por ejemplo FAC0297).
- Fechas dd/mm/aaaa. vencimiento: el de pago, nunca el "Vto. CAE"; vacío si no tiene. fecha_desde: si figura un período del servicio.
- Montos como en la factura. Factura A: neto = neto gravado, iva = total de IVA, no_gravado = no gravado o exento. B o C: neto e iva vacíos, solo total.
- items: el detalle, con monto (neto en la A, total en B/C) y tareas: números de tarea o expediente de 6 cifras.
- No copies CBU, CVU, alias, números de cuenta ni ningún dato bancario o personal.
- Si no se puede leer o no es una factura, explicá por qué en no_pude y dejá el resto vacío. No inventes: si un dato no está, dejalo vacío."""


def leer_pdf(cliente, ruta: Path, cfg) -> dict:
    """One invoice through the API. Returns the answer in the same shape as a Cowork respuesta.json entry."""
    import anthropic
    datos = base64.standard_b64encode(ruta.read_bytes()).decode("ascii")
    try:
        r = cliente.beta.messages.create(
            model=MODELO,
            max_tokens=16000,
            betas=["server-side-fallback-2026-07-01"],
            fallbacks="default",                       # a policy decline is retried on a fallback model
            output_config={"effort": "medium", "format": {"type": "json_schema", "schema": ESQUEMA}},
            messages=[{"role": "user", "content": [
                {"type": "document", "source": {"type": "base64", "media_type": "application/pdf", "data": datos}},
                {"type": "text", "text": INSTRUCCIONES.format(cuit_rua=cfg.cuit_rua)},
            ]}],
        )
    except anthropic.AuthenticationError as e:
        raise SinClave("la clave de la API no es válida (revisala en Configuración)") from e
    except anthropic.PermissionDeniedError as e:
        raise ErrorClaude("la clave no tiene permiso para usar la API") from e
    except anthropic.RateLimitError as e:
        raise ErrorClaude("la API está saturada o se llegó al límite de uso; probá en un rato") from e
    except anthropic.BadRequestError as e:
        raise ErrorClaude(f"la API rechazó el pedido: {e.message}") from e
    except anthropic.APIStatusError as e:
        raise ErrorClaude(f"error de la API ({e.status_code}); probá de nuevo más tarde") from e
    except anthropic.APIConnectionError as e:
        raise ErrorClaude("no hay conexión con la API de Claude") from e
    if r.stop_reason == "refusal":
        raise ErrorClaude("Claude no quiso leer este archivo")
    if r.stop_reason == "max_tokens":
        raise ErrorClaude("la respuesta quedó cortada")
    texto = next((b.text for b in r.content if b.type == "text"), "")
    try:
        x = json.loads(texto)
    except ValueError as e:
        raise ErrorClaude("la respuesta no vino en el formato esperado") from e
    x["archivo"] = ruta.name
    return x


def leer(cfg, lote, facturas: list, log=print) -> dict:
    """Reads the chosen invoices through the API and saves the answers as a request folder.
    Returns {"carpeta", "leidas", "errores": [(name, why)]}. Raises SinClave if there is no key."""
    k = clave()
    if not k:
        raise SinClave("falta la clave de la API de Claude (Configuración → Claude)")
    import anthropic
    cliente = anthropic.Anthropic(api_key=k)
    d = claude_cowork.raiz(cfg) / f"{lote.id} API"
    d.mkdir(parents=True, exist_ok=True)
    archivos, respuestas, errores, usados = [], [], [], set()
    for f in facturas:
        nombre = Path(f.archivo.nombre.split(" > ")[-1]).name
        base, ext, i = Path(nombre).stem, Path(nombre).suffix, 2
        while nombre.lower() in usados:
            nombre, i = f"{base} ({i}){ext}", i + 1
        usados.add(nombre.lower())
        try:
            x = leer_pdf(cliente, f.archivo.ruta, cfg)
        except SinClave:
            raise
        except ErrorClaude as ex:
            errores.append((f.archivo.nombre, str(ex)))
            log(f"CLAUDE {f.archivo.nombre}: {ex}")
            continue
        x["archivo"] = nombre
        archivos.append({"archivo": nombre, "sha256": f.archivo.sha256, "area": f.archivo.area,
                         "origen": f.archivo.nombre})
        respuestas.append(x)
        log(f"CLAUDE {f.archivo.nombre}: leída")
    (d / "pedido.json").write_text(json.dumps({"lote": lote.id, "fecha": datetime.now().isoformat(timespec="seconds"),
                                               "modo": "API", "modelo": MODELO, "archivos": archivos},
                                              ensure_ascii=False, indent=1), encoding="utf-8")
    (d / "respuesta.json").write_text(json.dumps({"facturas": respuestas}, ensure_ascii=False, indent=1),
                                      encoding="utf-8")
    return {"carpeta": d, "leidas": len(respuestas), "errores": errores}
