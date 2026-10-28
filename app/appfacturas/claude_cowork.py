"""Modo B: invoices the app could not read are handed to Claude (Cowork) through a folder.

  Facturas pendientes\\_Revisar con Claude\\<lote>\\
      pedido.md       what to do, in plain words (Claude reads this)
      pedido.json     the files and their SHA256 (the app reads this back)
      <copies of the PDFs>
      respuesta.json  written by Claude: one entry per file

The originals stay in _Entrada. When the batch is read again, an invoice whose SHA256 has an answer is
taken from it (source "claude": violet in the grid, always with a warning to review). Never sent: .txt files,
and PDFs with a CBU printed on them unless the user ticks them (Claude is also told not to copy bank details)."""
from __future__ import annotations

import json
import re
import shutil
from datetime import datetime
from pathlib import Path

from .model import Estado, Extraccion, Item
from .util import cbu_valido, cuit_valido, parse_fecha, parse_monto, sha256, solo_digitos

CARPETA = "_Revisar con Claude"

PEDIDO = """# Pedido de RUA EZ-Fac: leer facturas que la app no pudo leer

En esta carpeta hay {n} PDF de facturas de proveedores de RUA Asistencia (CUIT {cuit_rua}). La app no los pudo leer
(escaneados o con un formato que no conoce). Necesito que leas cada uno y escribas **`respuesta.json` en esta misma
carpeta**, con este formato exacto (un elemento por archivo, en el mismo orden):

```json
{modelo}
```

Reglas:
- `archivo`: el nombre del PDF tal cual está en la carpeta.
- `letra`: A, B, C o M (el recuadro grande de arriba). `tipo`: "Factura", "Nota de Crédito", "Nota de Débito", "Recibo" o "Nota de Venta".
- `cuit`: el CUIT del **proveedor que emite** la factura, solo números (nunca el de RUA, {cuit_rua}). `cuit_receptor`: el del cliente.
- `punto_de_venta` y `numero`: los dos números de "0002-00000189" (como números). Si no es una factura fiscal (sin CUIT ni CAE), dejalos en null y poné su número propio en `numero_no_fiscal` (por ejemplo "FAC0297").
- Fechas como dd/mm/aaaa. `vencimiento`: el de pago (nunca el "Vto. CAE"); vacío si no tiene. `fecha_desde`: si figura un período del servicio.
- Montos como en la factura ("121.000,00"). En una Factura A: `neto` es el neto gravado, `iva` el total de IVA, `no_gravado` lo no gravado o exento (si hay). En B o C: `neto` e `iva` vacíos y solo `total`.
- `items`: el detalle, con `monto` (neto en la A, total en la B/C) y `tareas`: los números de tarea o expediente de 6 cifras que figuren.
- **No copies CBU, CVU, alias, números de cuenta ni ningún dato bancario o personal.** Solo los campos de arriba.
- Si un PDF no se puede leer o no es una factura, poné en `no_pude` por qué y dejá el resto vacío. No inventes nada: si un dato no está, dejalo vacío.

Archivos:
{lista}
"""

MODELO = {"facturas": [{
    "archivo": "ejemplo.pdf", "razon_social": "", "cuit": "", "cuit_receptor": "", "letra": "", "tipo": "Factura",
    "punto_de_venta": None, "numero": None, "numero_no_fiscal": "", "fecha_emision": "", "fecha_desde": "",
    "vencimiento": "", "neto": "", "iva": "", "no_gravado": "", "total": "",
    "items": [{"descripcion": "", "monto": "", "tareas": []}], "no_pude": ""}]}


def raiz(cfg) -> Path:
    return cfg.base / CARPETA


def candidatas(lote) -> list:
    """Invoices of the batch worth sending: not read (or read without CUIT / number), only PDFs, and not already
    answered by Claude (those are in respondidas_sin_cargar)."""
    return [f for f in lote.facturas if f.estado == Estado.NO_LEIDA and f.archivo.ruta.suffix.lower() == ".pdf"
            and f.archivo.ruta.exists() and not (f.ext and f.ext.fuente == "claude")]


def respondidas_sin_cargar(lote) -> list:
    """Claude answered them, but what it read can't be loaded (e.g. an invalid CUIT)."""
    return [f for f in lote.facturas if f.estado == Estado.NO_LEIDA and f.ext and f.ext.fuente == "claude"]


def por_que_no(e) -> str:
    """Why Claude's reading can't be loaded, in plain words."""
    if not e.cuit and not e.numero_txt:
        return "no encontró el CUIT del proveedor"
    if e.cuit and not cuit_valido(e.cuit):
        return f"el CUIT {e.cuit} no es válido (no da el dígito verificador)"
    if e.cuit and (e.pv is None or e.nro is None):
        return "no encontró el punto de venta o el número"
    if not e.cuit and e.total is None:
        return "no encontró el total"
    return "faltan datos para cargarla"


def tiene_cbu(f) -> bool:
    """A CBU printed on the PDF (non-fiscal receipts often have one): not sent unless the user ticks it."""
    if f.ext and f.ext.cbu:
        return True
    try:
        from . import pdf
        doc = pdf.leer(f.archivo.ruta)
        texto = " ".join(p.texto for p in doc.paginas).replace(" ", "")
    except Exception:
        return False
    return any(cbu_valido(m) for m in re.findall(r"(?<!\d)\d{22}(?!\d)", texto))


def preparar(cfg, lote, facturas: list) -> Path:
    """Copies the PDFs and writes the request. Returns the folder (one per batch; written again if it exists)."""
    d = raiz(cfg) / lote.id
    d.mkdir(parents=True, exist_ok=True)
    archivos, usados = [], set()
    for f in facturas:
        nombre = Path(f.archivo.nombre.split(" > ")[-1]).name
        base, ext, i = Path(nombre).stem, Path(nombre).suffix, 2
        while nombre.lower() in usados:
            nombre, i = f"{base} ({i}){ext}", i + 1
        usados.add(nombre.lower())
        shutil.copy2(f.archivo.ruta, d / nombre)
        archivos.append({"archivo": nombre, "sha256": f.archivo.sha256, "area": f.archivo.area,
                         "origen": f.archivo.nombre})
    (d / "pedido.json").write_text(json.dumps({"lote": lote.id, "fecha": datetime.now().isoformat(timespec="seconds"),
                                               "archivos": archivos}, ensure_ascii=False, indent=1), encoding="utf-8")
    lista = "\n".join(f"- {a['archivo']}" for a in archivos)
    (d / "pedido.md").write_text(PEDIDO.format(n=len(archivos), cuit_rua=cfg.cuit_rua, lista=lista,
                                               modelo=json.dumps(MODELO, ensure_ascii=False, indent=1)), encoding="utf-8")
    return d


def pedidos(cfg) -> list[dict]:
    """Every request folder: {"carpeta", "archivos", "respondidos"} (newest first)."""
    out = []
    r = raiz(cfg) if cfg.base else None
    if not r or not r.is_dir():
        return out
    for d in sorted((x for x in r.iterdir() if x.is_dir()), reverse=True):
        try:
            p = json.loads((d / "pedido.json").read_text(encoding="utf-8"))
        except (OSError, ValueError):
            continue
        out.append({"carpeta": d, "archivos": len(p["archivos"]), "respondidos": len(_respuestas_de(d, p))})
    return out


def _respuestas_de(d: Path, pedido: dict) -> dict[str, dict]:
    """sha256 -> answer, for the files of this request that Claude answered."""
    try:
        r = json.loads((d / "respuesta.json").read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}
    por_nombre = {str(x.get("archivo", "")).lower(): x for x in (r.get("facturas") or []) if isinstance(x, dict)}
    out = {}
    for a in pedido["archivos"]:
        x = por_nombre.get(a["archivo"].lower())
        if x and not x.get("no_pude"):
            out[a["sha256"]] = x
    return out


def respuestas(cfg) -> dict[str, dict]:
    """All answers on disk, by the SHA256 of the original PDF (the newest request wins)."""
    out: dict[str, dict] = {}
    for p in reversed(pedidos(cfg)):
        try:
            ped = json.loads((p["carpeta"] / "pedido.json").read_text(encoding="utf-8"))
        except (OSError, ValueError):
            continue
        out.update(_respuestas_de(p["carpeta"], ped))
    return out


def _entero(v):
    d = solo_digitos(str(v if v is not None else ""))
    return int(d) if d else None


def extraccion(x: dict, cfg) -> Extraccion:
    """Claude's answer as the same Extraccion the parsers return (source "claude")."""
    e = Extraccion(fuente="claude")
    e.cuit = solo_digitos(str(x.get("cuit") or ""))
    if e.cuit == cfg.cuit_rua:
        e.cuit = ""
    e.cuit_receptor = solo_digitos(str(x.get("cuit_receptor") or ""))
    e.letra = str(x.get("letra") or "").strip().upper()[:1]
    tipo = str(x.get("tipo") or "Factura").strip()
    e.tipo = tipo if tipo in ("Factura", "Nota de Crédito", "Nota de Débito", "Recibo", "Nota de Venta") else "Factura"
    e.pv, e.nro = _entero(x.get("punto_de_venta")), _entero(x.get("numero"))
    e.numero_txt = str(x.get("numero_no_fiscal") or "").strip()
    e.razon_social = re.sub(r"\s+", " ", str(x.get("razon_social") or "")).strip()
    e.emision = parse_fecha(str(x.get("fecha_emision") or ""))
    e.desde = parse_fecha(str(x.get("fecha_desde") or ""))
    e.vto = parse_fecha(str(x.get("vencimiento") or ""))
    m = lambda k: parse_monto(str(x.get(k))) if x.get(k) not in (None, "") else None
    e.neto, e.iva, e.no_gravado, e.total = m("neto"), m("iva"), m("no_gravado"), m("total")
    for it in x.get("items") or []:
        if isinstance(it, dict):
            tareas = [solo_digitos(str(t)) for t in it.get("tareas") or [] if len(solo_digitos(str(t))) == 6]
            e.items.append(Item(descripcion=re.sub(r"\s+", " ", str(it.get("descripcion") or "")).strip(),
                                monto=parse_monto(str(it.get("monto"))) if it.get("monto") not in (None, "") else None,
                                tareas=[t for t in tareas if cfg.tarea_min <= int(t) <= cfg.tarea_max]))
    e.fiscal = bool(cuit_valido(e.cuit) and e.nro is not None)
    e.notas.append("La leyó Claude (Cowork): revisá CUIT, número y montos contra el PDF antes de confirmar")
    e.confianza = 0.5
    return e


def borrar(carpeta: Path) -> None:
    """Removes a request folder: only the app's copies, pedido.*, and respuesta.json are in it."""
    if carpeta.parent.name != CARPETA:
        raise ValueError(f"{carpeta} no es una carpeta de pedidos a Claude")
    shutil.rmtree(carpeta)
