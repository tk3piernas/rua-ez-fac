"""Invoices that are not the ARCA layout.

Covers fiscal invoices made with other billing software ("Punto de Venta:00002" + "Comprobante N°:";
or "Nº: 0001-00000001", "Neto:", "Fecha de venc. de pago:") and non-fiscal receipts without CUIT
or CAE ("Número: FAC0001", "Fecha: jun 13, 2026").

It works on lines rebuilt from the page positions, so the text order inside the PDF does not
matter. Confidence is lower than ARCA's: the result always goes to review."""
from __future__ import annotations

import re
from decimal import Decimal

from ..model import Extraccion, Item
from ..pdf import DocPdf, Span
from ..util import cbu_valido, cuit_valido, parse_fecha, parse_monto
from .base import Parser, es_monto, tareas_en, tipo_de_codigo

_MONTO_SOLO = re.compile(r"^-?\$?\s*-?\d[\d.,]*\d$")
_TOTALES = re.compile(r"(?i)^(sub-?\s?total|total\b|neto|i\.?\s?v\.?\s?a\b|pagada|importe|bonificaci|son\b|percep)")
_CUIT_RE = re.compile(r"(?<!\d)(\d{2})-?(\d{8})-?(\d)(?!\d)")


class Linea:
    def __init__(self, spans: list[Span]):
        self.spans = sorted(spans, key=lambda s: s.x0)
        self.y = min(s.yc for s in spans)

    @property
    def texto(self) -> str:
        return "  ".join(s.texto for s in self.spans)

    def montos(self) -> list[Span]:
        return [s for s in self.spans if _MONTO_SOLO.match(s.texto.strip()) and es_monto_flexible(s.texto)]


def es_monto_flexible(t: str) -> bool:
    t = t.strip()
    return es_monto(t) or bool(re.fullmatch(r"-?\$?\s*\d{1,3}(?:[.,]\d{3})*[.,]\d{2}", t))


def lineas(spans: list[Span]) -> list[Linea]:
    out: list[list[Span]] = []
    for s in sorted(spans, key=lambda s: s.yc):
        if out and abs(out[-1][0].yc - s.yc) <= 3:
            out[-1].append(s)
        else:
            out.append([s])
    return [Linea(g) for g in out]


def _buscar(pat: str, texto: str, grupo: int = 1) -> str:
    m = re.search(pat, texto)
    return m.group(grupo).strip() if m else ""


class Generico(Parser):
    nombre = "generico"

    def detectar(self, doc: DocPdf, nombre_archivo: str) -> float:
        t = doc.texto
        if not doc.tiene_texto:
            return 0.0
        if re.search(r"(?i)factura|comprobante|recibo|nota de cr[eé]dito", t) and re.search(r"(?i)\btotal\b", t):
            return 0.5
        return 0.0

    def extraer(self, doc: DocPdf, nombre_archivo: str, cfg) -> Extraccion:
        e = Extraccion(fuente=self.nombre)
        pg = doc.paginas[0]
        L = lineas(pg.spans)
        texto = "\n".join(l.texto for l in L)

        # --- type
        cod = _buscar(r"(?i)\bc[oó]d(?:igo)?\.?\s*:?\s*0*(\d{1,3})\b", texto)
        e.cod_comprobante = int(cod) if cod else None
        e.letra, e.tipo = tipo_de_codigo(e.cod_comprobante)
        caja = next((s.texto for s in pg.spans if s.size >= 20 and len(s.texto) == 1 and s.texto in "ABCM"), "")
        if caja and e.letra and caja != e.letra:
            e.notas.append(f"La letra del recuadro ({caja}) no coincide con el código {e.cod_comprobante}")
        e.letra = e.letra or caja
        if re.search(r"(?i)nota\s+de\s+cr[eé]dito", texto):
            e.tipo = "Nota de Crédito"
        elif re.search(r"(?i)nota\s+de\s+d[eé]bito", texto):
            e.tipo = "Nota de Débito"

        # --- number
        m = re.search(r"(?i)(?:N\s*[º°o]\.?|Nro\.?|N[uú]mero|Comprobante\s*N\s*[º°o]?\.?)\s*:?\s*(\d{4,5})\s*-\s*(\d{8})", texto)
        if m:
            e.pv, e.nro = int(m.group(1)), int(m.group(2))
        else:
            pv = _buscar(r"(?i)Punto\s+de\s+Venta\s*:?\s*(\d{1,5})", texto)
            nro = _buscar(r"(?i)(?:Comp\.?\s*Nro|Comprobante\s*N\s*[º°o]?\.?)\s*:?\s*(\d{1,8})\b", texto)
            if pv and nro:
                e.pv, e.nro = int(pv), int(nro)
        if e.nro is None:
            e.numero_txt = _buscar(r"(?i)\bN(?:[uú]mero|ro\.?|\s*[º°])\s*:?\s*([A-Z]{0,4}-?\d{2,10})\b", texto)

        # --- issuer CUIT: valid, not RUA. Preferred: on a line with the "CUIT" label in the top third of the
        # page (the issuer's header), over a bank's, a carrier's or an accountant's CUIT printed elsewhere
        cuits: list[str] = []
        puntaje: dict[str, tuple] = {}
        alto = pg.alto or 842
        for orden, l in enumerate(L):
            etiqueta = bool(re.search(r"(?i)c\.?\s?u\.?\s?i\.?\s?t", l.texto))
            arriba = l.y < alto / 3
            for mm in _CUIT_RE.finditer(l.texto):
                c = "".join(mm.groups())
                if c != cfg.cuit_rua and cuit_valido(c) and c[:2] in ("20", "23", "24", "27", "30", "33", "34") and c not in cuits:
                    cuits.append(c)
                    puntaje[c] = (not (etiqueta and arriba), not etiqueta, not arriba, orden)
        e.cuit = min(cuits, key=lambda c: puntaje[c]) if cuits else ""
        if len(cuits) > 1:
            e.notas.append(f"hay varios CUIT ({', '.join(cuits)}); tomé {e.cuit} como el del proveedor")
        c = cfg.cuit_rua
        e.cuit_receptor = c if re.search(rf"(?<!\d){c[:2]}-?{c[2:10]}-?{c[10:]}(?!\d)", texto) else ""
        e.condicion_receptor = _buscar(r"(?i)I\.?V\.?A\.?:\s*(Responsable Inscripto|Consumidor Final|Exento|Monotributo)", texto)
        cae = bool(re.search(r"(?i)\bC\.?A\.?E\.?\b", texto))
        e.fiscal = bool(e.cuit) and e.nro is not None and cae

        # --- issuer name: "FACTURA DE:" block, "Razón Social:", or the biggest text at the top
        e.razon_social = re.sub(r"\s+", " ", _nombre_emisor(pg, L)).strip()

        # --- dates
        e.emision = parse_fecha(_buscar(r"(?i)Fecha(?:\s+de\s+Emisi[oó]n)?\s*:\s*([^\n]{6,22})", texto))
        e.desde = parse_fecha(_buscar(r"(?i)Desde\s*:?\s*([\d/.-]{8,10})", texto))
        # never the CAE's due date ("Fecha de Vencimiento CAE"): it is ~10 days after the issue date
        sin_cae = "\n".join(l for l in texto.splitlines() if not re.search(r"(?i)\bC\.?A\.?E\b", l))
        vto = _buscar(r"(?i)(?:venc(?:imiento|\.)?\s+(?:de\s+)?pago|vto\.?\s+(?:para\s+el\s+)?pago|vencimiento)\s*:?\s*([^\n]{6,22})", sin_cae)
        e.vto = parse_fecha(vto)

        # --- totals (amount = rightmost amount on the label's line)
        def montos_de_label(pat):
            """Amounts printed to the right of a label, on the same line (one per matching label)."""
            out = []
            for l in L:
                for s in l.spans:
                    if re.match(pat, s.texto.strip(), re.I):
                        der = [m for m in l.montos() if m.x0 > s.x1]
                        if der:
                            out.append(parse_monto(der[-1].texto))
            return out

        def monto_linea(pat):
            v = montos_de_label(pat)
            return v[0] if v else None
        e.total = monto_linea(r"^(importe\s+)?total\b(?!\s*por)")
        e.neto = monto_linea(r"^(importe\s+)?neto(\s+gravado)?\b(?!\s*no\b)")
        sin_iva = [v for v in montos_de_label(r"^(importe\s+)?(neto\s+no\s+gravado|no\s+gravado|exento)\b") if v]
        e.no_gravado = sum(sin_iva, Decimal(0)) if sin_iva else None
        ivas = montos_de_label(r"^(importe\s+)?i\.?\s?v\.?\s?a\b")
        e.iva = sum((v for v in ivas if v is not None), Decimal(0)) if ivas else None
        if e.total is None:
            e.total = monto_linea(r"^sub-?\s?total\b")

        # --- CBU printed on the invoice (non-fiscal receipts usually carry it)
        for mm in re.finditer(r"(?<!\d)(\d{22})(?!\d)", texto.replace(" ", "")):
            if cbu_valido(mm.group(1)):
                e.cbu = mm.group(1)
                break

        e.items = _items(L, cfg)
        e.tareas_archivo = tareas_en(nombre_archivo, cfg.tarea_min, cfg.tarea_max)
        if not e.fiscal:
            e.notas.append("Comprobante no fiscal: no tiene " +
                           " ni ".join(x for x, ok in (("CUIT", e.cuit), ("CAE", cae),
                                                       ("punto de venta", e.nro is not None)) if not ok))
        falt = [c for c in ("total", "vto") if getattr(e, c) is None]
        if not (e.emision or e.desde):
            falt.append("fecha")
        if e.fiscal is False and not e.razon_social:
            falt.append("proveedor")
        if not e.items:
            falt.append("items")
        e.faltantes = falt
        e.confianza = max(0.0, 0.6 - 0.15 * len(falt))
        return e


def _nombre_emisor(pg, L: list[Linea]) -> str:
    sp = pg.spans
    lab = next((s for s in sp if re.match(r"(?i)^(factura|recibo)\s+de\s*:?$", s.texto)), None)
    if lab:
        partes = [s.texto for s in sorted(sp, key=lambda s: s.yc)
                  if abs(s.x0 - lab.x0) < 15 and lab.yc + 4 < s.yc < lab.yc + 45 and not s.texto.endswith(":")]
        if partes:
            return " ".join(partes[:2]).strip()
    m = next((l for l in L if re.match(r"(?i)^raz[oó]n\s+social\s*:", l.texto)), None)
    if m:
        v = re.sub(r"(?i)^raz[oó]n\s+social\s*:\s*", "", m.texto).split("  ")[0].strip()
        if v:
            return v
    arriba = [s for s in sp if s.y0 < pg.alto * 0.12 and len(s.texto) > 3 and not re.search(r"(?i)factura|original|electr", s.texto)]
    if arriba:
        return max(arriba, key=lambda s: s.size).texto.strip()
    return ""


def _items(L: list[Linea], cfg) -> list[Item]:
    hdr = next((l for l in L if re.search(r"(?i)\b(descripci[oó]n|detalle|producto)\b", l.texto)
                and re.search(r"(?i)\b(cantidad|importe|precio|sub-?total)\b", l.texto)), None)
    x_cant = None
    if hdr:
        c = next((s for s in hdr.spans if re.match(r"(?i)cantidad", s.texto)), None)
        x_cant = c.x0 - 6 if c else None
        x_desc0 = next((s.x0 for s in hdr.spans if re.search(r"(?i)descripci|detalle|producto", s.texto)), 0)
    else:
        x_desc0 = 0
    def es_total(l):
        return any(_TOTALES.match(s.texto.strip()) for s in l.spans)
    fin = next((l.y for l in L if es_total(l) and l.montos()), 1e9)
    filas = [l for l in L if l.y < fin and l is not hdr and l.montos() and not es_total(l)
             and (not hdr or l.y > hdr.y - 80)]
    if hdr:
        filas = [l for l in filas if not re.search(r"(?i)fecha|cuit|punto de venta|comprobante", l.texto)]
    items: list[Item] = []
    for i, fila in enumerate(filas):
        siguiente = filas[i + 1].y if i + 1 < len(filas) else fin
        limite = x_cant if x_cant else min(s.x0 for s in fila.montos()) - 6
        desc_spans = [s for l in L if fila.y - 2 <= l.y < siguiente - 2 and l is not hdr
                      for s in l.spans if s.x1 <= limite + 2 and not _MONTO_SOLO.match(s.texto)]
        codigo = " ".join(s.texto for s in desc_spans if s.x1 < x_desc0 - 2)
        desc = " ".join(s.texto for s in desc_spans if s.x1 >= x_desc0 - 2)
        monto = parse_monto(fila.montos()[-1].texto)
        items.append(Item(descripcion=re.sub(r"\s+", " ", desc).strip(), monto=monto,
                          tareas=tareas_en(f"{codigo} {desc}", cfg.tarea_min, cfg.tarea_max), codigo=codigo.strip()))
    return items


PARSER = Generico()
