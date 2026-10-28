"""What every parser shares: interface, AFIP codes, task numbers, span lookups."""
from __future__ import annotations

import re
from decimal import Decimal

from ..model import Extraccion
from ..pdf import DocPdf, Span

AFIP_NOMBRE_RE = re.compile(r"^(\d{11})_(\d{3})_(\d{4,5})_(\d{8})")
MONTO_RE = re.compile(r"^-?\$?\s*\d[\d.]*,\d{2}$|^-?\$?\s*\d{1,3}(,\d{3})*\.\d{2}$|^-?\d+,\d{2}$")


class Parser:
    """A parser = one file in this package with a module-level PARSER = MyParser()."""
    nombre = "base"

    def detectar(self, doc: DocPdf, nombre_archivo: str) -> float:
        """0 = not mine, 1 = surely mine."""
        raise NotImplementedError

    def extraer(self, doc: DocPdf, nombre_archivo: str, cfg) -> Extraccion:
        raise NotImplementedError


# ------------------------------------------------------------- AFIP codes

def tipo_de_codigo(cod: int | None) -> tuple[str, str]:
    """(letra, tipo) from the AFIP comprobante code. Includes FCE MiPyME (201-213)."""
    if cod is None:
        return "", "Factura"
    letra = ""
    for rango, l in (((1, 5), "A"), ((6, 10), "B"), ((11, 15), "C"), ((51, 55), "M"),
                     ((201, 205), "A"), ((206, 210), "B"), ((211, 215), "C")):
        if rango[0] <= cod <= rango[1]:
            letra = l
    if cod in (3, 8, 13, 53, 203, 208, 213):
        tipo = "Nota de Crédito"
    elif cod in (2, 7, 12, 52, 202, 207, 212):
        tipo = "Nota de Débito"
    elif cod in (4, 9, 15, 54):
        tipo = "Recibo"
    elif cod in (5, 10):
        tipo = "Nota de Venta"           # nota de venta al contado
    else:
        tipo = "Factura"
    return letra, tipo


def nombre_afip(nombre_archivo: str):
    """CUIT, cod, pv, nro from 'CUIT_TIPO_PV_NRO.pdf' (AFIP download name) or None."""
    m = AFIP_NOMBRE_RE.match(nombre_archivo)
    if not m:
        return None
    return m.group(1), int(m.group(2)), int(m.group(3)), int(m.group(4))


# ------------------------------------------------------------- task numbers

def tareas_en(texto: str, tmin: int, tmax: int) -> list[str]:
    """6-digit numbers in [tmin, tmax], not part of a longer number or an amount.

    Differs from the PS1 on purpose: "406419, 406800" and "406419/406800" are read
    (the PS1 skipped a number followed by a comma or glued to a slash)."""
    out = []
    texto = texto or ""
    for m in re.finditer(r"(?<![\d.,-])(\d{6})(?!\d|[.,]\d|-\d)", texto):
        v = m.group(1)
        antes = texto[max(0, m.start() - 30):m.start()]
        if _ES_PRECIO.search(antes):
            continue                         # "Precio unitario 350000", "$ 350000": an amount, not a task
        if tmin <= int(v) <= tmax and v not in out:
            out.append(v)
    return out


_ES_PRECIO = re.compile(r"(?i)(\$\s*|\b(?:precio|importe|total|subtotal|monto|unitario|p\.\s?unit\.?)\b(?:(?!tarea|exp|servicio)[^\d\n]){0,15})$")


# ------------------------------------------------------------- span helpers

def es_monto(t: str) -> bool:
    return bool(MONTO_RE.match(t.strip()))


def buscar_label(spans: list[Span], patron: str, y_min: float = -1, y_max: float = 1e9,
                 x_min: float = -1, x_max: float = 1e9) -> Span | None:
    rx = re.compile(patron, re.I)
    for s in spans:
        if y_min <= s.yc <= y_max and x_min <= s.x0 <= x_max and rx.match(s.texto):
            return s
    return None


def valor_de(spans: list[Span], patron: str, **kw) -> str:
    """Value printed to the right of a label on the same line (or after the label in the same span)."""
    lab = buscar_label(spans, patron, **kw)
    if not lab:
        return ""
    m = re.match(patron, lab.texto, re.I)
    resto = lab.texto[m.end():].strip()
    if resto:
        return resto
    cands = [s for s in spans if s is not lab and abs(s.yc - lab.yc) <= 3.5 and s.x0 >= lab.x1 - 1]
    cands.sort(key=lambda s: s.x0)
    for c in cands:
        if c.texto.endswith(":"):
            return ""          # the next thing on the line is another label
        return c.texto
    return ""


def suma(vals) -> Decimal:
    return sum((v for v in vals if v is not None), Decimal(0))
