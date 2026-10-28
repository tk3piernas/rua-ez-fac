"""ARCA (ex AFIP) electronic invoices: "Comprobante Autorizado" layout, any letter.

Reads by position: each label's value is the text printed to its right on the same line,
so the order in which the PDF stores its text does not matter."""
from __future__ import annotations

import re
from decimal import Decimal

from ..model import Extraccion, Item
from ..pdf import DocPdf, Span
from ..util import cuit_valido, parse_fecha, parse_monto, solo_digitos
from .base import (Parser, buscar_label, es_monto, nombre_afip, suma, tareas_en,
                   tipo_de_codigo, valor_de)

_TOTALES = r"^(Subtotal:\s*\$|Importe Otros Tributos|Importe Neto (No )?Gravado|Importe Exento|IVA [\d.,]+\s*%|CAE N|Importe Total)"


class Arca(Parser):
    nombre = "arca"

    def detectar(self, doc: DocPdf, nombre_archivo: str) -> float:
        t = doc.texto
        if "Punto de Venta" in t and re.search(r"Comp\.?\s*Nro", t):
            return 1.0 if re.search(r"COD\.\s*\d{2,3}", t) else 0.8
        return 0.0

    def extraer(self, doc: DocPdf, nombre_archivo: str, cfg) -> Extraccion:
        e = Extraccion(fuente=self.nombre)
        pg = doc.paginas[0]
        sp = pg.spans
        periodo = buscar_label(sp, r"Per[ií]odo Facturado Desde")
        if periodo:
            y_periodo = periodo.yc
        else:
            # invoices for products have no period: the issuer / receiver cut is just above the receiver's "CUIT:"
            cuits = sorted(s.yc for s in sp if re.match(r"CUIT:", s.texto))
            y_periodo = cuits[1] - 3 if len(cuits) > 1 else 175
        hdr_prod = buscar_label(sp, r"Producto\s*/\s*Servicio")
        y_items = hdr_prod.yc if hdr_prod else y_periodo + 80

        # --- type and number
        m = re.search(r"COD\.\s*(\d{1,3})", pg.texto)
        e.cod_comprobante = int(m.group(1)) if m else None
        e.letra, e.tipo = tipo_de_codigo(e.cod_comprobante)
        letra_caja = next((s.texto for s in sp if s.size >= 18 and s.texto in "ABCME" and len(s.texto) == 1
                           and s.y0 < y_periodo), "")
        if letra_caja and e.letra and letra_caja != e.letra:
            e.notas.append(f"La letra del recuadro ({letra_caja}) no coincide con el código {e.cod_comprobante}")
        e.letra = e.letra or letra_caja
        if re.search(r"NOTA DE CR[EÉ]DITO", pg.texto):
            e.tipo = "Nota de Crédito"
        elif re.search(r"NOTA DE D[EÉ]BITO", pg.texto):
            e.tipo = "Nota de Débito"
        e.pv = _int(valor_de(sp, r"Punto de Venta:", y_max=y_periodo))
        e.nro = _int(valor_de(sp, r"Comp\.?\s*Nro:?", y_max=y_periodo))

        # --- issuer (above "Período facturado")
        e.cuit = solo_digitos(valor_de(sp, r"CUIT:", y_max=y_periodo))
        e.razon_social = valor_de(sp, r"Raz[oó]n Social:", y_max=y_periodo)
        lab_rs = buscar_label(sp, r"Raz[oó]n Social:", y_max=y_periodo)
        if lab_rs:
            arriba = [s for s in sp if s.yc < lab_rs.yc - 10 and s.x1 < 280 and s.size >= 9.5
                      and not re.search(r"ORIGINAL|FACTURA|NOTA DE", s.texto)]
            if arriba:
                e.nombre_fantasia = re.sub(r"\s+", " ", arriba[0].texto).strip()
        e.domicilio = valor_de(sp, r"Domicilio Comercial:", y_max=y_periodo)
        lab_dom = buscar_label(sp, r"Domicilio Comercial:", y_max=y_periodo)
        if lab_dom and e.domicilio:
            val = next((s for s in sp if s.texto == e.domicilio and abs(s.yc - lab_dom.yc) <= 3.5), None)
            sigue = next((s for s in sp if val and 4 < s.yc - val.yc < 13 and abs(s.x0 - val.x0) < 3
                          and not s.texto.endswith(":")), None) if val else None
            if sigue:
                e.domicilio += " " + sigue.texto
        e.condicion_emisor = valor_de(sp, r"Condici[oó]n frente al IVA:", y_max=y_periodo)
        e.emision = parse_fecha(valor_de(sp, r"Fecha de Emisi[oó]n:", y_max=y_periodo))

        # --- period and due date
        e.desde = parse_fecha(valor_de(sp, r"Per[ií]odo Facturado Desde:"))
        e.vto = parse_fecha(valor_de(sp, r"Fecha de Vto\. para el pago:"))

        # --- receiver (between "Período" and the items header)
        e.cuit_receptor = solo_digitos(valor_de(sp, r"CUIT:", y_min=y_periodo + 3, y_max=y_items))
        e.condicion_receptor = valor_de(sp, r"Condici[oó]n frente al IVA:", y_min=y_periodo + 3, y_max=y_items)

        # --- totals
        tot = lambda pat: parse_monto(valor_de(sp, pat, y_min=y_items))
        e.total = tot(r"Importe Total:\s*\$?")
        e.neto = tot(r"Importe Neto Gravado:\s*\$?")
        sin_iva = [v for v in (tot(r"Importe Neto No Gravado:\s*\$?"), tot(r"Importe Exento:\s*\$?")) if v]
        e.no_gravado = suma(sin_iva) if sin_iva else None
        vals_otros = [parse_monto(valor_de(sp, r"Importe Otros Tributos:\s*\$?", y_min=s.yc - 1, y_max=s.yc + 1))
                      for s in sp if s.yc > y_items and s.texto.startswith("Importe Otros Tributos")]
        e.otros_tributos = max((v for v in vals_otros if v is not None), default=None)
        ivas = []
        for s in sp:
            if s.yc > y_items and re.match(r"IVA [\d.,]+\s*%:", s.texto):
                ivas.append(parse_monto(valor_de(sp, re.escape(s.texto), y_min=s.yc - 1, y_max=s.yc + 1)))
        e.iva = suma(ivas) if ivas else None
        if e.total is None:
            sub = tot(r"Subtotal:\s*\$?")
            e.total = sub

        # --- items
        e.items = _items(sp, hdr_prod, y_items, cfg)
        m = re.search(r"Fac\.?\s*([ABCM])\s*:\s*(\d{1,5})-(\d{1,8})", pg.texto)
        if m and e.tipo != "Factura":
            e.notas.append(f"Asociada a la Factura {m.group(1)} {int(m.group(2)):04d}-{int(m.group(3)):08d}")

        # --- AFIP download name wins for the key fields (it is what ARCA itself generated)
        afip = nombre_afip(nombre_archivo)
        if afip:
            cu, cod, pv, nro = afip
            for campo, v_nom, v_txt in (("CUIT", cu, e.cuit), ("punto de venta", pv, e.pv),
                                        ("número", nro, e.nro), ("código", cod, e.cod_comprobante)):
                if v_txt not in (None, "") and str(v_nom) != str(v_txt):
                    e.notas.append(f"El nombre del archivo dice {campo} {v_nom} y el PDF {v_txt}; uso el del nombre")
            e.cuit, e.pv, e.nro = cu, pv, nro
            if cod != e.cod_comprobante:
                e.cod_comprobante = cod
                e.letra, tipo = tipo_de_codigo(cod)
                e.tipo = tipo
        _completitud(e)
        return e


def _int(s: str) -> int | None:
    d = solo_digitos(s)
    return int(d) if d else None


def _items(sp: list[Span], hdr_prod: Span | None, y_items: float, cfg) -> list[Item]:
    if not hdr_prod:
        return []
    fin = min((s.yc for s in sp if s.yc > y_items + 4 and re.match(_TOTALES, s.texto)), default=1e9)
    hdr = [s for s in sp if abs(s.yc - hdr_prod.yc) <= 8]
    cant = next((s for s in hdr if s.texto.startswith("Cantidad")), None)
    x_desc_max = (cant.x0 - 2) if cant else 230
    x_cod_max = hdr_prod.x0 - 3
    zona = [s for s in sp if hdr_prod.yc + 6 < s.yc < fin - 2]

    # anchor rows: lines with at least two amounts right of the description column
    filas: list[list[Span]] = []
    for s in sorted((s for s in zona if s.x0 > x_desc_max), key=lambda s: s.yc):
        if filas and abs(filas[-1][0].yc - s.yc) <= 3:
            filas[-1].append(s)
        else:
            filas.append([s])
    anclas = [f for f in filas if sum(es_monto(s.texto) for s in f) >= 2]
    if not anclas:
        return []
    ys = [min(s.yc for s in f) for f in anclas]
    items = []
    for i, fila in enumerate(anclas):
        y0 = ys[i] - 5 if i else hdr_prod.yc + 6
        y1 = ys[i + 1] - 5 if i + 1 < len(ys) else fin
        textos = sorted((s for s in zona if y0 <= s.yc < y1 and s.x0 <= x_desc_max), key=lambda s: (round(s.yc), s.x0))
        codigo = " ".join(s.texto for s in textos if s.x1 <= x_cod_max + 1)
        desc = " ".join(s.texto for s in textos if s.x1 > x_cod_max + 1)
        fila = sorted(fila, key=lambda s: s.x0)
        montos = [s for s in fila if es_monto(s.texto)]
        alic = next((s for s in fila if re.fullmatch(r"\d+(?:[.,]\d+)?\s*%", s.texto)), None)
        bruto = parse_monto(montos[-1].texto)
        neto, iva = bruto, None
        if not alic and any(re.match(r"(?i)(exento|no gravado)", s.texto) for s in fila):
            iva = Decimal(0)                 # an exempt / not-taxed item carries no IVA
        if alic:
            antes = [s for s in montos if s.x1 <= alic.x0]
            if antes:
                neto = parse_monto(antes[-1].texto)
                iva = (bruto - neto) if bruto is not None and neto is not None else None
        tareas = tareas_en(f"{codigo} {desc}", cfg.tarea_min, cfg.tarea_max)
        items.append(Item(descripcion=re.sub(r"\s+", " ", desc).strip(), monto=neto, tareas=tareas,
                          codigo=codigo.strip(), iva=iva))
    return items


def _completitud(e: Extraccion) -> None:
    falt = []
    if not cuit_valido(e.cuit):
        falt.append("cuit")
    for campo in ("pv", "nro", "letra", "total", "vto"):
        if getattr(e, campo) in (None, ""):
            falt.append(campo)
    if not e.desde and not e.emision:
        falt.append("desde")
    if not e.items:
        falt.append("items")
    e.faltantes = falt
    e.confianza = max(0.0, 1.0 - 0.25 * len(falt))


PARSER = Arca()
