"""Plain-text report of a batch (same information as the PS1 -Simular output) + parity CSV."""
from __future__ import annotations

import csv
import json
from decimal import Decimal
from pathlib import Path

from .model import Estado, Lote, Nivel, a_json
from .util import fmt_fecha, fmt_monto, relativa


def texto(lote: Lote, cfg) -> str:
    L = []
    w = L.append
    w(f"Facturas pendientes: {cfg.base}")
    if lote.simular:
        w("*** MODO SIMULACIÓN: no se mueve ni se escribe nada ***")
    c = lote.contexto
    w(f"Proveedores en la lista: {c.get('proveedores_csv', 0)}")
    w(f"Facturas ya guardadas en carpetas: {c.get('guardadas', 0)}")
    w(f"Planilla: {c.get('planilla', '')}")
    w(f"Filas: una por {'ítem' if lote.modo_filas == 'item' else 'factura'}")
    w("")
    base = cfg.base
    for f in lote.facturas:
        a = f.archivo
        if f.estado == Estado.NO_LEIDA:
            continue
        if f.dup and f.a_duplicados and not f.existente:
            w(f"DUP  {a.area}\\{a.nombre}: {f.dup}; va a _Entrada\\_Duplicados.")
        elif f.existente:
            w(f"DUP  {a.area}\\{a.nombre} ya estaba guardada como {_rel(f.existente, base)}; la copia va a _Entrada\\_Duplicados.")
        elif f.destino:
            w(f"PDF  {a.area}\\{a.nombre}  ->  {_rel(f.destino, base)}" + ("   [proveedor nuevo]" if f.proveedor_nuevo else ""))
    for p in lote.plan:
        if p["op"] == "mover_zip":
            w(f"ZIP  {Path(p['de']).name}  ->  {_rel(Path(p['a']), base)}")
    for t in lote.txts:
        w(f"TXT  {t['area']}\\{t['nombre']}  ->  {t['proveedor']}  (CUIT {t['cuit'] or '?'}, {t['banco'] or 'sin CBU'})"
          + (f"  [{t['asociado']}]" if t['asociado'] else ""))
    if lote.pendientes:
        w("")
        w("Quedarían en _Entrada (revisar):")
        for p in lote.pendientes:
            w(f"  - {p}")
    w("")
    w("FILAS QUE SE AGREGARÍAN EN LA PLANILLA:")
    for f in lote.facturas:
        if f.dup and f.estado == Estado.DUPLICADA and f.clave:
            w(f"  NO SE CARGA: {f.proveedor} | {f.texto_link} {f.dup}")
    por_hoja: dict[str, Decimal] = {}
    for r in lote.filas:
        iva = "IVA según fórmula" if r.iva is None else f"IVA {fmt_monto(r.iva)}"
        w(f"  {r.hoja}: {r.proveedor} | {r.texto_link} | {'tarea ' + r.tarea if r.tarea else 'sin tarea'} | "
          f"desde {fmt_fecha(r.fecha)} | vence {fmt_fecha(r.vto)} | {fmt_monto(r.monto)} | {iva}")
        w(f"      servicio: {r.servicio}")
        if r.resaltar:
            w(f"      OJO, posible duplicado (se cargaría resaltada en amarillo)")
        por_hoja[r.hoja] = por_hoja.get(r.hoja, Decimal(0)) + (r.monto or 0)
    w("")
    w(f"Total del lote: {fmt_monto(lote.total)}  " + "  ".join(f"{h}: {fmt_monto(v)}" for h, v in por_hoja.items()))
    nuevos = [p for p in lote.plan if p["op"] == "proveedor_nuevo"]
    if nuevos:
        w("")
        w("PROVEEDORES NUEVOS (se agregarían a proveedores.csv y se crearía la carpeta):")
        for p in nuevos:
            w(f"  {p['cuit']};{p['area']};{p['nombre']};{p['carpeta']}")
    avisos = [a for a in lote.avisos if a.nivel != Nivel.INFO]
    info = [a for a in lote.avisos if a.nivel == Nivel.INFO]
    if avisos:
        w("")
        w("AVISOS:")
        for a in avisos:
            w(f"  [{a.nivel.value}] {a.texto}")
    if info:
        w("")
        w("Para tener en cuenta:")
        for a in info:
            w(f"  - {a.texto}")
    if lote.proveedores:
        w("")
        w("FILAS PARA LA HOJA PROVEEDORES (A:O):")
        for r in lote.proveedores:
            w("  " + " | ".join(str(v) for v in r["valores"]))
    if not lote.facturas and not lote.txts:
        w("No había nada nuevo en _Entrada.")
    return "\n".join(L)


def _rel(p: Path, base) -> str:
    return relativa(p, base)


def paridad_csv(lote: Lote, ruta: Path) -> None:
    """One line per invoice, the fields compared against the PS1."""
    with open(ruta, "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.writer(fh, delimiter=";")
        w.writerow(["archivo", "clave", "proveedor", "destino", "desde", "vto", "monto", "iva", "tareas", "estado", "dup", "posibles"])
        for f in lote.facturas:
            e = f.ext
            tareas = " / ".join(dict.fromkeys(t for i in (e.items if e else []) for t in i.tareas))
            w.writerow([f.archivo.nombre, f.clave, f.proveedor,
                        _rel(f.destino or f.existente, None) if (f.destino or f.existente) else "",
                        fmt_fecha(e.desde or e.emision) if e else "", fmt_fecha(e.vto) if e else "",
                        f.monto if f.monto is not None else "", "" if f.iva is None else f.iva,
                        tareas, f.estado.value, f.dup, " | ".join(f.posibles)])


def guardar(lote: Lote, carpeta: Path, informe: str) -> Path:
    d = carpeta / lote.id
    d.mkdir(parents=True, exist_ok=True)
    (d / "informe.txt").write_text(informe, encoding="utf-8")
    (d / "lote.json").write_text(json.dumps(lote, default=a_json, ensure_ascii=False, indent=1), encoding="utf-8")
    paridad_csv(lote, d / "paridad.csv")
    return d
