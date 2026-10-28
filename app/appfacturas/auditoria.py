"""Revisar: read-only audit of the workbook against the folders (like the PS1 -Revisar). Changes nothing."""
from __future__ import annotations

import time
from collections import defaultdict
from dataclasses import dataclass, field
from pathlib import Path
from urllib.parse import unquote

from . import maestro
from . import planilla as pl_mod
from .util import norm, relativa

DIAS_VIEJO = 7


@dataclass
class Hallazgo:
    titulo: str
    explicacion: str
    items: list[str] = field(default_factory=list)


def ruta_de_link(cfg, link: str) -> Path | None:
    """Local file a SharePoint link of column J points to (None if it is not under 'Facturas pendientes')."""
    if not link or not cfg.url_base or not cfg.base:
        return None
    base_url = unquote(cfg.url_base).rstrip("/").lower()
    url = unquote(link).strip()
    if not url.lower().startswith(base_url + "/"):
        return None
    return cfg.base.joinpath(*url[len(base_url) + 1:].split("/"))


def _donde(f) -> str:
    return f"{f.hoja} fila {f.fila} ({f.proveedor}{', ' + f.j_texto if f.j_texto else ''})"


def revisar(cfg, hoy: float | None = None) -> list[Hallazgo]:
    """Everything worth a look, in order. Empty lists are kept so the screen can say 'OK'."""
    hoy = hoy or time.time()
    pl = pl_mod.leer(cfg.planilla, cfg.areas)
    if not pl.leida:
        return [Hallazgo("No pude leer la planilla", "Cerrala en Excel si está bloqueada o revisá la ruta.", [pl.error])]
    sin_link = Hallazgo("Filas sin link", "Filas con proveedor pero sin link en la columna J.")
    rotos = Hallazgo("Links a archivos que no están", "El link de J apunta a un PDF que ya no está en su carpeta "
                     "(se movió, se renombró o se borró).")
    repetidas = Hallazgo("Filas repetidas", "Mismo link, o mismo proveedor con el mismo texto en J, en más de una fila.")
    pisado = Hallazgo("Filas con el IVA (H) escrito a mano o por la app",
                      "H tiene un número en lugar de la fórmula: Factura B o C (0), o una Factura A con IVA distinto. "
                      "Está bien si es eso; revisá las que no.")
    sin_fila = Hallazgo("PDF guardados que no están en la planilla", "Facturas en las carpetas de los proveedores "
                        "sin ninguna fila que las nombre. Puede ser algo ya pagado o algo que falta cargar.")
    sin_ficha = Hallazgo("Proveedores sin datos de pago", "Proveedores de la planilla que no están en la hoja "
                         "Proveedores ni en la memoria (PROVEEDORES RUA.xlsx).")
    viejos = Hallazgo(f"Archivos esperando hace más de {DIAS_VIEJO} días en _Entrada",
                      "Quedaron sin leer o sin confirmar.")

    linkeados: set[str] = set()
    vistos: dict[str, list] = defaultdict(list)
    for f in pl.filas:
        if not f.j_link and not f.j_texto:
            sin_link.items.append(_donde(f))
        if f.j_link:
            vistos["link:" + unquote(f.j_link).lower()].append(f)
            local = ruta_de_link(cfg, f.j_link)
            if local is not None:
                linkeados.add(str(local).lower())
                if not local.exists():
                    rotos.items.append(f"{_donde(f)} → {relativa(local, cfg.base)}")
        if f.j_texto:
            vistos["texto:" + norm(f.proveedor) + "|" + norm(f.j_texto)].append(f)
        if f.h_pisado is not None:
            pisado.items.append(f"{_donde(f)}: H = {f.h_pisado}")
    ya = set()
    for fs in vistos.values():
        clave = tuple(sorted((f.hoja, f.fila) for f in fs))
        if len(fs) > 1 and clave not in ya:
            ya.add(clave)
            repetidas.items.append(" y ".join(_donde(f) for f in fs))

    if cfg.base:
        for a in cfg.areas:
            d = cfg.base / a
            if d.is_dir():
                for p in sorted(d.rglob("*.pdf")):
                    if str(p).lower() not in linkeados:
                        sin_fila.items.append(relativa(p, cfg.base))
        for a in cfg.areas:
            d = cfg.entrada / a
            if d.is_dir():
                for p in sorted(d.iterdir()):
                    if p.is_file() and hoy - p.stat().st_mtime > DIAS_VIEJO * 86400:
                        viejos.items.append(f"{relativa(p, cfg.base)} ({int((hoy - p.stat().st_mtime) // 86400)} días)")

    con_ficha = {norm(p["nombre"]) for p in pl.proveedores if p.get("nombre")}
    try:
        con_ficha |= {norm(m["valores"][1]) for m in maestro.leer(cfg).values()}
    except Exception:
        pass
    for nombre in sorted({f.proveedor for f in pl.filas}, key=str.lower):
        if norm(nombre) not in con_ficha:
            sin_ficha.items.append(nombre)
    return [rotos, repetidas, sin_link, sin_fila, sin_ficha, pisado, viejos]


def informe(hallazgos: list[Hallazgo]) -> str:
    """Plain text for 'Copiar informe'."""
    lineas = []
    for h in hallazgos:
        lineas.append(f"{h.titulo}: {len(h.items) if h.items else 'OK'}")
        lineas += [f"  - {x}" for x in h.items]
    return "\n".join(lineas)
