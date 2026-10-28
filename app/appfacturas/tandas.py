"""Closing a round of payments ("tanda"): export the workbook as 'DEUDA A PAGAR ARG tanda N', note which
batches went into it and, optionally, empty the workbook to start the next one. Can be undone."""
from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path

from . import ejecutor, preparar
from .excel import ExcelError
from .util import sha256

ARCHIVO = "tandas.json"


def leer(datos: Path) -> list[dict]:
    try:
        return json.loads((datos / ARCHIVO).read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return []


def _guardar(datos: Path, tandas: list[dict]) -> None:
    datos.mkdir(parents=True, exist_ok=True)
    from .util import guardar_json
    guardar_json(datos / ARCHIVO, tandas)


def _de_la_planilla(tandas: list[dict], planilla) -> list[dict]:
    return [t for t in tandas if str(t["planilla"]).lower() == str(planilla).lower() and not t.get("deshecha")]


def siguiente(datos: Path, planilla) -> int:
    return len(_de_la_planilla(leer(datos), planilla)) + 1


def lotes_abiertos(datos: Path, planilla) -> list[str]:
    """Confirmed batches loaded in this workbook that are not in a closed round yet (oldest first)."""
    cerrados = {l for t in _de_la_planilla(leer(datos), planilla) for l in t["lotes"]}
    out = []
    for d in ejecutor.lotes_confirmados(datos / "lotes"):
        if d.name in cerrados:
            continue
        pl = ejecutor.donde(d)[1]
        if pl and str(pl).lower() == str(planilla).lower():
            out.append(d.name)
    return out


def tanda_de(datos: Path) -> dict[str, int]:
    """batch id -> round number (for Historial)."""
    return {l: t["n"] for t in leer(datos) if not t.get("deshecha") for l in t["lotes"]}


def cerrar(cfg, datos: Path, nombre: str = "", vaciar: bool = False, log=print) -> dict:
    """Exports the workbook as round N and notes its batches; with `vaciar`, then the reset (one local backup)."""
    n = siguiente(datos, cfg.planilla)
    destino = preparar.exportar_planilla(cfg, nombre or f"DEUDA A PAGAR ARG tanda {n}", log)
    t = {"n": n, "fecha": datetime.now().isoformat(timespec="seconds"), "planilla": str(cfg.planilla),
         "exportada": str(destino), "sha256": sha256(destino), "lotes": lotes_abiertos(datos, cfg.planilla),
         "vaciada": False, "deshecha": False}
    tandas = leer(datos)
    tandas.append(t)
    _guardar(datos, tandas)                  # noted before the reset: if that fails the export is still known
    if vaciar:
        res = preparar.reset_planilla(cfg, datos / "backups", log=log, solo_local=True)
        t["vaciada"], t["filas"] = True, sum(res["filas"].values())
        _guardar(datos, tandas)
    log(f"TANDA {n} cerrada: {destino.name}, {len(t['lotes'])} lotes")
    return t


def ultima(datos: Path, planilla) -> dict | None:
    ts = _de_la_planilla(leer(datos), planilla)
    return ts[-1] if ts else None


def deshacer(cfg, datos: Path, log=print) -> list[str]:
    """Undoes the last round of this workbook: the workbook back as before the reset (if it was emptied and
    nobody changed it since), the exported copy deleted (if nobody changed it), and the round forgotten."""
    t = ultima(datos, cfg.planilla)
    if not t:
        raise ExcelError("no hay ninguna tanda cerrada para deshacer")
    problemas = []
    if t.get("vaciada"):
        preparar.volver_del_reset(cfg, datos / "backups", log)       # raises if it can't: nothing else is touched
    p = preparar.borrar_exportacion(Path(t["exportada"]), t["sha256"])
    if p:
        problemas.append(p)
    tandas = leer(datos)
    for x in tandas:
        if x["n"] == t["n"] and x["planilla"] == t["planilla"] and not x.get("deshecha"):
            x["deshecha"] = True
    _guardar(datos, tandas)
    log(f"TANDA {t['n']} deshecha")
    return problemas
