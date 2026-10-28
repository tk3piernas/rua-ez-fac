"""Provider memory: PROVEEDORES RUA.xlsx, next to the workbook in OneDrive.

The debt workbooks can be exported per batch, so the provider details (CBU, alias, contact...) live in their own
workbook: one row per provider, same columns A:O as the Proveedores sheet + P "Actualizado". The app reads it
to fill the Proveedores sheet and writes it (through Excel, like the debt workbook) when you paste a provider's
details or a batch brings a .txt. Every change goes to a history file, so it can be undone."""
from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path

from .util import norm, sha256, solo_digitos

NOMBRE = "PROVEEDORES RUA.xlsx"
HOJA = "Proveedores"
COLS = ["Área", "Proveedor", "CUIT", "Condición frente al IVA", "Domicilio", "Contacto (nombre)",
        "Teléfono", "Email", "Banco", "Tipo de cuenta", "N° de cuenta", "CBU / CVU", "Alias",
        "Titular de la cuenta", "Observaciones"]
LETRAS = "ABCDEFGHIJKLMNO"
TEXTO = set("CGKL")                   # CUIT, phone, account number and CBU go as text
PRIMERA, ULTIMA = 2, 3000


def ruta(cfg) -> Path | None:
    return Path(cfg.planilla).parent / NOMBRE if cfg.planilla else None


def clave(valores: list) -> str:
    """CUIT, or the normalised name for providers without CUIT."""
    cuit = solo_digitos(str(valores[2] or ""))
    return cuit or ("SC:" + norm(str(valores[1] or "")) if valores[1] else "")


def crear(r: Path) -> None:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    wb = Workbook()
    ws = wb.active
    ws.title = HOJA
    for j, h in enumerate(COLS + ["Actualizado"], start=1):
        c = ws.cell(1, j, h)
        c.font, c.fill = Font(bold=True, color="FFFFFF"), PatternFill("solid", fgColor="14147A")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[c.column_letter].width = 30 if h in ("Proveedor", "Domicilio", "Observaciones") else 17
    for f in range(PRIMERA, 501):
        for col in TEXTO:
            ws[f"{col}{f}"].number_format = "@"
    ws.freeze_panes = "C2"
    r.parent.mkdir(parents=True, exist_ok=True)
    wb.save(r)


def leer(cfg) -> dict[str, dict]:
    """clave -> {"fila", "valores" (15, as text)}. Empty if the file doesn't exist yet."""
    return leer_archivo(ruta(cfg))


def leer_archivo(r: Path | None) -> dict[str, dict]:
    if not r or not r.is_file():
        return {}
    from openpyxl import load_workbook
    wb = load_workbook(r, read_only=True, data_only=True)
    out = {}
    try:
        ws = wb[HOJA] if HOJA in wb.sheetnames else wb.worksheets[0]
        for f, fila in enumerate(ws.iter_rows(min_row=PRIMERA, max_col=15, values_only=True), start=PRIMERA):
            vals = [_texto(v) for v in fila] + [""] * (15 - len(fila))
            k = clave(vals)
            if k:
                out[k] = {"fila": f, "valores": vals}
    finally:
        wb.close()
    return out


def buscar(memoria: dict, cuit: str = "", nombre: str = "") -> dict | None:
    if cuit and cuit in memoria:
        return memoria[cuit]
    return memoria.get("SC:" + norm(nombre)) if nombre else None


def _texto(v) -> str:
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).strip()


def combinar(antes: list | None, nuevos: list, solo_vacios: bool = False) -> list:
    """New values over the old ones. Empty new values never erase; with solo_vacios only blanks get filled."""
    if not antes:
        return [_texto(v) for v in nuevos]
    out = list(antes)
    for i, v in enumerate(nuevos):
        v = _texto(v)
        if v and (not out[i] or not solo_vacios):
            out[i] = v
    return out


def desde_texto(texto: str, area: str = "") -> tuple[list, list[str]]:
    """Pasted text (mail, WhatsApp, .txt) -> (15 values for the memory, warnings)."""
    from .proveedores import nombre_prolijo
    from .txt_proveedor import desde_texto as leer_txt
    d = leer_txt(texto)
    c0 = d.cuentas[0] if d.cuentas else None
    obs = [f"Otra cuenta: {c.banco} CBU/CVU {c.cbu}" + (f" alias {c.alias}" if c.alias else "") for c in d.cuentas[1:]]
    vals = [area, d.nombre or nombre_prolijo(d.titular), d.cuit, d.condicion_iva, d.domicilio, d.contacto,
            d.telefono, d.email, c0.banco if c0 else "", c0.tipo if c0 else "", c0.nro if c0 else "",
            c0.cbu if c0 else "", c0.alias if c0 else "", (c0.titular if c0 else "") or d.titular, ". ".join(obs)]
    return [_texto(v) for v in vals], list(d.avisos)


def chequeos(vals: list) -> list[tuple[bool, str]]:
    """(ok, text) for the fields that can be checked: CUIT and CBU check digits, bank vs CBU."""
    from .bancos import banco_de
    from .util import cbu_valido, cuit_valido
    out = []
    cuit, cbu = solo_digitos(vals[2]), solo_digitos(vals[11])
    if not vals[1]:
        out.append((False, "Falta el nombre del proveedor"))
    if cuit:
        out.append((cuit_valido(cuit), f"CUIT {cuit}" + (" válido" if cuit_valido(cuit) else ": dígito verificador mal")))
    else:
        out.append((False, "Sin CUIT (se guarda por nombre)"))
    if cbu:
        ok = cbu_valido(cbu)
        banco = banco_de(cbu) if ok else ""
        out.append((ok, f"CBU/CVU {cbu}" + (f" válido ({banco})" if ok else ": dígito verificador mal")))
    else:
        out.append((False, "Sin CBU/CVU"))
    return out


def diferencias(antes: list, despues: list) -> list[tuple[str, str, str]]:
    return [(COLS[i], a, d) for i, (a, d) in enumerate(zip(antes, despues)) if a != d]


# ------------------------------------------------------------------ write and undo

def guardar(cfg, fichas: list[tuple[list, bool]], historial: Path, motivo: str, log=print) -> list[dict]:
    """Writes each (valores, solo_vacios) into the memory: same provider -> same row, else the first empty row.
    Returns the changes (also appended to `historial`)."""
    from . import excel
    r = ruta(cfg)
    if not r:
        raise excel.ExcelError("no sé dónde está la planilla, así que no sé dónde guardar la memoria de proveedores")
    creado = not r.is_file()
    if creado:
        crear(r)
    memoria = leer(cfg)
    libro = excel.abrir(cfg, r)
    cambios = []
    try:
        if not libro.saved:
            raise excel.LibroConCambios(f"{NOMBRE} está abierta con cambios sin guardar: guardala y volvé a probar")
        h = libro.hoja(HOJA)
        ocupadas = {m["fila"] for m in memoria.values()}
        for vals, solo_vacios in fichas:
            k = clave(vals)
            if not k:
                continue
            ya = memoria.get(k)
            nuevos = combinar(ya["valores"] if ya else None, vals, solo_vacios)
            if ya and nuevos == ya["valores"]:
                continue
            fila = ya["fila"] if ya else next(f for f in range(PRIMERA, ULTIMA + 1)
                                              if f not in ocupadas and not h.formula(f, "B") and not h.formula(f, "C"))
            for i, v in enumerate(nuevos):
                h.poner(fila, LETRAS[i], v, texto=LETRAS[i] in TEXTO)
            h.poner(fila, "P", datetime.now().strftime("%d/%m/%Y %H:%M"))
            ocupadas.add(fila)
            memoria[k] = {"fila": fila, "valores": nuevos}
            cambios.append({"fila": fila, "antes": ya["valores"] if ya else None, "despues": nuevos})
            log(f"MEMORIA fila {fila}: {nuevos[1]} ({'actualizado' if ya else 'nuevo'})")
        if cambios:
            libro.guardar()
    finally:
        libro.cerrar()
    if cambios or creado:
        hist = _leer_hist(historial)
        hist.append({"id": datetime.now().strftime("%Y%m%d%H%M%S%f"), "fecha": datetime.now().isoformat(timespec="seconds"),
                     "motivo": motivo, "ruta": str(r), "creado": creado, "sha256": sha256(r), "cambios": cambios,
                     "deshecho": False})
        historial.parent.mkdir(parents=True, exist_ok=True)
        historial.write_text(json.dumps(hist, ensure_ascii=False, indent=1), encoding="utf-8")
    return cambios


def _leer_hist(historial: Path) -> list[dict]:
    try:
        return json.loads(historial.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return []


def ultimo(historial: Path) -> dict | None:
    return next((e for e in reversed(_leer_hist(historial)) if not e["deshecho"] and e["cambios"]), None)


def ultimo_de(historiales: list[Path]) -> tuple[Path, dict] | None:
    """The newest change not undone among several histories (the one of pasted data and one per batch)."""
    todos = [(h, e) for h in historiales for e in _leer_hist(h) if not e["deshecho"] and e["cambios"]]
    return max(todos, key=lambda he: he[1]["id"], default=None)


def deshacer(cfg, historial: Path, todo: bool = False, log=print, id_: str = "") -> list[str]:
    """Undoes the last change of `historial` (or all of them, newest first; or the one with `id_`). Only rows
    still as the app left them go back; if the app created the file and nobody touched it since, it is deleted.
    Works on the file saved in the history, whatever mode (test / real) the app is in now."""
    from . import excel
    hist = _leer_hist(historial)
    pendientes = [e for e in reversed(hist) if not e["deshecho"]]
    if id_:
        pendientes = [e for e in pendientes if e["id"] == id_]
    elif not todo:
        pendientes = pendientes[:1]
    problemas: list[str] = []
    for e in pendientes:
        r = Path(e["ruta"])
        if not r.is_file():
            problemas.append(f"{r.name} ya no está; no hay nada que deshacer")
        elif e["creado"] and sha256(r) == e["sha256"]:
            r.unlink()
            log(f"BORRO  {r.name} (la había creado la app)")
        elif e["cambios"]:
            memoria = {m["fila"]: m["valores"] for m in leer_archivo(r).values()}
            libro = excel.abrir(cfg, r)
            try:
                if not libro.saved:
                    raise excel.LibroConCambios(f"{r.name} está abierta con cambios sin guardar: guardala y volvé a deshacer")
                h = libro.hoja(HOJA)
                for c in reversed(e["cambios"]):
                    if memoria.get(c["fila"]) != c["despues"]:
                        problemas.append(f"{r.name} fila {c['fila']} cambió después; no la toco")
                        continue
                    for i, v in enumerate(c["antes"] or [""] * 15):
                        h.poner(c["fila"], LETRAS[i], v, texto=LETRAS[i] in TEXTO)
                    h.poner(c["fila"], "P", "")
                    log(f"MEMORIA fila {c['fila']} vuelve a como estaba")
                libro.guardar()
            finally:
                libro.cerrar()
            if e["creado"] and not leer_archivo(r):
                r.unlink()
                log(f"BORRO  {r.name} (la había creado la app y quedó vacía)")
        e["deshecho"] = True
        # saved after each change: if a later one fails, the ones already undone are not tried again
        historial.write_text(json.dumps(hist, ensure_ascii=False, indent=1), encoding="utf-8")
    return problemas
