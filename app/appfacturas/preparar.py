"""Setup (create the local and OneDrive folders + an empty workbook if missing) and workbook reset."""
from __future__ import annotations

import os
import shutil
from datetime import datetime
from pathlib import Path

from . import config
from .excel import PRIMERA, ULTIMA, ExcelError, abrir

AREAS = ("Vehicular", "Hogar")
SUBCARPETAS = ["_Entrada/Vehicular", "_Entrada/Hogar", "_Entrada/_Duplicados", "_Entrada/_Procesados",
               "Vehicular", "Hogar", "_Script"]
CSV_ENCABEZADO = "CUIT;Area;Proveedor;Carpeta\r\n"
NOMBRE_PLANILLA = "DEUDA A PAGAR ARG.xlsx"
H_R1C1 = '=IF(RC[-1]="","",ROUND(RC[-1]*R4C2,2))'          # H = G * B4 (IVA)
I_R1C1 = '=IF(RC[-2]="","",RC[-2]+RC[-1])'                 # I = G + H
COLS_PROVEEDORES = ["Área", "Proveedor", "CUIT", "Condición frente al IVA", "Domicilio", "Contacto (nombre)",
                    "Teléfono", "Email", "Banco", "Tipo de cuenta", "N° de cuenta", "CBU / CVU", "Alias",
                    "Titular de la cuenta", "Observaciones"]


# ------------------------------------------------------------------ where things go

def raiz_onedrive() -> Path | None:
    """The synced OneDrive for Business folder of RUA ('OneDrive - Rua Asistencia')."""
    cands = []
    if os.environ.get("OneDriveCommercial"):
        cands.append(Path(os.environ["OneDriveCommercial"]))
    home = Path(os.environ.get("USERPROFILE") or Path.home())
    if home.exists():
        todas = sorted(p for p in home.glob("OneDrive - *") if p.is_dir())
        cands += [p for p in todas if "rua" in p.name.lower()] + todas
    return next((c for c in cands if c.is_dir()), None)


def plan_setup(cfg) -> dict:
    """What setup would create. Never overwrites anything."""
    base = cfg.base
    raiz = base.parent if base else raiz_onedrive()
    if not base and raiz:
        base = raiz / "Facturas pendientes"
    crear_dirs, crear_archivos = [], []
    app = config.carpeta_app()
    for d in ("datos", "datos/lotes", "datos/backups", "datos/entrada_ui"):
        if not (app / d).is_dir():
            crear_dirs.append(app / d)
    if base:
        if not base.is_dir():
            crear_dirs.append(base)
        for s in SUBCARPETAS:
            if not (base / s).is_dir():
                crear_dirs.append(base / s)
        if not (base / "_Script" / "proveedores.csv").is_file():
            crear_archivos.append(base / "_Script" / "proveedores.csv")
    planilla = cfg.planilla or (raiz / NOMBRE_PLANILLA if raiz else None)
    if planilla and not planilla.is_file():
        crear_archivos.append(planilla)
    return {"base": base, "raiz": raiz, "planilla": planilla, "dirs": crear_dirs, "archivos": crear_archivos,
            "sin_onedrive": base is None}


def ejecutar_setup(cfg, plan: dict | None = None) -> list[str]:
    plan = plan or plan_setup(cfg)
    hecho = []
    for d in plan["dirs"]:
        d.mkdir(parents=True, exist_ok=True)
        hecho.append(f"Carpeta creada: {d}")
    for f in plan["archivos"]:
        if f.exists():
            continue
        f.parent.mkdir(parents=True, exist_ok=True)
        if f.name == "proveedores.csv":
            f.write_text(CSV_ENCABEZADO, encoding="utf-8-sig", newline="")
        elif f.suffix.lower() == ".xlsx":
            crear_planilla_vacia(f)
        hecho.append(f"Archivo creado: {f}")
    return hecho


# ------------------------------------------------------------------ new empty workbook

def crear_planilla_vacia(ruta: Path) -> Path:
    """DEUDA A PAGAR ARG layout: Vehicular/Hogar (headers row 11, data 12-210, H/I formulas, summary H4:I9),
    Proveedores (headers row 4, A-O) and Listas."""
    from openpyxl import Workbook
    from openpyxl.formatting.rule import FormulaRule
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.worksheet.datavalidation import DataValidation

    navy, naranja = "14147A", "FF542C"
    titulo = Font(bold=True, size=14, color=navy)
    enc = Font(bold=True, color="FFFFFF")
    relleno = PatternFill("solid", fgColor=navy)
    fino = Side(style="thin", color="D0D4E0")
    borde = Border(bottom=fino)
    wb = Workbook()
    wb.remove(wb.active)
    for i, area in enumerate(AREAS):
        ws = wb.create_sheet(area)
        ws["A1"] = f"Deuda a pagar · {area}"
        ws["A1"].font = titulo
        ws["A4"], ws["B4"] = "Alícuota IVA", 0.21
        ws["B4"].number_format = "0%"
        ws["A5"], ws["B5"] = "Hoy", "=TODAY()"
        ws["B5"].number_format = "dd/mm/yyyy"
        rng = lambda c: f"{c}{PRIMERA}:{c}{ULTIMA}"
        resumen = [("Total pendiente", f"=SUM({rng('G')})"), ("Total IVA", f"=SUM({rng('H')})"),
                   ("Total + IVA", f"=SUM({rng('I')})"), ("# de facturas", f"=COUNTA({rng('A')})"),
                   ("Vencido", f'=SUMIFS({rng("I")},{rng("F")},"<"&$B$5)'),
                   ("Vencen pronto (7 días)", f'=SUMIFS({rng("I")},{rng("F")},">="&$B$5,{rng("F")},"<="&$B$5+7)')]
        for j, (t, f) in enumerate(resumen):
            ws.cell(4 + j, 8, t).font = Font(bold=True, color=navy)
            c = ws.cell(4 + j, 9, f)
            c.number_format = "0" if t.startswith("#") else '"$" #,##0.00'
        heads = ["Proveedor", "N° de tarea", "Servicio realizado", "Empresa", "Fecha realizada",
                 "Fecha de vencimiento", "Monto total", "IVA", "Monto total + IVA", "Link a la factura"]
        for j, h in enumerate(heads, start=1):
            c = ws.cell(11, j, h)
            c.font, c.fill = enc, relleno
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.row_dimensions[11].height = 30
        for r in range(PRIMERA, ULTIMA + 1):
            ws.cell(r, 2).number_format = "@"
            for c in (5, 6):
                ws.cell(r, c).number_format = "dd/mm/yyyy"
            for c in (7, 8, 9):
                ws.cell(r, c).number_format = '"$" #,##0.00'
            ws.cell(r, 8, f'=IF(G{r}="","",ROUND(G{r}*$B$4,2))')
            ws.cell(r, 9, f'=IF(G{r}="","",G{r}+H{r})')
            for c in range(1, 11):
                ws.cell(r, c).border = borde
        for col, w in zip("ABCDEFGHIJ", (34, 13, 34, 18, 14, 14, 15, 13, 17, 30)):
            ws.column_dimensions[col].width = w
        ws.freeze_panes = "A12"
        ws.conditional_formatting.add(f"A{PRIMERA}:J{ULTIMA}", FormulaRule(
            formula=[f'AND($F{PRIMERA}<>"",$F{PRIMERA}<$B$5,$A{PRIMERA}<>"")'], fill=PatternFill("solid", fgColor="FDE2E1")))
        ws.conditional_formatting.add(f"A{PRIMERA}:J{ULTIMA}", FormulaRule(
            formula=[f'AND($F{PRIMERA}<>"",$F{PRIMERA}>=$B$5,$F{PRIMERA}<=$B$5+7,$A{PRIMERA}<>"")'],
            fill=PatternFill("solid", fgColor="FFF6D5")))
        dv = DataValidation(type="list", formula1=f"=Listas!${'AB'[i]}$2:${'AB'[i]}$50", allow_blank=True)
        ws.add_data_validation(dv)
        dv.add(f"D{PRIMERA}:D{ULTIMA}")
    wp = wb.create_sheet("Proveedores")
    wp["A1"] = "Proveedores"
    wp["A1"].font = titulo
    for j, h in enumerate(COLS_PROVEEDORES, start=1):
        c = wp.cell(4, j, h)
        c.font, c.fill = enc, relleno
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        wp.column_dimensions[c.column_letter].width = 16 if j not in (2, 5, 15) else 30
    for r in range(5, 301):
        for c in (3, 7, 11, 12):
            wp.cell(r, c).number_format = "@"
    wp.freeze_panes = "A5"
    wl = wb.create_sheet("Listas")
    wl["A1"], wl["B1"] = "Empresas Vehicular", "Empresas Hogar"
    for c in ("A1", "B1"):
        wl[c].font = Font(bold=True, color=naranja)
    wl.column_dimensions["A"].width = wl.column_dimensions["B"].width = 28
    ruta.parent.mkdir(parents=True, exist_ok=True)
    wb.save(ruta)
    return ruta


# ------------------------------------------------------------------ reset

def respaldo_reset(planilla: Path, backups: Path) -> list[Path]:
    """Two copies before a reset: next to the workbook (OneDrive keeps it) and in datos\\backups."""
    marca = f"{datetime.now():%Y-%m-%d %H%M}"
    junto = planilla.with_name(f"{planilla.stem} - respaldo antes del reset {marca}{planilla.suffix}")
    backups.mkdir(parents=True, exist_ok=True)
    local = backups / f"{datetime.now():%Y-%m-%d_%H%M%S} antes del reset {planilla.name}"
    shutil.copy2(planilla, junto)
    shutil.copy2(planilla, local)
    return [junto, local]


def reset_planilla(cfg, backups: Path, vaciar_proveedores: bool = False, log=print) -> dict:
    """Empties the workbook IN PLACE (same file, so the OneDrive sharing links keep working):
    A-G and J of rows 12-210 (values, fills, comments, links) and H back to its formula. I is never touched.
    Optionally empties Proveedores A5:O300. Returns what was done (with the backups)."""
    if not cfg.planilla or not Path(cfg.planilla).is_file():
        raise ExcelError(f"no encuentro la planilla ({cfg.planilla})")
    libro = abrir(cfg)
    try:
        if not libro.saved:
            raise ExcelError("la planilla está abierta con cambios sin guardar: guardala o cerrala y volvé a probar")
        copias = respaldo_reset(Path(cfg.planilla), backups)
        filas = {}
        for area in AREAS:
            h = libro.hoja(area)
            filas[area] = sum(1 for r in range(PRIMERA, ULTIMA + 1) if h.formula(r, "A") not in ("", None))
            ref = next((r for r in range(PRIMERA, ULTIMA + 1) if str(h.formula(r, "H")).startswith("=")), None)
            h.vaciar_rango(f"A{PRIMERA}:J{ULTIMA}", sin_columna="I", formula_h_desde=ref)
            log(f"{area}: {filas[area]} filas vaciadas")
        provs = 0
        if vaciar_proveedores:
            hp = libro.hoja("Proveedores")
            provs = sum(1 for r in range(5, 301) if hp.formula(r, "A") not in ("", None) or hp.formula(r, "B") not in ("", None))
            hp.limpiar("A5:O300")
            log(f"Proveedores: {provs} filas vaciadas")
        libro.guardar()
        return {"filas": filas, "proveedores": provs, "respaldos": [str(c) for c in copias]}
    finally:
        libro.cerrar()
