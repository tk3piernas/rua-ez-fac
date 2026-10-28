"""Setup and workbook reset (no real invoices needed: this file also goes to the repo)."""
from pathlib import Path

import pytest
from openpyxl import load_workbook
from openpyxl.comments import Comment
from openpyxl.styles import PatternFill

from appfacturas import planilla, preparar
from appfacturas.config import Config


def _cfg(tmp_path):
    base = tmp_path / "OneDrive - Rua Asistencia" / "Facturas pendientes"
    return Config(base=base, planilla=base.parent / "DEUDA A PAGAR ARG.xlsx", excel_motor="openpyxl",
                  url_base="https://rua.example/Documents/Facturas pendientes")


def test_setup_crea_todo_y_no_pisa(tmp_path, monkeypatch):
    monkeypatch.setenv("APPFACTURAS_HOME", str(tmp_path / "app"))
    cfg = _cfg(tmp_path)
    cfg.base.parent.mkdir(parents=True)
    plan = preparar.plan_setup(cfg)
    assert plan["planilla_nueva"] == cfg.planilla and cfg.planilla not in plan["archivos"] and cfg.base in plan["dirs"]
    assert plan["onedrive_sin_bajar"]                      # empty OneDrive folder: may still be syncing
    preparar.ejecutar_setup(cfg, plan)
    assert not cfg.planilla.exists()                       # A9: never without its own confirmation
    hecho = preparar.ejecutar_setup(cfg, plan, crear_planilla=True)
    assert hecho
    for s in preparar.SUBCARPETAS:
        assert (cfg.base / s).is_dir()
    csv = cfg.base / "_Script" / "proveedores.csv"
    assert csv.read_text(encoding="utf-8-sig").startswith("CUIT;Area;Proveedor;Carpeta")
    assert (tmp_path / "app" / "datos" / "lotes").is_dir()
    pl = planilla.leer(cfg.planilla)
    assert pl.leida and pl.filas == [] and float(pl.alicuota["Vehicular"]) == 0.21
    # second run: nothing to do, nothing overwritten
    csv.write_text("CUIT;Area;Proveedor;Carpeta\r\n1;Hogar;X;X\r\n", encoding="utf-8-sig")
    plan2 = preparar.plan_setup(cfg)
    assert plan2["dirs"] == [] and plan2["archivos"] == []
    assert "1;Hogar;X" in csv.read_text(encoding="utf-8-sig")


def test_reset_vacia_en_el_lugar_y_respalda(tmp_path):
    cfg = _cfg(tmp_path)
    preparar.crear_planilla_vacia(cfg.planilla)
    wb = load_workbook(cfg.planilla)
    ws = wb["Vehicular"]
    ws["A12"], ws["B12"], ws["G12"], ws["H12"] = "Remolques X", "412345", 1000, 0        # H overwritten (Factura C)
    ws["J12"] = '=HYPERLINK("https://x/a.pdf","Factura 0001-00000001")'
    ws["A13"], ws["G13"] = "Otro", 500
    ws["J13"].hyperlink = "https://x/b.pdf"
    ws["A13"].comment = Comment("posible duplicado", "x")
    for c in "ABCDEFGHIJ":
        ws[f"{c}13"].fill = PatternFill("solid", fgColor="FFFFF2CC")
    wb["Proveedores"]["A5"], wb["Proveedores"]["B5"] = "Vehicular", "Remolques X"
    wb.save(cfg.planilla)

    res = preparar.reset_planilla(cfg, tmp_path / "backups", vaciar_proveedores=False, log=lambda s: None)
    assert res["filas"] == {"Vehicular": 2, "Hogar": 0, "Persona": 0, "ServiciosYa": 0}
    assert len(res["respaldos"]) == 2 and all(__import__("pathlib").Path(p).is_file() for p in res["respaldos"])

    wb = load_workbook(cfg.planilla)
    ws = wb["Vehicular"]
    for c in "ABCDEFGJ":
        assert ws[f"{c}12"].value is None and ws[f"{c}13"].value is None
    assert ws["H12"].value == '=IF(G12="","",ROUND(G12*$B$4,2))'      # formula back
    assert ws["I12"].value == '=IF(G12="","",G12+H12)'                 # I untouched
    assert ws["A13"].comment is None and ws["J13"].hyperlink is None
    assert ws["A13"].fill.fill_type is None
    assert ws["B4"].value == 0.21 and ws["A11"].value == "Proveedor"   # layout kept
    assert wb["Proveedores"]["B5"].value == "Remolques X"              # kept unless asked

    preparar.reset_planilla(cfg, tmp_path / "backups", vaciar_proveedores=True, log=lambda s: None)
    assert load_workbook(cfg.planilla)["Proveedores"]["B5"].value is None


def test_agrega_hojas_de_areas_nuevas(tmp_path):
    """A workbook from before Persona / ServiciosYa gets them as empty copies of Hogar; the rest is kept."""
    cfg = _cfg(tmp_path)
    preparar.crear_planilla_vacia(cfg.planilla)
    wb = load_workbook(cfg.planilla)
    del wb["Persona"], wb["ServiciosYa"]
    wb["Hogar"]["A12"], wb["Hogar"]["G12"] = "Plomero X", 1000
    wb.save(cfg.planilla)
    assert preparar.hojas_faltantes(cfg.planilla) == ["Persona", "ServiciosYa"]
    assert preparar.plan_setup(cfg)["hojas"] == ["Persona", "ServiciosYa"]

    res = preparar.agregar_hojas(cfg, tmp_path / "backups", log=lambda s: None)
    assert res["hojas"] == ["Persona", "ServiciosYa"] and len(res["respaldos"]) == 2
    wb = load_workbook(cfg.planilla)
    assert wb.sheetnames[:4] == ["Vehicular", "Hogar", "Persona", "ServiciosYa"]
    ws = wb["Persona"]
    assert ws["A1"].value == "Deuda a pagar · Persona" and ws["A11"].value == "Proveedor"
    assert ws["A12"].value is None and ws["H12"].value == '=IF(G12="","",ROUND(G12*$B$4,2))'
    assert [d.formula1 for d in ws.data_validations.dataValidation] == ["=Listas!$C$2:$C$50"]
    assert wb["Hogar"]["A12"].value == "Plomero X"
    assert preparar.hojas_faltantes(cfg.planilla) == []


def test_exportar_copia_con_nombre_y_deshacer(tmp_path):
    from appfacturas.util import sha256
    cfg = _cfg(tmp_path)
    preparar.crear_planilla_vacia(cfg.planilla)
    destino = preparar.exportar_planilla(cfg, "Deuda 28/10 tanda 3", log=lambda s: None)
    assert destino == cfg.planilla.parent / "Planillas exportadas" / "Deuda 28-10 tanda 3.xlsx"
    assert sha256(destino) == sha256(cfg.planilla)
    try:
        preparar.exportar_planilla(cfg, "Deuda 28-10 tanda 3.xlsx", log=lambda s: None)
        assert False, "no debería pisar"
    except preparar.ExcelError as e:
        assert "ya hay" in str(e)
    assert preparar.borrar_exportacion(destino, sha256(destino)) == "" and not destino.exists()


def test_deshacer_dos_veces_no_toca_nada(tmp_path):
    import json
    import pytest
    from appfacturas import ejecutor
    lote = tmp_path / "L1"
    lote.mkdir()
    (lote / "estado.json").write_text(json.dumps({"estado": "deshecho", "problemas_deshacer": []}), encoding="utf-8")
    with pytest.raises(ejecutor.YaDeshecho):
        ejecutor.deshacer(lote, log=lambda s: None)
    assert json.loads((lote / "estado.json").read_text(encoding="utf-8"))["problemas_deshacer"] == []


def test_onedrive_sin_bajar(tmp_path):
    assert preparar.onedrive_sin_bajar(tmp_path / "no existe")
    assert preparar.onedrive_sin_bajar(tmp_path) == "la carpeta del OneDrive está vacía"
    (tmp_path / "otro archivo.docx").write_bytes(b"x")
    assert preparar.onedrive_sin_bajar(tmp_path) == ""


def test_volver_del_reset(tmp_path):
    """A10: the reset can be undone while nobody changed the workbook; the same file is written in place."""
    cfg = _cfg(tmp_path)
    preparar.crear_planilla_vacia(cfg.planilla)
    wb = load_workbook(cfg.planilla)
    wb["Vehicular"]["A12"], wb["Vehicular"]["G12"] = "Remolques X", 1000
    wb.save(cfg.planilla)
    backups = tmp_path / "datos" / "backups"
    res = preparar.reset_planilla(cfg, backups, log=lambda s: None)
    assert Path(res["respaldos"][0]).parent == cfg.planilla.parent / "Respaldos"     # A12
    assert preparar.reset_para_volver(cfg, backups)
    preparar.volver_del_reset(cfg, backups, log=lambda s: None)
    assert load_workbook(cfg.planilla)["Vehicular"]["A12"].value == "Remolques X"
    assert preparar.reset_para_volver(cfg, backups) is None                          # only once


def test_no_vuelve_si_la_planilla_cambio_despues(tmp_path):
    cfg = _cfg(tmp_path)
    preparar.crear_planilla_vacia(cfg.planilla)
    backups = tmp_path / "datos" / "backups"
    preparar.reset_planilla(cfg, backups, log=lambda s: None)
    wb = load_workbook(cfg.planilla)
    wb["Vehicular"]["A12"] = "cargada después del reset"
    wb.save(cfg.planilla)
    assert preparar.reset_para_volver(cfg, backups) is None
    with pytest.raises(preparar.ExcelError):
        preparar.volver_del_reset(cfg, backups)


def test_respaldos_locales_quedan_20(tmp_path):
    """A12: datos\backups keeps only the newest 20 copies."""
    from appfacturas import excel
    backups = tmp_path / "backups"
    backups.mkdir()
    for i in range(25):
        (backups / f"2026-01-{i + 1:02d}_000000 DEUDA A PAGAR ARG.xlsx").write_bytes(b"x")
    pl = tmp_path / "DEUDA A PAGAR ARG.xlsx"
    pl.write_bytes(b"y")
    excel.respaldar(pl, backups)
    quedan = sorted(p.name for p in backups.glob("*.xlsx"))
    assert len(quedan) == 20 and quedan[0].startswith("2026-01-07")
