"""Interfaz (NiceGUI) en una pestaña local del navegador: http://127.0.0.1:<puerto>.

La lógica vive en el núcleo (lote, ejecutor, excel); esto solo muestra, edita y confirma."""
from __future__ import annotations

import socket
import sys
import webbrowser
from datetime import date
from pathlib import Path

from nicegui import app, run, ui

from . import NOMBRE, __version__, config, ejecutor, excel, reporte
from .lote import Armador, Entrada, aplicar_exclusiones, entradas_de_carpeta
from .model import Estado, Nivel
from .util import ahora_id, fmt_fecha, fmt_monto, parse_fecha, parse_monto

CSS = """
:root { --navy: #14147a; --naranja: #ff542c; --linea: #e3e6f0; }
body { background: linear-gradient(180deg, #e9ebf7 0px, #f6f7fc 320px) fixed; color: #1d2030;
       font-family: "Segoe UI", Inter, system-ui, sans-serif; }
.ez-header { background: linear-gradient(90deg, #0d0d52, #14147a 55%, #26259c) !important;
             border-bottom: 3px solid var(--naranja); box-shadow: 0 4px 18px rgba(20,20,122,.25); }
.ez-logo { background: #fff; border-radius: 50%; padding: 3px; }
.seccion { background: #fff; border: 1px solid var(--linea); border-radius: 18px; position: relative; overflow: hidden;
           box-shadow: 0 8px 28px rgba(20,20,122,.07) !important; }
.seccion::before { content: ""; position: absolute; left: 0; right: 0; top: 0; height: 4px;
                   background: linear-gradient(90deg, var(--navy), var(--naranja)); }
.chip-icono { width: 38px; height: 38px; border-radius: 12px; background: rgba(255,84,44,.12); color: var(--naranja);
              display: flex; align-items: center; justify-content: center; }
.titulo-seccion { font-size: 1.12rem; font-weight: 700; color: var(--navy); }
.tarjeta { min-width: 160px; border-radius: 14px; background: linear-gradient(180deg, #fff, #f5f6fc);
           border: 1px solid var(--linea); box-shadow: none !important; padding: 12px 18px; }
.tarjeta .titulo { font-size: .75rem; color: #6b7088; text-transform: uppercase; letter-spacing: .4px; }
.tarjeta .valor { color: var(--navy); font-size: 1.3rem; font-weight: 800; }
.q-btn { border-radius: 10px; text-transform: none; font-weight: 600; letter-spacing: .2px; }
.q-tab { text-transform: none; font-weight: 600; color: #4a4f6a; }
.q-tab__label { font-size: .95rem; }
.q-tab--active { color: var(--navy); }
.q-tabs .q-tab__indicator { background: var(--naranja); height: 3px; border-radius: 3px; }
.pill { border-radius: 999px; padding: 4px 14px; font-size: .8rem; font-weight: 700; }
.pill-prueba { background: #dcfce7; color: #166534; }
.pill-real { background: #ffe4dc; color: #b42d0e; }
.pill-sim { background: #fef3c7; color: #92400e; }
.q-uploader { border-radius: 14px; }
.q-uploader__header { background: linear-gradient(90deg, var(--navy), #2b2aa8); }
.ag-theme-balham { --ag-header-background-color: #f1f2fa; --ag-header-foreground-color: #14147a;
                   --ag-odd-row-background-color: #fafbfe; --ag-border-radius: 10px; --ag-font-size: 13px; }
.fila-amarilla { background-color: #fff6d5 !important; }
.fila-roja { background-color: #fde2e1 !important; }
.fila-violeta { background-color: #efe7fb !important; }
.fila-excluida { color: #9aa0a6 !important; text-decoration: line-through; }
.contenido { width: 100%; max-width: 1600px; margin: 0 auto; padding-bottom: 140px; }
.alfa { position: fixed; right: 18px; bottom: 18px; z-index: 2000; width: 300px; border-radius: 18px; background: #fff;
        border: 1px solid #f6d5c8; box-shadow: 0 12px 32px rgba(20,20,122,.18) !important; }
.chequeo-ok { color: #15803d; } .chequeo-falta { color: #dc2626; } .chequeo-aviso { color: #d97706; }
"""


def seccion(titulo: str, icono: str | None = None, ancho: str = "w-full"):
    """A white card with a coloured top line and its title centered: each area of the screen is one of these."""
    c = ui.card().classes(f"seccion {ancho} items-center gap-3 p-6")
    with c:
        with ui.row().classes("w-full justify-center items-center gap-3"):
            if icono:
                with ui.element("div").classes("chip-icono"):
                    ui.icon(icono).classes("text-2xl")
            ui.label(titulo).classes("titulo-seccion")
        ui.separator().classes("w-full opacity-60")
    return c


def tarjeta(titulo: str, valor: str) -> None:
    with ui.card().classes("tarjeta items-center gap-1"):
        ui.label(titulo).classes("titulo")
        ui.label(valor).classes("valor")


class Estado_:
    """Single user, single window: the state lives in the process."""
    cfg: config.Config | None = None
    usar_prueba: bool = False
    lote = None
    trabajo: Path | None = None
    resultado: dict | None = None
    log: list[str] = []


S = Estado_()


# ------------------------------------------------------------------ paths

def datos() -> Path:
    return config.carpeta_app() / "datos"


def carpeta_subidos() -> Path:
    return datos() / "entrada_ui"


def base_activa(cfg) -> Path | None:
    if S.usar_prueba:
        from .paridad import carpeta_prueba
        return carpeta_prueba() / "Facturas pendientes"
    return config.cargar().base


def cfg_activa() -> config.Config:
    base = base_activa(None)
    return config.cargar(base=base)


def puede_escribir(cfg) -> bool:
    return S.usar_prueba or cfg.modo_real


def contar_entrada(cfg) -> dict:
    out = {}
    for a in cfg.areas:
        d = cfg.entrada / a if cfg.base else None
        out[a] = len([f for f in d.iterdir() if f.is_file()]) if d and d.is_dir() else 0
    sub = carpeta_subidos()
    out["subidos"] = len([f for f in sub.rglob("*") if f.is_file()]) if sub.exists() else 0
    return out


def entradas(cfg):
    ents, txts = entradas_de_carpeta(cfg)
    sub = carpeta_subidos()
    if sub.exists():
        for d in sorted(sub.iterdir()):
            area = d.name if d.name in cfg.areas else ""
            for f in sorted(d.iterdir(), key=lambda p: p.name.lower()):
                s = f.suffix.lower()
                if s in (".pdf", ".zip"):
                    ents.append(Entrada(f, area, "elegida" if area else ""))
                elif s == ".txt":
                    txts.append((f, area or "Vehicular"))
    return ents, txts


def armar(cfg):
    lid = ahora_id()
    trabajo = datos() / "lotes" / lid
    ents, txts = entradas(cfg)
    lote = Armador(cfg, simular=False, trabajo=trabajo).procesar(ents, txts)
    reporte.guardar(lote, trabajo.parent, reporte.texto(lote, cfg))
    return lote, trabajo


# ------------------------------------------------------------------ grid rows

def filas_grid(lote) -> list[dict]:
    bloq = {a.ref for a in lote.avisos if a.nivel == Nivel.BLOQUEA and a.ref}
    fact = {f.id: f for f in lote.facturas}
    out = []
    for i, r in enumerate(lote.filas):
        f = fact.get(r.factura_id)
        e = f.ext if f else None
        out.append({
            "idx": i, "incluir": not r.excluida, "area": r.hoja, "proveedor": r.proveedor,
            "factura": r.texto_link, "fecha": fmt_fecha(r.fecha), "vto": fmt_fecha(r.vto), "tarea": r.tarea,
            "servicio": r.servicio, "monto": float(r.monto) if r.monto is not None else None,
            "iva": "fórmula" if r.iva is None else f"{r.iva:.2f}".replace(".", ","),
            "fuente": (e.fuente if e else "") + ("" if not e or e.fiscal else " · no fiscal"),
            "nuevo": "sí" if f and f.proveedor_nuevo else "", "resaltar": r.resaltar,
            "bloquea": r.factura_id in bloq, "claude": bool(e and e.fuente.startswith("claude")),
            "vencida": bool(r.vto and r.vto < date.today()),
        })
    return out


COLUMNAS = [
    {"field": "incluir", "headerName": "Cargar", "width": 85, "editable": True, "cellDataType": "boolean"},
    {"field": "area", "headerName": "Área", "width": 95},
    {"field": "proveedor", "headerName": "Proveedor", "editable": True, "minWidth": 170, "flex": 1},
    {"field": "factura", "headerName": "Factura", "width": 190},
    {"field": "fecha", "headerName": "Fecha servicio", "editable": True, "width": 120},
    {"field": "vto", "headerName": "Vencimiento", "editable": True, "width": 120,
     "cellClassRules": {"text-red-700 font-bold": "data.vencida"}},
    {"field": "tarea", "headerName": "N° de tarea", "editable": True, "width": 130},
    {"field": "servicio", "headerName": "Servicio", "editable": True, "minWidth": 200, "flex": 2},
    {"field": "monto", "headerName": "Monto", "editable": True, "width": 130, "type": "rightAligned",
     ":valueFormatter": "p => p.value == null ? '' : p.value.toLocaleString('es-AR', {minimumFractionDigits: 2, maximumFractionDigits: 2})"},
    {"field": "iva", "headerName": "IVA (H)", "width": 100, "type": "rightAligned"},
    {"field": "fuente", "headerName": "Leída con", "width": 120},
    {"field": "nuevo", "headerName": "Prov. nuevo", "width": 105},
]


def aplicar_edicion(lote, args: dict) -> str | None:
    """Applies one grid edit to the batch. Returns an error message, or None."""
    data = args.get("data") or {}
    col = args.get("colId")
    i = data.get("idx")
    if i is None or i >= len(lote.filas):
        return None
    r = lote.filas[i]
    v = data.get(col)
    if col == "incluir":
        r.excluida = not bool(v)
        return None
    r.editada = True
    if col == "proveedor":
        r.proveedor = str(v or "").strip()
    elif col == "tarea":
        r.tarea = str(v or "").strip()
    elif col == "servicio":
        r.servicio = str(v or "").strip()
    elif col in ("fecha", "vto"):
        d = parse_fecha(str(v or ""))
        if v and not d:
            return f"No entiendo la fecha '{v}' (usá dd/mm/aaaa)"
        setattr(r, col, d)
    elif col == "monto":
        m = parse_monto(str(v)) if v not in (None, "") else None
        if v not in (None, "") and m is None:
            return f"No entiendo el monto '{v}'"
        r.monto = m
    return None


# ------------------------------------------------------------------ page

STATIC = Path(__file__).parent / "static"
app.add_static_files("/ezfac", STATIC)

@ui.page("/")
def pagina():
    ui.add_css(CSS)
    ui.colors(primary="#14147a", secondary="#ff542c", accent="#ff542c", positive="#15803d",
              negative="#dc2626", warning="#d97706", info="#2b2aa8")
    S.cfg = cfg_activa()
    with ui.header().classes("ez-header items-center justify-between px-6 py-2"):
        with ui.row().classes("items-center gap-3"):
            ui.image("/ezfac/logo.png").classes("ez-logo w-10 h-10")
            with ui.column().classes("gap-0"):
                ui.label(NOMBRE).classes("text-xl font-bold leading-tight")
                ui.label(f"Carga de facturas de proveedores · v{__version__}").classes("text-xs opacity-75")
        lbl_modo = ui.label().classes("pill")
        ui.button("Salir", icon="logout", on_click=lambda: (ui.notify("Cerrando la app…"), _salir())).props("flat color=white")

    def refrescar_modo():
        cfg = S.cfg
        if S.usar_prueba:
            lbl_modo.text, tipo = "Trabajando sobre la CARPETA DE PRUEBA", "pill-prueba"
        elif puede_escribir(cfg):
            lbl_modo.text, tipo = "Trabajando sobre carpetas y planilla REALES", "pill-real"
        else:
            lbl_modo.text, tipo = "REAL · solo simulación (escritura real apagada)", "pill-sim"
        lbl_modo.classes(replace=f"pill {tipo}")

    with ui.column().classes("contenido items-center gap-4 px-4"):
        with ui.tabs().props("align=center").classes("w-full") as tabs:
            t_cargar = ui.tab("Cargar", icon="upload_file")
            t_rev = ui.tab("Revisión", icon="fact_check")
            t_res = ui.tab("Resultado", icon="task_alt")
            t_hist = ui.tab("Historial", icon="history")
            t_conf = ui.tab("Configuración", icon="settings")
        with ui.tab_panels(tabs, value=t_cargar).classes("w-full bg-transparent"):
            with ui.tab_panel(t_cargar).classes("items-center"):
                panel_cargar(tabs, t_rev, refrescar_modo)
            with ui.tab_panel(t_rev).classes("items-center"):
                panel_revision(tabs, t_res)
            with ui.tab_panel(t_res).classes("items-center"):
                panel_resultado()
            with ui.tab_panel(t_hist).classes("items-center"):
                panel_historial()
            with ui.tab_panel(t_conf).classes("items-center"):
                panel_config(refrescar_modo)
    refrescar_modo()
    aviso_alfa()


def aviso_alfa() -> None:
    with ui.card().classes("alfa p-3") as tarjeta_alfa:
        with ui.row().classes("w-full items-center no-wrap gap-3"):
            ui.image("/ezfac/gatito.svg").classes("w-20 h-20 shrink-0")
            with ui.column().classes("gap-0 grow"):
                ui.label("Esto es una alfa").classes("font-bold text-base").style("color: var(--navy)")
                ui.label("Mandar bugs o sugerencias a Matheo").classes("text-sm text-gray-600")
            ui.button(icon="close", on_click=lambda: tarjeta_alfa.set_visibility(False)).props(
                "flat round dense size=sm color=grey-7").classes("self-start")


# ------------------------------------------------------------------ Cargar

def panel_cargar(tabs, t_rev, refrescar_modo):
    with ui.row().classes("w-full justify-center items-stretch gap-6"):
        with seccion("1. Archivos", "upload_file", "w-[420px]"):
            area = ui.select({"": "Detectar por proveedor", "Vehicular": "Vehicular", "Hogar": "Hogar"},
                             value="", label="Área de lo que subís").classes("w-72")

            async def subido(e):
                dest = carpeta_subidos() / (area.value or "_Detectar")
                dest.mkdir(parents=True, exist_ok=True)
                await e.file.save(dest / Path(e.file.name).name)
                contar.refresh()

            ui.upload(multiple=True, auto_upload=True, on_upload=subido,
                      label="Arrastrá PDF, zip o .txt (o hacé clic)").props('accept=".pdf,.zip,.txt"').classes("w-full")
            ui.label("Se copian a la app; los originales quedan donde estaban.").classes("text-xs text-gray-500 text-center")

        with seccion("2. Qué se va a leer", "folder_open", "w-[340px]"):
            @ui.refreshable
            def contar():
                c = contar_entrada(S.cfg) if S.cfg.base else {}
                with ui.column().classes("items-center gap-1"):
                    for etiqueta, n in ((r"_Entrada\Vehicular", c.get("Vehicular", 0)),
                                        (r"_Entrada\Hogar", c.get("Hogar", 0)),
                                        ("Subidos acá", c.get("subidos", 0))):
                        with ui.row().classes("gap-2 items-baseline"):
                            ui.label(etiqueta).classes("text-sm text-gray-600")
                            ui.label(str(n)).classes("text-xl font-bold").style("color: var(--navy)")
            contar()

            def vaciar_subidos():
                import shutil
                shutil.rmtree(carpeta_subidos(), ignore_errors=True)
                contar.refresh()
            ui.button("Quitar lo subido", icon="delete_sweep", on_click=vaciar_subidos).props("flat dense")

        with seccion("3. Leer", "search", "w-[300px]"):
            spinner = ui.spinner(size="lg")
            spinner.visible = False

            async def leer():
                spinner.visible = True
                try:
                    S.lote, S.trabajo = await run.io_bound(armar, S.cfg)
                    S.resultado = None
                except Exception as ex:
                    ui.notify(f"No pude leer: {ex}", type="negative", multi_line=True)
                    return
                finally:
                    spinner.visible = False
                ui.notify(f"Leí {len(S.lote.facturas)} facturas y {len(S.lote.txts)} .txt", type="positive")
                panel_revision.refresh()
                panel_historial.refresh()
                tabs.set_value(t_rev)
            ui.button("Leer facturas", icon="search", on_click=leer).props("size=lg color=secondary unelevated").classes("px-6")
            ui.label("Leer no mueve ni escribe nada.").classes("text-xs text-gray-500 text-center")
    ui.timer(5.0, lambda: contar.refresh())


# ------------------------------------------------------------------ Revisión

@ui.refreshable
def panel_revision(tabs=None, t_res=None):
    lote = S.lote
    if not lote:
        with seccion("Revisión", "fact_check", "w-[560px]"):
            ui.label("Todavía no hay un lote leído. Andá a Cargar y tocá 'Leer facturas'.").classes("text-gray-500 text-center")
        return
    cfg = S.cfg
    niveles = {n: [a for a in lote.avisos if a.nivel == n] for n in Nivel}

    with ui.column().classes("w-full items-center gap-5"):
        # --- summary
        with seccion("Resumen del lote", "summarize"):
            @ui.refreshable
            def resumen():
                act = [r for r in lote.filas if not r.excluida]
                por_area = {}
                for r in act:
                    por_area[r.hoja] = por_area.get(r.hoja, 0) + (r.monto or 0)
                with ui.row().classes("w-full justify-center gap-4"):
                    tarjetas = [("Facturas", str(len(lote.facturas))), ("Filas a cargar", str(len(act))),
                                ("Total del lote", fmt_monto(sum((r.monto or 0 for r in act), 0)))]
                    tarjetas += [(f"Total {a}", fmt_monto(v)) for a, v in por_area.items()]
                    tarjetas += [("Avisos", f"{len(niveles[Nivel.BLOQUEA])} frenan · {len(niveles[Nivel.REVISAR])} revisar")]
                    for titulo, valor in tarjetas:
                        tarjeta(titulo, valor)
            resumen()

        # --- warnings
        if any(niveles.values()):
            with seccion("Avisos", "warning"):
                if niveles[Nivel.BLOQUEA]:
                    with ui.expansion(f"Frenan la carga ({len(niveles[Nivel.BLOQUEA])})", icon="block", value=True).classes("w-full bg-red-50 rounded"):
                        for a in niveles[Nivel.BLOQUEA]:
                            ui.label("• " + a.texto).classes("text-sm")
                if niveles[Nivel.REVISAR]:
                    with ui.expansion(f"Para revisar ({len(niveles[Nivel.REVISAR])})", icon="report",
                                      value=len(niveles[Nivel.REVISAR]) <= 12).classes("w-full bg-amber-50 rounded"):
                        for a in niveles[Nivel.REVISAR]:
                            ui.label("• " + a.texto).classes("text-sm")
                if niveles[Nivel.INFO]:
                    with ui.expansion(f"Para tener en cuenta ({len(niveles[Nivel.INFO])})", icon="info").classes("w-full rounded"):
                        for a in niveles[Nivel.INFO]:
                            ui.label("• " + a.texto).classes("text-sm")

        # --- rows
        with seccion("Filas que se cargan en la planilla", "table_rows"):
            ui.label("Doble clic para editar Proveedor, fechas, N° de tarea, Servicio o Monto; destildá 'Cargar' para dejar "
                     "una fila afuera. Amarillo = posible duplicado · rojo = frenada · violeta = leída por Claude."
                     ).classes("text-xs text-gray-500 text-center")
            filas = filas_grid(lote)
            grid = ui.aggrid({
                "columnDefs": COLUMNAS, "rowData": filas,
                "stopEditingWhenCellsLoseFocus": True, "singleClickEdit": False,
                "rowClassRules": {"fila-amarilla": "data.resaltar", "fila-roja": "data.bloquea",
                                  "fila-violeta": "data.claude", "fila-excluida": "!data.incluir"},
                "defaultColDef": {"resizable": True, "sortable": True},
                "autoSizeStrategy": {"type": "fitGridWidth"},
            }, theme="balham").classes("w-full").style(f"height: {min(620, 64 + 28 * len(filas))}px")

            def editado(e):
                err = aplicar_edicion(lote, e.args)
                if err:
                    ui.notify(err, type="warning")
                resumen.refresh()
                total_filas.refresh()
            grid.on("cellValueChanged", editado)

            @ui.refreshable
            def total_filas():
                act = [r for r in lote.filas if not r.excluida]
                ui.label(f"Se cargarían {len(act)} filas por {fmt_monto(sum((r.monto or 0 for r in act), 0))}").classes("font-bold")
            total_filas()

        # --- what is not loaded / new providers / files
        no_cargan = [f for f in lote.facturas if f.estado in (Estado.DUPLICADA, Estado.NO_LEIDA) or
                     (f.dup and f.estado == Estado.REVISAR)]
        if no_cargan or lote.pendientes:
            with seccion(f"No se cargan ({len(no_cargan)})", "do_not_disturb_on"):
                with ui.column().classes("w-full gap-1"):
                    for f in no_cargan:
                        motivo = f.dup or ("no la pude leer" if f.estado == Estado.NO_LEIDA else "")
                        ui.label(f"• {f.archivo.nombre}: {motivo}").classes("text-sm")
                    for p in lote.pendientes:
                        ui.label(f"• {p}").classes("text-sm")
        if lote.proveedores:
            with seccion(f"Proveedores nuevos para la hoja Proveedores ({len(lote.proveedores)})", "person_add"):
                ui.aggrid({"columnDefs": [{"field": c, "headerName": h, "width": w} for c, h, w in
                                          [("area", "Área", 90), ("nombre", "Proveedor", 220), ("cuit", "CUIT", 120),
                                           ("banco", "Banco", 150), ("cbu", "CBU / CVU", 200), ("alias", "Alias", 150),
                                           ("obs", "Observaciones", 400)]],
                           "autoSizeStrategy": {"type": "fitGridWidth"},
                           "rowData": [{"area": v[0], "nombre": v[1], "cuit": v[2], "banco": v[8], "cbu": v[11],
                                        "alias": v[12], "obs": v[14]} for v in (p["valores"] for p in lote.proveedores)]},
                          theme="balham").classes("w-full").style(f"height: {64 + 28 * len(lote.proveedores)}px")
        movs = [p for p in lote.plan if p["op"] in ("mover", "copiar", "mover_zip", "a_duplicados", "txt_procesado")]
        if movs:
            with seccion(f"Archivos que se mueven ({len(movs)})", "drive_file_move"):
                with ui.expansion("Ver la lista", icon="list").classes("w-full"):
                    for p in movs:
                        a = p.get("a") or "_Entrada\\_Duplicados"
                        ui.label(f"• {Path(p['de']).name}  →  {_rel(a, cfg)}").classes("text-sm")

        # --- confirm
        with seccion("Confirmar", "check_circle", "w-[760px]"):
            with ui.row().classes("justify-center items-center gap-4"):
                btn = ui.button("Confirmar y cargar", icon="check", on_click=lambda: confirmar(tabs, t_res)).props("size=lg color=positive unelevated").classes("px-6")
                if not puede_escribir(cfg):
                    btn.disable()
                    btn.tooltip("La escritura en la planilla real todavía no está habilitada (Configuración)")
                ui.button("Copiar filas al portapapeles", icon="content_copy", on_click=copiar).props("outline")
            if not puede_escribir(cfg):
                ui.label("Sobre lo real, por ahora la app solo simula: habilitalo en Configuración.").classes("text-xs text-amber-700")
            ui.label(f"Informe guardado en datos\\lotes\\{lote.id}").classes("text-xs text-gray-500")


def _rel(p, cfg) -> str:
    try:
        return str(Path(p).relative_to(cfg.base))
    except Exception:
        return str(p)


async def copiar():
    try:
        await run.io_bound(excel.copiar_portapapeles, S.lote)
        ui.notify("Filas copiadas. En Excel: clic en la columna A de la primera fila vacía y Ctrl+V.", type="positive")
    except Exception as ex:
        ui.notify(f"No pude copiar: {ex}", type="negative")


async def confirmar(tabs, t_res):
    lote, cfg = S.lote, S.cfg
    excluidas = aplicar_exclusiones(lote)
    act = [r for r in lote.filas if not r.excluida]
    movs = len([p for p in lote.plan if p["op"] in ("mover", "copiar", "mover_zip", "a_duplicados", "txt_procesado")])
    with ui.dialog() as dlg, ui.card().classes("items-center"):
        ui.label("¿Confirmo?").classes("text-lg font-bold")
        ui.label(f"Se mueven {movs} archivos y se escriben {len(act)} filas "
                 f"y {len(lote.proveedores)} proveedores en {'la COPIA de prueba' if S.usar_prueba else 'la planilla REAL'}.").classes("text-center")
        if excluidas:
            ui.label(f"{len(excluidas)} facturas excluidas quedan en _Entrada.")
        with ui.row().classes("justify-center"):
            ui.button("Sí, cargar", on_click=lambda: dlg.submit(True)).props("color=positive")
            ui.button("Cancelar", on_click=lambda: dlg.submit(False)).props("flat")
    if not await dlg:
        return
    n = ui.notification("Cargando…", spinner=True, timeout=None)
    try:
        S.resultado = await run.io_bound(lambda: ejecutor.Ejecutor(cfg, lote, S.trabajo, log=lambda s: None).ejecutar())
    finally:
        n.dismiss()
    if S.resultado["excel"] == "pendiente" and "cambios sin guardar" in S.resultado["excel_mensaje"]:
        with ui.dialog() as d2, ui.card().classes("items-center"):
            ui.label("La planilla está abierta con cambios sin guardar.").classes("font-bold")
            ui.label("Los archivos ya se movieron. ¿Guardo tus cambios y escribo las filas, o las dejo pendientes?").classes("text-center")
            with ui.row().classes("justify-center"):
                ui.button("Guardar mis cambios y escribir", on_click=lambda: d2.submit(True)).props("color=positive")
                ui.button("Dejar pendientes", on_click=lambda: d2.submit(False)).props("flat")
        if await d2:
            est, msg = await run.io_bound(ejecutor.reintentar_excel, cfg, S.trabajo, True, lambda s: None)
            S.resultado["excel"], S.resultado["excel_mensaje"] = est, msg
    panel_resultado.refresh()
    panel_historial.refresh()
    tabs.set_value(t_res)


# ------------------------------------------------------------------ Resultado

@ui.refreshable
def panel_resultado():
    r = S.resultado
    if not r or not S.trabajo:
        with seccion("Resultado", "task_alt", "w-[640px]"):
            ui.label("Todavía no se confirmó ningún lote en esta sesión. "
                     "El último lote se puede deshacer desde Historial.").classes("text-gray-500 text-center")
        return
    est = ejecutor.estado(S.trabajo)
    ok = not r["errores"] and r["excel"] in ("escrito", "no corresponde")
    xl = est.get("excel_escrito") or {}
    with ui.column().classes("w-full items-center gap-5"):
        with seccion("Listo" if ok else "Listo, con cosas pendientes", "task_alt" if ok else "pending_actions", "w-[760px]"):
            with ui.row().classes("justify-center gap-4"):
                for titulo, valor in (("Pasos con archivos", str(len(r["hechos"]))), ("Errores", str(len(r["errores"]))),
                                      ("Filas en la planilla", str(xl.get("filas", 0))),
                                      ("Proveedores", str(xl.get("proveedores", 0)))):
                    tarjeta(titulo, valor)
            ui.label(f"Planilla: {r['excel']} {r.get('excel_mensaje', '')}").classes(
                "text-center " + ("text-green-700" if r["excel"] == "escrito" else "text-amber-700"))
            for e in r["errores"]:
                ui.label("! " + e).classes("text-red-700 text-sm")
        with seccion("Qué se hizo", "list", "w-[760px]"):
            with ui.expansion("Ver el detalle", icon="list").classes("w-full"):
                for h in r["hechos"]:
                    ui.label("• " + h).classes("text-sm")
        with seccion("Acciones", "build", "w-[760px]"):
            with ui.row().classes("justify-center gap-4"):
                if est.get("excel") == "pendiente":
                    async def reintentar():
                        e2, msg = await run.io_bound(ejecutor.reintentar_excel, S.cfg, S.trabajo, False, lambda s: None)
                        S.resultado["excel"], S.resultado["excel_mensaje"] = e2, msg
                        ui.notify("Filas escritas" if e2 == "escrito" else f"Siguen pendientes: {msg}",
                                  type="positive" if e2 == "escrito" else "warning")
                        panel_resultado.refresh()
                    ui.button("Reintentar planilla", icon="refresh", on_click=reintentar)
                    ui.button("Copiar filas al portapapeles", icon="content_copy", on_click=copiar).props("outline")
                ui.button("Deshacer este lote", icon="undo", on_click=lambda: deshacer(S.trabajo)).props("color=negative outline")


async def deshacer(trabajo: Path):
    with ui.dialog() as dlg, ui.card().classes("items-center"):
        ui.label(f"¿Deshago el lote {trabajo.name}?").classes("font-bold")
        ui.label("Borra las filas que agregó en la planilla y devuelve los archivos a donde estaban "
                 "(solo lo que sigue igual a como lo dejó la app).").classes("text-center")
        with ui.row().classes("justify-center"):
            ui.button("Sí, deshacer", on_click=lambda: dlg.submit(True)).props("color=negative")
            ui.button("Cancelar", on_click=lambda: dlg.submit(False)).props("flat")
    if not await dlg:
        return
    try:
        problemas = await run.io_bound(lambda: ejecutor.deshacer(trabajo, log=lambda s: None, cfg=S.cfg))
    except Exception as ex:
        ui.notify(f"No pude deshacer (no se cambió nada): {ex}", type="negative", multi_line=True)
        return
    if problemas:
        ui.notify("Deshecho, salvo: " + "; ".join(problemas), type="warning", multi_line=True, timeout=0, close_button=True)
    else:
        ui.notify("Lote deshecho: todo volvió a como estaba.", type="positive")
    S.resultado = None
    panel_resultado.refresh()
    panel_historial.refresh()


# ------------------------------------------------------------------ Historial

@ui.refreshable
def panel_historial():
    d = datos() / "lotes"
    lotes = sorted((x for x in d.iterdir() if x.is_dir()), reverse=True) if d.exists() else []
    with seccion("Lotes anteriores", "history", "w-[760px]"):
        if not lotes:
            ui.label("Todavía no hay lotes.").classes("text-gray-500")
            return
        ultimo = ejecutor.ultimo_confirmado(d)
        filas = []
        for x in lotes[:40]:
            e = ejecutor.estado(x)
            filas.append({"lote": x.name, "estado": e.get("estado", "solo leído / simulado"),
                          "planilla": e.get("excel", ""), "errores": len(e.get("errores", []))})
        ui.aggrid({"columnDefs": [{"field": "lote", "headerName": "Lote", "width": 180},
                                  {"field": "estado", "headerName": "Estado", "width": 220},
                                  {"field": "planilla", "headerName": "Planilla", "width": 150},
                                  {"field": "errores", "headerName": "Errores", "width": 100}],
                   "autoSizeStrategy": {"type": "fitGridWidth"}, "rowData": filas}, theme="balham").classes("w-full").style(f"height: {min(600, 64 + 28 * len(filas))}px")
        if ultimo:
            ui.button(f"Deshacer el último lote confirmado ({ultimo.name})", icon="undo",
                      on_click=lambda: deshacer(ultimo)).props("color=negative outline")


# ------------------------------------------------------------------ Configuración

def panel_config(refrescar_modo):
    cfg_real = config.cargar()
    with ui.column().classes("w-full items-center gap-5"):
        with seccion("Setup", "construction", "w-[760px]"):
            ui.label("Crea las carpetas que falten (en la app y en tu OneDrive de RUA), una planilla vacía si no "
                     "tenés, y revisa que esté todo bien. Nunca pisa nada que ya exista.").classes("text-sm text-gray-600 text-center")
            ui.button("Revisar y preparar todo", icon="construction", on_click=dialogo_setup).props(
                "color=secondary unelevated size=md").classes("px-5")

        with seccion("Dónde trabaja la app", "folder", "w-[760px]"):
            with ui.column().classes("items-center gap-1"):
                ui.label(f"Facturas pendientes (real): {cfg_real.base}").classes("text-sm text-center")
                ui.label(f"Planilla (real): {cfg_real.planilla}").classes("text-sm text-center")
                ui.label(f"Configuración: {cfg_real.archivo}").classes("text-xs text-gray-500 text-center")

            def cambiar_prueba(e):
                S.usar_prueba = bool(e.value)
                S.cfg = cfg_activa()
                S.lote = None
                panel_revision.refresh()
                refrescar_modo()
                ui.notify("Ahora trabajás sobre la " + ("carpeta de PRUEBA" if S.usar_prueba else "carpeta REAL"))
            ui.switch("Trabajar sobre la carpeta de prueba (carpeta de la app\\prueba)", value=S.usar_prueba,
                      on_change=cambiar_prueba)

            async def rearmar():
                from . import paridad
                await run.io_bound(paridad.preparar)
                ui.notify("Carpeta de prueba rearmada con una copia de la planilla y el lote de prueba", type="positive")
            ui.button("Rearmar la carpeta de prueba", icon="restart_alt", on_click=rearmar).props("outline")

        with seccion("Escritura en la planilla y las carpetas reales", "lock_open", "w-[760px]"):
            ui.label("Mientras esté apagado, sobre lo real la app solo lee y simula.").classes("text-xs text-gray-500")

            def cambiar_real(e):
                config.guardar_opcion(cfg_real.archivo, "app", "modo_real", bool(e.value))
                S.cfg = cfg_activa()
                refrescar_modo()
                panel_revision.refresh()
            ui.switch("Habilitar escritura real", value=cfg_real.modo_real, on_change=cambiar_real)

        with seccion("Empezar la planilla de cero", "restart_alt", "w-[760px]"):
            ui.label("Deja vacías las filas 12 a 210 de Vehicular y Hogar: valores, colores, comentarios y links. "
                     "Vuelve a poner la fórmula del IVA en H, no toca la columna I ni el formato, y es el mismo "
                     "archivo (los links compartidos siguen andando). Antes guarda dos copias de respaldo.").classes(
                "text-sm text-gray-600 text-center")
            ui.button("Reset de la planilla…", icon="delete_forever", on_click=dialogo_reset).props(
                "color=negative outline")

        with seccion("Cómo se carga", "tune", "w-[760px]"):
            def cambiar_filas(e):
                config.guardar_opcion(cfg_real.archivo, "reglas", "filas", e.value)
                S.cfg = cfg_activa()
                ui.notify("Se aplica al próximo lote que leas")
            ui.select({"item": "Una fila por ítem", "factura": "Una fila por factura (como el script viejo)"},
                      value=cfg_real.filas, label="Filas", on_change=cambiar_filas).classes("w-80")

            def cambiar_servicio(e):
                config.guardar_opcion(cfg_real.archivo, "planilla", "escribir_servicio", bool(e.value))
                S.cfg = cfg_activa()
            ui.switch("Escribir el servicio de la factura en la columna C", value=cfg_real.escribir_servicio,
                      on_change=cambiar_servicio)


def mostrar_chequeo(lineas: list[str]) -> None:
    with ui.column().classes("w-full gap-1"):
        for l in lineas:
            if not l.strip():
                continue
            for pref, icono, clase in (("OK", "check_circle", "chequeo-ok"), ("FALTA", "cancel", "chequeo-falta"),
                                       ("AVISO", "warning", "chequeo-aviso")):
                if l.startswith(pref):
                    with ui.row().classes("items-start no-wrap gap-2"):
                        ui.icon(icono).classes(f"{clase} text-lg")
                        ui.label(l[len(pref):].strip()).classes("text-sm break-all")
                    break
            else:
                ui.label(l).classes("text-sm font-bold" if l.startswith(("Todo", "Hay")) else "text-xs text-gray-500")


async def dialogo_setup():
    from . import instalacion, preparar
    cfg_real = config.cargar()
    plan = await run.io_bound(preparar.plan_setup, cfg_real)
    ok, lineas = await run.io_bound(instalacion.chequeo)
    with ui.dialog() as dlg, ui.card().classes("seccion w-[760px] max-w-full gap-3 p-6"):
        ui.label("Setup").classes("titulo-seccion")

        @ui.refreshable
        def cuerpo():
            if plan["sin_onedrive"]:
                ui.label("No encuentro el OneDrive de RUA sincronizado en esta PC (carpeta 'OneDrive - Rua Asistencia'). "
                         "Abrí OneDrive, iniciá sesión con tu cuenta de RUA y volvé a tocar Setup.").classes("text-red-700 text-sm")
            pend = plan["dirs"] + plan["archivos"]
            if pend:
                ui.label(f"Voy a crear {len(pend)} cosas:").classes("font-bold")
                with ui.column().classes("gap-0 w-full"):
                    for x in pend:
                        tipo = "archivo" if x in plan["archivos"] else "carpeta"
                        ui.label(f"• {tipo}: {x}").classes("text-xs break-all")
            else:
                ui.label("No falta crear nada.").classes("font-bold chequeo-ok")
            ui.separator()
            ui.label("Chequeo").classes("font-bold")
            mostrar_chequeo(lineas)
        cuerpo()

        async def crear():
            nonlocal plan, ok, lineas
            try:
                hecho = await run.io_bound(preparar.ejecutar_setup, cfg_real, plan)
            except Exception as ex:
                ui.notify(f"No pude crear todo: {ex}", type="negative", multi_line=True)
                return
            ui.notify(f"Listo: {len(hecho)} cosas creadas", type="positive")
            plan = await run.io_bound(preparar.plan_setup, config.cargar())
            ok, lineas = await run.io_bound(instalacion.chequeo)
            S.cfg = cfg_activa()
            cuerpo.refresh()
            btn.set_visibility(bool(plan["dirs"] or plan["archivos"]))

        with ui.row().classes("w-full justify-center gap-3"):
            btn = ui.button("Crear lo que falta", icon="add_circle", on_click=crear).props("color=secondary unelevated")
            btn.set_visibility(bool(plan["dirs"] or plan["archivos"]))
            ui.button("Cerrar", on_click=dlg.close).props("flat")
    dlg.open()


async def dialogo_reset():
    from . import preparar
    cfg = S.cfg
    if not puede_escribir(cfg):
        ui.notify("Para vaciar la planilla real primero habilitá la escritura real (o usá la carpeta de prueba).",
                  type="warning", multi_line=True)
        return
    with ui.dialog() as dlg, ui.card().classes("seccion w-[620px] max-w-full items-center gap-3 p-6"):
        ui.label("Reset de la planilla").classes("titulo-seccion")
        ui.label(f"{'COPIA DE PRUEBA' if S.usar_prueba else 'PLANILLA REAL'}: {cfg.planilla}").classes(
            "text-xs text-center break-all " + ("chequeo-ok" if S.usar_prueba else "chequeo-falta"))
        ui.label("Se vacían las filas cargadas de Vehicular y Hogar. Antes se guardan dos copias: una al lado de la "
                 "planilla y otra en datos\\backups.").classes("text-sm text-center")
        provs = ui.checkbox("También vaciar la hoja Proveedores")
        conf = ui.input("Escribí BORRAR para confirmar").classes("w-64")
        with ui.row().classes("justify-center gap-3"):
            ok_btn = ui.button("Vaciar la planilla", icon="delete_forever",
                               on_click=lambda: dlg.submit(True)).props("color=negative unelevated")
            ui.button("Cancelar", on_click=lambda: dlg.submit(False)).props("flat")
        ok_btn.disable()
        conf.on_value_change(lambda e: ok_btn.enable() if (e.value or "").strip().upper() == "BORRAR" else ok_btn.disable())
    if not await dlg:
        return
    n = ui.notification("Vaciando la planilla…", spinner=True, timeout=None)
    try:
        res = await run.io_bound(preparar.reset_planilla, cfg, datos() / "backups", bool(provs.value), lambda s: None)
    except Exception as ex:
        ui.notify(f"No pude vaciar la planilla (no se cambió nada): {ex}", type="negative", multi_line=True)
        return
    finally:
        n.dismiss()
    total = sum(res["filas"].values())
    ui.notify(f"Planilla vacía: {total} filas borradas" + (f" y {res['proveedores']} proveedores" if provs.value else "")
              + f". Respaldo: {res['respaldos'][0]}", type="positive", multi_line=True, timeout=0, close_button=True)
    S.lote = None
    panel_revision.refresh()


# ------------------------------------------------------------------ start

def _puerto_ocupado(p: int) -> bool:
    with socket.socket() as s:
        s.settimeout(0.3)
        return s.connect_ex(("127.0.0.1", p)) == 0


def _salir() -> None:
    """Stop the server; if something keeps it alive, force the exit a moment later."""
    import os
    import threading
    app.shutdown()
    threading.Timer(1.5, lambda: os._exit(0)).start()


@app.get("/api/version")
def _api_version():
    return {"version": __version__}


def _version_corriendo(url: str) -> str:
    import json
    import urllib.request
    try:
        with urllib.request.urlopen(url + "/api/version", timeout=2) as r:
            return json.loads(r.read().decode()).get("version", "")
    except Exception:
        return ""                     # versions before 0.5.2 have no /api/version


def _cerrar_viejo(puerto: int) -> bool:
    """Kill the old app still listening on the port (only if it is a python process)."""
    import subprocess
    import time
    if sys.platform != "win32":
        return False
    sin_ventana = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    try:
        net = subprocess.run(["netstat", "-ano", "-p", "TCP"], capture_output=True, text=True,
                             creationflags=sin_ventana).stdout
    except Exception:
        return False
    pids = set()
    for linea in net.splitlines():
        c = linea.split()
        if len(c) >= 5 and c[1].endswith(f":{puerto}") and c[3].upper() in ("LISTENING", "ESCUCHANDO"):
            pids.add(c[4])
    for pid in pids:
        t = subprocess.run(["tasklist", "/FI", f"PID eq {pid}", "/FO", "CSV", "/NH"], capture_output=True,
                           text=True, creationflags=sin_ventana).stdout.lower()
        if "python" in t:
            subprocess.run(["taskkill", "/PID", pid, "/F"], capture_output=True, creationflags=sin_ventana)
    for _ in range(25):
        if not _puerto_ocupado(puerto):
            return True
        time.sleep(0.2)
    return False


def _navegador_app() -> str | None:
    """Edge (comes with Windows) or Chrome, to open the app in its own window without tabs or address bar."""
    if sys.platform != "win32":
        return None
    import os
    cands = []
    try:
        import winreg
        for exe in ("msedge.exe", "chrome.exe"):
            for raiz in (winreg.HKEY_CURRENT_USER, winreg.HKEY_LOCAL_MACHINE):
                try:
                    with winreg.OpenKey(raiz, rf"Software\Microsoft\Windows\CurrentVersion\App Paths\{exe}") as k:
                        cands.append(winreg.QueryValue(k, None))
                except OSError:
                    pass
    except ImportError:
        pass
    for var, sub in (("ProgramFiles(x86)", r"Microsoft\Edge\Application\msedge.exe"),
                     ("ProgramFiles", r"Microsoft\Edge\Application\msedge.exe"),
                     ("ProgramFiles", r"Google\Chrome\Application\chrome.exe"),
                     ("LOCALAPPDATA", r"Google\Chrome\Application\chrome.exe")):
        if os.environ.get(var):
            cands.append(str(Path(os.environ[var]) / sub))
    return next((c for c in cands if c and Path(c).is_file()), None)


def abrir_ventana(url: str) -> None:
    import subprocess
    exe = _navegador_app()
    if exe:
        try:
            subprocess.Popen([exe, f"--app={url}", "--window-size=1500,950"])
            return
        except OSError:
            pass
    webbrowser.open(url)


def _abrir_cuando_este_lista(url: str, puerto: int) -> None:
    import threading
    import time

    def esperar():
        for _ in range(150):           # up to 30 s
            if _puerto_ocupado(puerto):
                abrir_ventana(url)
                return
            time.sleep(0.2)
    threading.Thread(target=esperar, daemon=True).start()


def main(abrir_navegador: bool = True) -> None:
    cfg = config.cargar()
    url = f"http://127.0.0.1:{cfg.puerto}"
    if _puerto_ocupado(cfg.puerto):
        if _version_corriendo(url) == __version__ or not _cerrar_viejo(cfg.puerto):
            if abrir_navegador:
                abrir_ventana(url)    # same version already running (or could not close the old one): just show it
            return
    if sys.stdout is None:            # pythonw: no console, log to a file
        log = datos() / "app.log"
        log.parent.mkdir(parents=True, exist_ok=True)
        sys.stdout = sys.stderr = open(log, "a", encoding="utf-8", buffering=1)
    if abrir_navegador:
        _abrir_cuando_este_lista(url, cfg.puerto)
    ui.run(host="127.0.0.1", port=cfg.puerto, title=NOMBRE, reload=False, show=False,
           language="es", show_welcome_message=False, favicon=STATIC / "ezfac.ico")
