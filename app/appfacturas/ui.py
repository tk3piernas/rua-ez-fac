"""Interfaz (NiceGUI) en una pestaña local del navegador: http://127.0.0.1:<puerto>.

La lógica vive en el núcleo (lote, ejecutor, excel); esto solo muestra, edita y confirma."""
from __future__ import annotations

import socket
import sys
import time
import webbrowser
from datetime import date
from pathlib import Path

from nicegui import app, run, ui

from . import NOMBRE, __version__, config, ejecutor, excel, reporte
from .lote import (Armador, Entrada, aplicar_exclusiones, entradas_de_carpeta, iva_al_editar_monto,
                   renombrar_proveedor)
from .model import Estado, Nivel
from .util import ahora_id, fmt_fecha, fmt_monto, parse_fecha, parse_monto, relativa, sha256

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
.alfa { background: #ff542c; color: #fff; border-radius: 999px; padding: 2px 10px 2px 3px; margin-left: 6px; cursor: default; }
.q-dialog .seccion { overflow-y: auto; max-height: 88vh; }
.dialogo-botones { position: sticky; bottom: -24px; background: #fff; padding: 10px 0 4px; z-index: 1;
                   border-top: 1px solid var(--linea); }
/* paths: wrap at spaces (or anywhere only if one piece is too long), never 'DEU / DA' */
.ruta { overflow-wrap: anywhere; word-break: normal; }
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
    ocupado: bool = False                 # a confirm, undo or retry is running: those buttons wait
    leer_ahora: dict = {}                 # client id -> read now (set by its Cargar panel; used by the header badge)


S = Estado_()


def de_a_uno(fn):
    """Confirm / undo / retry never run twice at the same time (a double click duplicated rows)."""
    import functools

    @functools.wraps(fn)
    async def envuelta(*a, **kw):
        if S.ocupado:
            ui.notify("Esperá: todavía estoy terminando lo anterior.", type="info")
            return
        S.ocupado = True
        try:
            return await fn(*a, **kw)
        finally:
            S.ocupado = False
    return envuelta


async def con_espera(mensaje: str, fn, *args):
    """A long task: spinner while it runs, and the action buttons disabled (S.ocupado) until it ends."""
    n = ui.notification(mensaje, spinner=True, timeout=None)
    previo, S.ocupado = S.ocupado, True
    try:
        return await run.io_bound(fn, *args)
    finally:
        S.ocupado = previo
        n.dismiss()


def mientras_libre(boton, permitido: bool = True, motivo: str = ""):
    """Button disabled while something is running, and always disabled when not `permitido`
    (a plain disable() would be undone by the binding)."""
    if not permitido:
        boton.disable()
        if motivo:
            boton.tooltip(motivo)
        return boton
    return boton.bind_enabled_from(S, "ocupado", backward=lambda o: not o)


SOLO_SIMULAR = "Estás en 'Solo simular': para eso cambiá el modo en Configuración (Practicar o Cargar en la planilla real)."


def puede_tocar(trabajo: Path) -> bool:
    """A batch's files and rows can be changed (undo, continue): test batches always; real ones only in
    'Cargar en la planilla real'."""
    return ejecutor.donde(trabajo)[0] == "prueba" or config.cargar().modo_real


# ------------------------------------------------------------------ paths

def datos() -> Path:
    return config.carpeta_app() / "datos"


def carpeta_subidos() -> Path:
    return datos() / "entrada_ui"


def base_activa(cfg) -> Path | None:
    if S.usar_prueba:
        from .carpeta_prueba import carpeta_prueba
        return carpeta_prueba() / "Facturas pendientes"
    return config.cargar().base


def cfg_activa() -> config.Config:
    base = base_activa(None)
    return config.cargar(base=base)


def puede_escribir(cfg) -> bool:
    return S.usar_prueba or cfg.modo_real


INICIO = time.time()                         # files uploaded before this moment come from another session


def lista_subidos() -> list[Path]:
    sub = carpeta_subidos()
    return sorted((f for f in sub.rglob("*") if f.is_file()), key=lambda f: (f.parent.name, f.name.lower())) if sub.exists() else []


def quitar_subidos() -> tuple[int, int]:
    """Deletes what was uploaded through the app (never the originals). Returns (deleted, left)."""
    sub = carpeta_subidos()
    archivos = [f for f in sub.rglob("*") if f.is_file()] if sub.exists() else []
    quedan = 0
    for f in archivos:
        try:
            f.unlink()
        except OSError:
            quedan += 1
    for d in sorted((d for d in sub.rglob("*") if d.is_dir()), reverse=True) if sub.exists() else []:
        try:
            d.rmdir()
        except OSError:
            pass
    return len(archivos) - quedan, quedan


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


def archivos_para_leer(cfg) -> list[Path]:
    """What 'Leer facturas' would read now: _Entrada\\<area> and what was uploaded through the app."""
    out = []
    if cfg and cfg.base:
        for a in cfg.areas:
            d = cfg.entrada / a
            if d.is_dir():
                out += [f for f in d.iterdir() if f.is_file()]
    return out + lista_subidos()


def _marcas(fs: list[Path]) -> list[str]:
    out = []
    for f in fs:
        try:
            out.append(f"{str(f).lower()}|{int(f.stat().st_mtime)}")
        except OSError:
            pass
    return out


def _archivo_lectura() -> Path:
    return datos() / "ultima_lectura.json"


def recordar_lectura(cfg, fs: list[Path]) -> None:
    """Which files the last reading saw (per 'Facturas pendientes'), to tell new ones apart later."""
    import json
    try:
        d = json.loads(_archivo_lectura().read_text(encoding="utf-8"))
    except (OSError, ValueError):
        d = {}
    d[str(cfg.base).lower()] = _marcas(fs)
    _archivo_lectura().parent.mkdir(parents=True, exist_ok=True)
    _archivo_lectura().write_text(json.dumps(d, ensure_ascii=False), encoding="utf-8")


def para_leer(cfg) -> tuple[int, int]:
    """(files waiting to be read, how many of them are new since the last reading; 0 if never read)."""
    import json
    fs = archivos_para_leer(cfg)
    try:
        vistos = json.loads(_archivo_lectura().read_text(encoding="utf-8")).get(str(cfg.base).lower())
    except (OSError, ValueError):
        vistos = None
    if vistos is None:
        return len(fs), 0
    return len(fs), len(set(_marcas(fs)) - set(vistos))


def armar(cfg):
    lid = ahora_id()
    trabajo = datos() / "lotes" / lid
    leidos = archivos_para_leer(cfg)
    ents, txts = entradas(cfg)
    lote = Armador(cfg, simular=False, trabajo=trabajo).procesar(ents, txts)
    reporte.guardar(lote, trabajo.parent, reporte.texto(lote, cfg))
    recordar_lectura(cfg, leidos)
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


def _columnas(cols: list[dict]) -> list[dict]:
    """Every fixed width is also its minimum; every header shows its full name on hover."""
    return [{**c, **({"minWidth": c["width"]} if "width" in c else {}), "headerTooltip": c["headerName"]} for c in cols]


COLUMNAS = _columnas([
    {"field": "incluir", "headerName": "Cargar", "width": 85, ":editable": "params => !params.data.bloquea",
     "cellDataType": "boolean"},
    {"field": "area", "headerName": "Área", "width": 105},
    {"field": "proveedor", "headerName": "Proveedor", "editable": True, "minWidth": 190, "flex": 1, "tooltipField": "proveedor"},
    {"field": "factura", "headerName": "Factura", "width": 190, "tooltipField": "factura"},
    {"field": "fecha", "headerName": "Fecha servicio", "editable": True, "width": 115},
    {"field": "vto", "headerName": "Vencimiento", "editable": True, "width": 115,
     "cellClassRules": {"text-red-700 font-bold": "data.vencida"}},
    {"field": "tarea", "headerName": "N° de tarea", "editable": True, "width": 110},
    {"field": "servicio", "headerName": "Servicio", "editable": True, "minWidth": 260, "flex": 2, "tooltipField": "servicio"},
    {"field": "monto", "headerName": "Monto", "editable": True, "width": 130, "type": "rightAligned",
     ":valueFormatter": "p => p.value == null ? '' : p.value.toLocaleString('es-AR', {minimumFractionDigits: 2, maximumFractionDigits: 2})"},
    {"field": "iva", "headerName": "IVA (H)", "width": 100, "type": "rightAligned"},
    {"field": "fuente", "headerName": "Leída con", "width": 120},
    {"field": "nuevo", "headerName": "Prov. nuevo", "width": 105},
])


def aplicar_edicion(lote, args: dict) -> str | None:
    """Applies one grid edit to the batch. Returns an error message, or None."""
    data = args.get("data") or {}
    col = args.get("colId")
    i = data.get("idx")
    if i is None or i >= len(lote.filas):
        return None
    r = lote.filas[i]
    v = data.get(col)
    if r.bloqueada:
        r.excluida = True
        return "Esta factura tiene un aviso que la frena (en rojo): no se carga"
    if col == "incluir":
        r.excluida = not bool(v)
        return None
    r.editada = True
    if col == "proveedor":
        return renombrar_proveedor(lote, S.cfg, r.factura_id, str(v or ""))
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
        antes = r.iva
        r.monto = m
        r.iva = iva_al_editar_monto(lote, r)
        if antes not in (None, 0) and r.iva is None:
            return "Cambió el monto: H vuelve a la fórmula de la planilla (el IVA leído era del monto anterior)"
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
            aviso_alfa()
            with ui.row().classes("items-center gap-1 ml-4"):
                ui.button("Abrir planilla", icon="table_view", on_click=abrir_planilla).props("flat color=white").tooltip(
                    "Abre en Excel la planilla en la que está escribiendo la app (la real o la copia de prueba)")
                ui.button("Exportar planilla", icon="file_download", on_click=dialogo_exportar).props("flat color=white")
                aviso_para_leer()
        # mode on the right and, last of all, Salir at the far right (also when the header wraps)
        with ui.row().classes("items-center gap-3 ml-auto"):
            lbl_modo = ui.label().classes("pill")
            ui.button("Salir", icon="logout", on_click=_salir).props("flat color=white")

    def refrescar_modo():
        cfg = S.cfg
        if S.usar_prueba:
            lbl_modo.text, tipo = "Trabajando sobre la CARPETA DE PRUEBA", "pill-prueba"
        elif puede_escribir(cfg):
            lbl_modo.text, tipo = "Trabajando sobre carpetas y planilla REALES", "pill-real"
        else:
            lbl_modo.text, tipo = "Solo simulando: no se mueve ni se escribe nada", "pill-sim"
        lbl_modo.classes(replace=f"pill {tipo}")

    with ui.column().classes("contenido items-center gap-4 px-4"):
        with ui.tabs().props("align=center").classes("w-full") as tabs:
            t_cargar = ui.tab("Cargar", icon="upload_file")
            t_rev = ui.tab("Revisión", icon="fact_check")
            t_res = ui.tab("Resultado", icon="task_alt")
            t_hist = ui.tab("Historial", icon="history")
            t_revisar = ui.tab("Revisar", icon="rule")
            t_prov = ui.tab("Proveedores", icon="contacts")
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
            with ui.tab_panel(t_revisar).classes("items-center"):
                panel_revisar()
            with ui.tab_panel(t_prov).classes("items-center"):
                panel_proveedores()
            with ui.tab_panel(t_conf).classes("items-center"):
                panel_config(refrescar_modo)
    refrescar_modo()


@ui.refreshable
def aviso_para_leer() -> None:
    if not S.cfg:
        return
    n, nuevos = para_leer(S.cfg)
    if not n:
        return
    texto = f"{n} para leer" + (f" · {nuevos} nuevos" if nuevos else "")
    async def clic():
        f = S.leer_ahora.get(ui.context.client.id)
        if f:
            await f()
    ui.button(texto, icon="mark_email_unread", on_click=clic).props(
        "unelevated rounded dense color=white text-color=primary").classes("px-3").tooltip(
        "Hay archivos esperando en _Entrada o subidos acá" + (" (y algunos llegaron después de la última lectura)"
                                                                if nuevos else "") + ". Clic: ir a Cargar y leerlos.")


def aviso_alfa() -> None:
    with ui.row().classes("alfa items-center gap-2 no-wrap"):
        ui.image("/ezfac/simba.png").classes("w-7 h-7 rounded-full").props("fit=cover")
        ui.label("ALPHA").classes("font-bold text-xs")
        with ui.tooltip().classes("text-sm"):
            ui.label("¡Esto es una alpha! Mandá bugs y sugerencias a Matheo.")


# ------------------------------------------------------------------ Cargar

def panel_cargar(tabs, t_rev, refrescar_modo):
    with ui.row().classes("w-full justify-center items-stretch gap-6"):
        with seccion("1. Archivos", "upload_file", "w-[420px]"):
            area = ui.select({"": "Detectar por proveedor", **{a: a for a in config.AREAS}},
                             value="", label="Área de lo que subís").classes("w-72")

            async def subido(e):
                if Path(e.file.name).suffix.lower() == ".txt" and not area.value:
                    ui.notify(f"{e.file.name}: para los .txt elegí el área arriba (no se puede detectar por "
                              "proveedor). No lo subí.", type="warning", multi_line=True)
                    return
                dest = carpeta_subidos() / (area.value or "_Detectar")
                dest.mkdir(parents=True, exist_ok=True)
                await e.file.save(dest / Path(e.file.name).name)
                contar.refresh()

            subidor = ui.upload(multiple=True, auto_upload=True, on_upload=subido,
                                label="Arrastrá PDF, zip o .txt (o hacé clic)").props('accept=".pdf,.zip,.txt"').classes("w-full")
            ui.label("Se copian a la app; los originales quedan donde estaban.").classes("text-xs text-gray-500 text-center")

        with seccion("2. Qué se va a leer", "folder_open", "w-[340px]"):
            @ui.refreshable
            def contar():
                c = contar_entrada(S.cfg) if S.cfg.base else {}
                with ui.column().classes("items-center gap-1"):
                    for etiqueta, n in ([(f"_Entrada\\{a}", c.get(a, 0)) for a in config.AREAS]
                                        + [("Subidos acá", c.get("subidos", 0))]):
                        with ui.row().classes("gap-2 items-baseline"):
                            ui.label(etiqueta).classes("text-sm text-gray-600")
                            ui.label(str(n)).classes("text-xl font-bold").style("color: var(--navy)")
                    subidos = lista_subidos()
                    viejos = [f for f in subidos if f.stat().st_mtime < INICIO]
                    if viejos:
                        ui.label(f"Quedaron {len(viejos)} archivos subidos en otra sesión: se suman al próximo "
                                 "lote. Quitalos si no van.").classes("text-xs text-amber-700 text-center")
                    with ui.column().classes("gap-0 w-full"):
                        for f in subidos:
                            with ui.row().classes("items-center gap-1 no-wrap w-full"):
                                ui.label(f"{f.parent.name.replace('_Detectar', 'Detectar')} · {f.name}").classes(
                                    "text-xs text-gray-700 truncate grow" + (" text-amber-700" if f in viejos else ""))
                                ui.button(icon="close", on_click=lambda f=f: quitar_uno(f)).props(
                                    "flat round dense size=xs").tooltip("Quitar este archivo")
                    b = ui.button("Quitar lo subido", icon="delete_sweep", on_click=vaciar_subidos).props("flat dense")
                    if not c.get("subidos"):
                        b.disable()
                        b.tooltip("No hay nada subido acá")

            def quitar_uno(f: Path):
                try:
                    f.unlink()
                    ui.notify(f"Quité {f.name}", type="positive")
                except OSError as ex:
                    ui.notify(f"No pude quitar {f.name}: {ex}", type="warning")
                contar.refresh()

            def vaciar_subidos():
                n, quedan = quitar_subidos()
                subidor.reset()                      # also the list of the upload box
                if quedan:
                    ui.notify(f"Quité {n} archivos; {quedan} no se pudieron borrar (¿están abiertos?)", type="warning")
                else:
                    ui.notify(f"Quité {n} archivos subidos" if n else "No había nada subido", type="positive" if n else "info")
                contar.refresh()
            contar()

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

            async def ir_y_leer():
                """Header badge: read from any tab (leer() then opens Revisión)."""
                n = ui.notification("Leyendo las facturas…", spinner=True, timeout=None)
                try:
                    await leer()
                finally:
                    n.dismiss()
            S.leer_ahora[ui.context.client.id] = ir_y_leer
            ui.label("Leer no mueve ni escribe nada.").classes("text-xs text-gray-500 text-center")
    ui.timer(5.0, lambda: (contar.refresh(), aviso_para_leer.refresh()))


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
                # fixed widths that never shrink (fitting everything to the width made the headers unreadable);
                # Proveedor and Servicio take the free space, and a narrow window scrolls sideways
                "defaultColDef": {"resizable": True, "sortable": True, "wrapHeaderText": True, "autoHeaderHeight": True},
                "tooltipShowDelay": 300,
            }, theme="balham").classes("w-full").style(f"height: {min(620, 64 + 28 * len(filas))}px")

            def editado(e):
                err = aplicar_edicion(lote, e.args)
                if err:
                    ui.notify(err, type="warning", multi_line=True)
                grid.options["rowData"] = filas_grid(lote)     # other rows of the provider, H, undone edits
                grid.update()
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
        seccion_claude(lote, cfg)
        if lote.proveedores:
            with seccion(f"Proveedores nuevos para la hoja Proveedores ({len(lote.proveedores)})", "person_add"):
                ui.aggrid({"columnDefs": _columnas([{"field": c, "headerName": h, "width": w, "tooltipField": c}
                                                    for c, h, w in
                                                    [("area", "Área", 90), ("nombre", "Proveedor", 220), ("cuit", "CUIT", 120),
                                                     ("banco", "Banco", 150), ("cbu", "CBU / CVU", 200), ("alias", "Alias", 150),
                                                     ("obs", "Observaciones", 400)]]),
                           "defaultColDef": {"resizable": True, "wrapHeaderText": True, "autoHeaderHeight": True},
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
                if puede_escribir(cfg):
                    mientras_libre(btn)
                else:
                    btn.disable()
                    btn.tooltip("Estás en 'Solo simular': para cargar, cambiá el modo en Configuración")
                ui.button("Copiar filas al portapapeles", icon="content_copy", on_click=copiar).props("outline")
            if not puede_escribir(cfg):
                ui.label("Estás en 'Solo simular': para cargar, cambiá el modo en Configuración.").classes("text-xs text-amber-700")
            ui.label(f"Informe guardado en datos\\lotes\\{lote.id}").classes("text-xs text-gray-500")


def _rel(p, cfg) -> str:
    return relativa(p, cfg.base)


async def copiar():
    try:
        await run.io_bound(excel.copiar_portapapeles, S.lote)
        ui.notify("Filas copiadas. En Excel: clic en la columna A de la primera fila vacía y Ctrl+V.", type="positive")
    except Exception as ex:
        ui.notify(f"No pude copiar: {ex}", type="negative")


@de_a_uno
async def confirmar(tabs, t_res):
    lote, cfg = S.lote, S.cfg
    excluidas = aplicar_exclusiones(lote)
    act = [r for r in lote.filas if not r.excluida]
    movs = len([p for p in lote.plan if p["op"] in ("mover", "copiar", "mover_zip", "a_duplicados", "txt_procesado")])
    hay_avisos = any(a.nivel in (Nivel.REVISAR, Nivel.BLOQUEA) for a in lote.avisos)
    if not hay_avisos and not excluidas:
        await _confirmado(tabs, t_res, lote, cfg)
        return
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
    await _confirmado(tabs, t_res, lote, cfg)


async def _confirmado(tabs, t_res, lote, cfg):
    n = ui.notification("Cargando en " + ("la COPIA de prueba…" if S.usar_prueba else "la planilla REAL…"),
                        spinner=True, timeout=None)
    try:
        S.resultado = await run.io_bound(lambda: ejecutor.Ejecutor(cfg, lote, S.trabajo, log=lambda s: None,
                                                                    modo="prueba" if S.usar_prueba else "real").ejecutar())
    except ejecutor.YaConfirmado as ex:
        ui.notify(f"No hice nada: {ex}.", type="info")
        return
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
            est, msg = await con_espera("Guardando tus cambios y escribiendo las filas…", ejecutor.reintentar_excel, cfg, S.trabajo, True, lambda s: None)
            S.resultado["excel"], S.resultado["excel_mensaje"] = est, msg
    panel_resultado.refresh()
    panel_historial.refresh()
    tabs.set_value(t_res)


@de_a_uno
async def continuar(trabajo: Path, con_archivos: bool = True):
    if not puede_tocar(trabajo):
        ui.notify(SOLO_SIMULAR, type="warning", multi_line=True)
        return
    pagina_ = ui.context.client.layout
    n = ui.notification("Terminando el lote…", spinner=True, timeout=None)
    try:
        r = await run.io_bound(lambda: ejecutor.continuar(S.cfg, trabajo, log=lambda s: None, con_archivos=con_archivos))
    except Exception as ex:
        with pagina_:
            ui.notify(f"No pude seguir: {ex}", type="negative", multi_line=True)
        return
    finally:
        n.dismiss()
    est = ejecutor.estado(trabajo)
    with pagina_:
        if not est.get("archivos_pendientes") and r["excel"] in ("escrito", "no corresponde"):
            ui.notify("Listo: el lote quedó completo.", type="positive")
        else:
            ui.notify("Todavía quedan cosas pendientes: " + "; ".join(r["errores"] + ([r["excel_mensaje"]] if r["excel_mensaje"] else [])),
                      type="warning", multi_line=True, timeout=0, close_button=True)
    if S.trabajo == trabajo and S.resultado:
        S.resultado.update({k: r[k] for k in ("excel", "excel_mensaje", "errores", "sin_archivo")})
        S.resultado["hechos"] = S.resultado.get("hechos", []) + r["hechos"]
    panel_resultado.refresh()
    panel_historial.refresh()


def fecha_lote(lid: str) -> str:
    """'2026-10-28_104143' -> '28/10/2026 10:41' (the id itself if it has another shape)."""
    import re
    m = re.match(r"(\d{4})-(\d{2})-(\d{2})_(\d{2})(\d{2})", lid)
    return f"{m[3]}/{m[2]}/{m[1]} {m[4]}:{m[5]}" if m else lid


def texto_donde(trabajo: Path) -> str:
    modo, planilla = ejecutor.donde(trabajo)
    lugar = {"prueba": "la COPIA de prueba", "real": "la planilla REAL"}.get(modo, "la planilla")
    return f"{lugar}{f' ({planilla})' if planilla else ''}"


# ------------------------------------------------------------------ modo B: hand the unread ones to Claude (Cowork)

def seccion_claude(lote, cfg):
    from . import claude_cowork
    cands = claude_cowork.candidatas(lote)
    sin_cargar = claude_cowork.respondidas_sin_cargar(lote)
    peds = claude_cowork.pedidos(cfg) if cfg.base else []
    if not cands and not peds and not sin_cargar:
        return
    with seccion("Pasarle a Claude lo que no pude leer", "smart_toy"):
        for f in sin_cargar:
            ui.label(f"• {f.archivo.nombre}: Claude la leyó, pero {claude_cowork.por_que_no(f.ext)}. "
                     "Queda en _Entrada para cargarla a mano.").classes("text-sm text-amber-700")
        if cands:
            ui.label(("1 PDF no se pudo leer" if len(cands) == 1 else f"{len(cands)} PDF no se pudieron leer")
                     + ". Puedo copiarlos a una carpeta con un pedido para Claude "
                     "(Cowork): los originales quedan en _Entrada. Nunca van los .txt, y los PDF con un CBU impreso "
                     "quedan destildados.").classes("text-sm text-gray-600 text-center")
            marcados = {}
            with ui.column().classes("gap-0 w-full"):
                for f in cands:
                    con_cbu = claude_cowork.tiene_cbu(f)
                    marcados[f.id] = ui.checkbox(f"{f.archivo.nombre}" + ("  (tiene un CBU impreso)" if con_cbu else ""),
                                                 value=not con_cbu).classes("text-sm")

            @de_a_uno
            async def preparar_claude():
                if not puede_escribir(cfg):
                    ui.notify(SOLO_SIMULAR, type="warning", multi_line=True)
                    return
                elegidas = [f for f in cands if marcados[f.id].value]
                if not elegidas:
                    ui.notify("No elegiste ningún PDF", type="info")
                    return
                try:
                    d = await con_espera("Preparando el pedido para Claude…", claude_cowork.preparar, cfg, lote, elegidas)
                except Exception as ex:
                    ui.notify(f"No pude preparar el pedido: {ex}", type="negative", multi_line=True)
                    return
                ui.notify(f"Listo: {len(elegidas)} PDF en {d}. Abrí Claude (Cowork) en esa carpeta y pedile: "
                          "'Leé pedido.md y hacé lo que dice'. Cuando termine, tocá Leer facturas de nuevo.",
                          type="positive", multi_line=True, timeout=0, close_button=True)
                panel_revision.refresh()
            @de_a_uno
            async def leer_con_api():
                from . import claude_api
                if not puede_escribir(cfg):
                    ui.notify(SOLO_SIMULAR, type="warning", multi_line=True)
                    return
                elegidas = [f for f in cands if marcados[f.id].value]
                if not elegidas:
                    ui.notify("No elegiste ningún PDF", type="info")
                    return
                with ui.dialog() as dlg, ui.card().classes("items-center"):
                    ui.label(f"¿Mando {len(elegidas)} PDF a la API de Claude?").classes("font-bold")
                    ui.label("Se envían a Anthropic para leerlos (solo los PDF elegidos, nunca los .txt). Cuesta unos "
                             "centavos de dólar por factura, que se cobran en la cuenta de la clave.").classes(
                        "text-sm text-center")
                    with ui.row().classes("justify-center"):
                        ui.button("Sí, leer con Claude", on_click=lambda: dlg.submit(True)).props("color=secondary")
                        ui.button("Cancelar", on_click=lambda: dlg.submit(False)).props("flat")
                if not await dlg:
                    return
                try:
                    r = await con_espera(f"Claude está leyendo {len(elegidas)} PDF…", claude_api.leer, cfg, lote,
                                         elegidas, lambda s: None)
                except claude_api.SinClave as ex:
                    ui.notify(f"No pude: {ex}.", type="warning", multi_line=True)
                    return
                except Exception as ex:
                    ui.notify(f"No pude leer con Claude: {ex}", type="negative", multi_line=True)
                    return
                if r["errores"]:
                    ui.notify("Claude no pudo con: " + "; ".join(f"{n} ({m})" for n, m in r["errores"]),
                              type="warning", multi_line=True, timeout=0, close_button=True)
                if r["leidas"]:
                    ui.notify(f"Claude leyó {r['leidas']}. Vuelvo a leer el lote con eso.", type="positive")
                    f_leer = S.leer_ahora.get(ui.context.client.id)
                    if f_leer:
                        await f_leer()
                        return
                panel_revision.refresh()

            from . import claude_api
            hay_clave = bool(claude_api.origen_clave())
            with ui.row().classes("justify-center gap-3"):
                mientras_libre(ui.button("Leer ahora con la API de Claude", icon="bolt", on_click=leer_con_api).props(
                    "color=secondary unelevated"), puede_escribir(cfg) and hay_clave,
                    SOLO_SIMULAR if not puede_escribir(cfg) else "Falta la clave de la API (Configuración → Claude)")
                mientras_libre(ui.button("Preparar el pedido para Claude (Cowork)", icon="smart_toy",
                                         on_click=preparar_claude).props("outline"), puede_escribir(cfg), SOLO_SIMULAR)
        if peds:
            ui.label("Pedidos a Claude").classes("font-bold text-sm mt-2")
            for pd in peds[:10]:
                with ui.row().classes("items-center gap-2 w-full no-wrap"):
                    ui.label(f"{fecha_lote(pd['carpeta'].name)}: {pd['respondidos']} de {pd['archivos']} respondidos"
                             ).classes("text-sm grow")
                    if sys.platform == "win32":
                        import os
                        ui.button("Abrir la carpeta", icon="folder_open",
                                  on_click=lambda d=pd["carpeta"]: os.startfile(d)).props("flat dense")

                    def borrar(d=pd["carpeta"]):
                        try:
                            claude_cowork.borrar(d)
                            ui.notify("Pedido borrado (los originales siguen en _Entrada)", type="positive")
                        except Exception as ex:
                            ui.notify(f"No pude borrarlo: {ex}", type="warning")
                        panel_revision.refresh()
                    ui.button("Borrar", icon="delete", on_click=borrar).props("flat dense color=negative")
            if any(pd["respondidos"] for pd in peds):
                ui.button("Leer facturas de nuevo (con lo de Claude)", icon="refresh",
                          on_click=lambda: S.leer_ahora.get(ui.context.client.id) and S.leer_ahora[ui.context.client.id]()
                          ).props("outline")


@ui.refreshable
def panel_clave_claude():
    """Modo A: the API key, kept in the Windows Credential Manager (never in a file)."""
    from . import claude_api
    with seccion("Claude para lo que no se puede leer", "bolt", "w-[760px]"):
        ui.label("Con una clave de la API de Claude, en Revisión aparece 'Leer ahora con la API de Claude' para los "
                 "PDF que la app no pudo leer. La clave se guarda en el Administrador de credenciales de Windows, "
                 "nunca en un archivo. Cada factura cuesta unos centavos de dólar en la cuenta de la clave.").classes(
            "text-sm text-gray-600 text-center")
        origen = claude_api.origen_clave()
        ui.label(f"Clave: {'guardada en ' + origen if origen else 'no hay ninguna'}").classes(
            "text-sm text-center " + ("chequeo-ok" if origen else "text-gray-500"))
        with ui.row().classes("justify-center items-center gap-3"):
            entrada = ui.input("Clave de la API (sk-ant-…)", password=True, password_toggle_button=True).classes("w-96")

            def guardar():
                try:
                    claude_api.guardar_clave(entrada.value)
                except Exception as ex:
                    ui.notify(f"No la guardé: {ex}", type="warning")
                    return
                entrada.value = ""
                ui.notify("Clave guardada en el Administrador de credenciales de Windows", type="positive")
                panel_clave_claude.refresh()
                panel_revision.refresh()

            def borrar():
                claude_api.borrar_clave()
                ui.notify("Clave borrada del Administrador de credenciales", type="positive")
                panel_clave_claude.refresh()
                panel_revision.refresh()
            ui.button("Guardar la clave", icon="key", on_click=guardar).props("color=secondary unelevated")
            if origen.startswith("Administrador"):
                ui.button("Borrar la clave", icon="delete", on_click=borrar).props("flat color=negative")


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
            if est.get("archivos_pendientes"):
                ui.label(f"{est.get('filas_sin_archivo', 0)} filas no se escribieron porque su archivo no se movió "
                         f"(sigue en _Entrada). Tocá 'Reintentar archivos' cuando el archivo esté libre.").classes(
                    "text-amber-700 text-sm text-center")
        with seccion("Qué se hizo", "list", "w-[760px]"):
            with ui.expansion("Ver el detalle", icon="list").classes("w-full"):
                for h in r["hechos"]:
                    ui.label("• " + h).classes("text-sm")
        with seccion("Acciones", "build", "w-[760px]"):
            with ui.row().classes("justify-center gap-4"):
                if est.get("archivos_pendientes"):
                    mientras_libre(ui.button("Reintentar archivos", icon="drive_file_move",
                                             on_click=lambda: continuar(S.trabajo)))
                if est.get("excel") == "pendiente":
                    @de_a_uno
                    async def reintentar():
                        if not puede_tocar(S.trabajo):
                            ui.notify(SOLO_SIMULAR, type="warning", multi_line=True)
                            return
                        e2, msg = await con_espera("Escribiendo las filas…", ejecutor.reintentar_excel, S.cfg, S.trabajo, False, lambda s: None)
                        S.resultado["excel"], S.resultado["excel_mensaje"] = e2, msg
                        ui.notify("Filas escritas" if e2 == "escrito" else f"Siguen pendientes: {msg}",
                                  type="positive" if e2 == "escrito" else "warning")
                        panel_resultado.refresh()
                    mientras_libre(ui.button("Reintentar planilla", icon="refresh", on_click=reintentar))
                    ui.button("Copiar filas al portapapeles", icon="content_copy", on_click=copiar).props("outline")
                mientras_libre(ui.button("Deshacer este lote", icon="undo", on_click=lambda: deshacer(S.trabajo))
                               .props("color=negative outline"))


@de_a_uno
async def deshacer(trabajo: Path):
    if not puede_tocar(trabajo):
        ui.notify(SOLO_SIMULAR, type="warning", multi_line=True)
        return
    with ui.dialog() as dlg, ui.card().classes("items-center"):
        ui.label(f"¿Deshago el lote {trabajo.name}?").classes("font-bold")
        escritas = ejecutor.estado(trabajo).get("excel") == "escrito"
        ui.label(("Borra las filas que agregó en " + texto_donde(trabajo) + " y devuelve" if escritas else "Devuelve")
                 + " los archivos a donde estaban (solo lo que sigue igual a como lo dejó la app).").classes("text-center")
        with ui.row().classes("justify-center"):
            ui.button("Sí, deshacer", on_click=lambda: dlg.submit(True)).props("color=negative")
            ui.button("Cancelar", on_click=lambda: dlg.submit(False)).props("flat")
    if not await dlg:
        return
    pagina_ = ui.context.client.layout          # the button may be gone when the undo ends (panels refresh)
    try:
        problemas = await con_espera("Deshaciendo el lote…", lambda: ejecutor.deshacer(trabajo, log=lambda s: None, cfg=S.cfg))
    except ejecutor.YaDeshecho as ex:
        with pagina_:
            ui.notify(f"No hice nada: {ex}.", type="info")
        return
    except Exception as ex:
        with pagina_:
            ui.notify(f"No pude deshacer (no se cambió nada): {ex}", type="negative", multi_line=True)
        return
    with pagina_:
        if problemas:
            ui.notify("Deshecho, salvo: " + "; ".join(problemas), type="warning", multi_line=True, timeout=0,
                      close_button=True)
        else:
            ui.notify("Lote deshecho: todo volvió a como estaba.", type="positive")
    S.resultado = None
    panel_resultado.refresh()
    panel_historial.refresh()


# ------------------------------------------------------------------ Historial

@ui.refreshable
def panel_historial():
    from . import tandas
    d = datos() / "lotes"
    lotes = sorted((x for x in d.iterdir() if x.is_dir()), reverse=True) if d.exists() else []
    panel_tandas()
    n_tanda = tandas.tanda_de(datos())
    with seccion("Lotes anteriores", "history", "w-[760px]"):
        if not lotes:
            ui.label("Todavía no hay lotes.").classes("text-gray-500")
            return
        ultimo = ejecutor.ultimo_confirmado(d)
        filas = []
        for x in lotes[:40]:
            e = ejecutor.estado(x)
            modo = ejecutor.donde(x)[0] if e.get("estado") else ""
            filas.append({"lote": fecha_lote(x.name), "tanda": n_tanda.get(x.name, ""),
                          "estado": ejecutor.etiqueta(x), "donde": modo.upper(),
                          "planilla": e.get("excel", ""), "errores": len(e.get("errores", []))})
        ui.aggrid({"columnDefs": [{"field": "lote", "headerName": "Fecha", "width": 160},
                                  {"field": "tanda", "headerName": "Tanda", "width": 80},
                                  {"field": "estado", "headerName": "Estado", "width": 200},
                                  {"field": "donde", "headerName": "Dónde", "width": 100},
                                  {"field": "planilla", "headerName": "Planilla", "width": 150},
                                  {"field": "errores", "headerName": "Errores", "width": 100}],
                   "autoSizeStrategy": {"type": "fitGridWidth"}, "rowData": filas}, theme="balham").classes("w-full").style(f"height: {min(600, 64 + 28 * len(filas))}px")
        if ultimo:
            e = ejecutor.estado(ultimo)
            ui.label(f"Último lote: {ultimo.name} ({ejecutor.etiqueta(ultimo)}), cargado en {texto_donde(ultimo)}").classes(
                "text-sm text-gray-600 text-center ruta")
            with ui.row().classes("justify-center gap-4"):
                if ejecutor.etiqueta(ultimo) == "quedó a medias" or e.get("archivos_pendientes") or e.get("excel") == "pendiente":
                    mientras_libre(ui.button("Terminar este lote", icon="play_arrow",
                                             on_click=lambda: continuar(ultimo)).props("outline"),
                                   puede_tocar(ultimo), SOLO_SIMULAR)
                mientras_libre(ui.button(f"Deshacer el último lote ({ultimo.name}, {ejecutor.donde(ultimo)[0].upper() or '?'})",
                                         icon="undo", on_click=lambda: deshacer(ultimo)).props("color=negative outline"),
                               puede_tocar(ultimo), SOLO_SIMULAR)


def panel_tandas():
    """Close the current round: export it as 'tanda N', note its batches and, optionally, empty the workbook."""
    from . import tandas
    cfg = S.cfg
    with seccion("Tanda de pagos", "inventory_2", "w-[760px]"):
        if not cfg or not cfg.planilla:
            ui.label("No sé cuál es la planilla.").classes("text-gray-500")
            return
        n = tandas.siguiente(datos(), cfg.planilla)
        abiertos = tandas.lotes_abiertos(datos(), cfg.planilla)
        ui.label(f"Tanda actual: {n} · {len(abiertos)} lotes cargados desde el último cierre "
                 f"({'COPIA de prueba' if S.usar_prueba else 'planilla REAL'})").classes("text-sm text-center")
        ui.label("Cerrar la tanda guarda una copia de la planilla como 'DEUDA A PAGAR ARG tanda N' en Planillas "
                 "exportadas, anota qué lotes entraron y, si querés, deja la planilla vacía para la próxima.").classes(
            "text-xs text-gray-500 text-center")
        with ui.row().classes("justify-center gap-3"):
            mientras_libre(ui.button(f"Cerrar la tanda {n}…", icon="inventory_2", on_click=cerrar_tanda).props(
                "color=secondary unelevated"), puede_escribir(cfg), SOLO_SIMULAR)
            u = tandas.ultima(datos(), cfg.planilla)
            if u:
                mientras_libre(ui.button(f"Deshacer el cierre de la tanda {u['n']}", icon="undo",
                                         on_click=deshacer_tanda).props("color=negative outline"),
                               puede_escribir(cfg), SOLO_SIMULAR)


@de_a_uno
async def cerrar_tanda():
    from . import tandas
    cfg = S.cfg
    if not puede_escribir(cfg):
        ui.notify(SOLO_SIMULAR, type="warning", multi_line=True)
        return
    n = tandas.siguiente(datos(), cfg.planilla)
    with ui.dialog() as dlg, ui.card().classes("seccion w-[620px] max-w-full items-center gap-3 p-6"):
        ui.label(f"Cerrar la tanda {n}").classes("titulo-seccion")
        ui.label(f"{'COPIA DE PRUEBA' if S.usar_prueba else 'PLANILLA REAL'}: {cfg.planilla}").classes(
            "text-xs text-center ruta " + ("chequeo-ok" if S.usar_prueba else "chequeo-falta"))
        nombre = ui.input("Nombre de la copia", value=f"DEUDA A PAGAR ARG tanda {n}").classes("w-96")
        vaciar = ui.checkbox("Además dejar la planilla vacía para la próxima tanda (reset)")
        conf = ui.input("Para vaciarla, escribí BORRAR").classes("w-64")
        conf.bind_visibility_from(vaciar, "value")
        ui.label("Se puede deshacer desde acá mientras nadie toque la planilla ni la copia.").classes(
            "text-xs text-gray-500 text-center")
        with ui.row().classes("dialogo-botones w-full justify-center gap-3"):
            ui.button("Cerrar la tanda", icon="inventory_2", on_click=lambda: dlg.submit(True)).props(
                "color=secondary unelevated")
            ui.button("Cancelar", on_click=lambda: dlg.submit(False)).props("flat")
    if not await dlg:
        return
    if vaciar.value and (conf.value or "").strip().upper() != "BORRAR":
        ui.notify("Para vaciar la planilla hay que escribir BORRAR. No hice nada.", type="warning")
        return
    try:
        t = await con_espera("Cerrando la tanda…", tandas.cerrar, cfg, datos(), nombre.value, bool(vaciar.value),
                             lambda s: None)
    except Exception as ex:
        ui.notify(f"No pude cerrar la tanda: {ex}", type="negative", multi_line=True)
        panel_historial.refresh()
        return
    ui.notify(f"Tanda {t['n']} cerrada: {Path(t['exportada']).name}, {len(t['lotes'])} lotes"
              + (", planilla vacía" if t.get("vaciada") else ""), type="positive", multi_line=True)
    S.lote = None
    panel_revision.refresh()
    panel_historial.refresh()
    volver_reset_boton.refresh()


@de_a_uno
async def deshacer_tanda():
    from . import tandas
    if not puede_escribir(S.cfg):
        ui.notify(SOLO_SIMULAR, type="warning", multi_line=True)
        return
    u = tandas.ultima(datos(), S.cfg.planilla)
    with ui.dialog() as dlg, ui.card().classes("items-center"):
        ui.label(f"¿Deshago el cierre de la tanda {u['n']}?").classes("font-bold")
        ui.label(("La planilla vuelve a como estaba antes de vaciarla (tiene que estar cerrada en Excel) y "
                  if u.get("vaciada") else "") + "se borra la copia exportada si nadie la tocó.").classes("text-sm text-center")
        with ui.row().classes("justify-center"):
            ui.button("Sí, deshacer", on_click=lambda: dlg.submit(True)).props("color=negative")
            ui.button("Cancelar", on_click=lambda: dlg.submit(False)).props("flat")
    if not await dlg:
        return
    try:
        problemas = await con_espera("Deshaciendo el cierre…", tandas.deshacer, S.cfg, datos(), lambda s: None)
    except Exception as ex:
        ui.notify(f"No pude deshacer (no se cambió nada): {ex}", type="negative", multi_line=True)
        return
    ui.notify("Cierre deshecho" + (": " + "; ".join(problemas) if problemas else ""),
              type="warning" if problemas else "positive", multi_line=True)
    panel_historial.refresh()
    volver_reset_boton.refresh()


# ------------------------------------------------------------------ Revisar (audit, read only)

R = {"hallazgos": None, "cuando": ""}


@de_a_uno
async def revisar_ahora():
    from . import auditoria
    try:
        R["hallazgos"] = await con_espera("Revisando la planilla y las carpetas…", auditoria.revisar, S.cfg)
    except Exception as ex:
        ui.notify(f"No pude revisar: {ex}", type="negative", multi_line=True)
        return
    from datetime import datetime
    R["cuando"] = datetime.now().strftime("%d/%m/%Y %H:%M")
    panel_revisar.refresh()


@ui.refreshable
def panel_revisar():
    from . import auditoria
    with seccion("Revisar la planilla", "rule", "w-[860px]"):
        ui.label("Compara la planilla con las carpetas de los proveedores y busca cosas para mirar: links rotos, filas "
                 "repetidas o sin link, PDF que no están en ninguna fila, proveedores sin datos de pago y archivos "
                 "viejos en _Entrada. Solo lee: no cambia nada.").classes("text-sm text-gray-600 text-center")
        ui.label(f"Sobre: {S.cfg.planilla}").classes("text-xs text-gray-500 text-center ruta")
        with ui.row().classes("justify-center gap-3"):
            mientras_libre(ui.button("Revisar ahora", icon="rule", on_click=revisar_ahora).props(
                "color=secondary unelevated"))
            if R["hallazgos"]:
                texto = auditoria.informe(R["hallazgos"])
                ui.button("Copiar informe", icon="content_copy",
                          on_click=lambda: (ui.clipboard.write(texto), ui.notify("Informe copiado", type="positive"))
                          ).props("outline")
    if not R["hallazgos"]:
        return
    with seccion(f"Resultado ({R['cuando']})", "fact_check", "w-[860px]"):
        for h in R["hallazgos"]:
            n = len(h.items)
            titulo = f"{h.titulo}: " + (str(n) if n else "OK")
            with ui.expansion(titulo, icon="check_circle" if not n else "report",
                              value=0 < n <= 10).classes("w-full rounded " + ("" if not n else "bg-amber-50")):
                ui.label(h.explicacion).classes("text-xs text-gray-600")
                for x in h.items[:300]:
                    ui.label("• " + x).classes("text-sm ruta")
                if n > 300:
                    ui.label(f"… y {n - 300} más (están todos en 'Copiar informe')").classes("text-xs text-gray-500")


# ------------------------------------------------------------------ open / export the workbook

def abrir_planilla() -> None:
    r = S.cfg.planilla if S.cfg else None
    if not r or not Path(r).is_file():
        ui.notify(f"No encuentro la planilla ({r or 'sin configurar'}). Mirá Configuración → Setup.", type="warning",
                  multi_line=True)
        return
    import os
    os.startfile(r)
    ui.notify(f"Abriendo {'la COPIA DE PRUEBA' if S.usar_prueba else Path(r).name}…")


async def dialogo_exportar():
    from . import preparar
    cfg = S.cfg
    if not cfg.planilla or not Path(cfg.planilla).is_file():
        ui.notify(f"No encuentro la planilla ({cfg.planilla})", type="warning")
        return
    carpeta = Path(cfg.planilla).parent / preparar.CARPETA_EXPORTADAS
    with ui.dialog() as dlg, ui.card().classes("seccion w-[620px] max-w-full items-center gap-3 p-6"):
        ui.label("Exportar planilla").classes("titulo-seccion")
        ui.label("Guarda una copia de la planilla tal como está ahora (si está abierta en Excel, con lo que tenga) "
                 "con el nombre que elijas. La planilla sigue igual; si querés empezarla de cero, después usá "
                 "Configuración → Reset.").classes("text-sm text-center text-gray-600")
        nombre = ui.input("Nombre", value=f"DEUDA A PAGAR ARG {date.today():%d-%m-%Y}").classes("w-96")
        ui.label(f"Se guarda en: {carpeta}").classes("text-xs text-center ruta " +
                                                     ("chequeo-ok" if S.usar_prueba else "text-gray-500"))

        hecho = {}

        @de_a_uno
        async def exportar():
            try:
                destino = await con_espera("Exportando la planilla…", preparar.exportar_planilla, cfg, nombre.value)
            except Exception as ex:
                ui.notify(f"No pude exportar: {ex}", type="negative", multi_line=True)
                return
            hecho.update(destino=destino, sha=sha256(destino))
            botones.refresh()

        def deshacer_export():
            problema = preparar.borrar_exportacion(hecho["destino"], hecho["sha"])
            ui.notify(problema or f"Listo: borré {hecho['destino'].name}", type="warning" if problema else "positive")
            if not problema:
                hecho.clear()
                botones.refresh()

        @ui.refreshable
        def botones():
            with ui.row().classes("dialogo-botones w-full justify-center gap-3"):
                if not hecho:
                    mientras_libre(ui.button("Exportar", icon="file_download", on_click=exportar).props("color=secondary unelevated"))
                    ui.button("Cancelar", on_click=dlg.close).props("flat")
                    return
                import os
                ui.label(f"✔ Exportada: {hecho['destino'].name}").classes("w-full text-center font-bold chequeo-ok")
                ui.button("Abrir", icon="open_in_new", on_click=lambda: os.startfile(hecho["destino"])).props("outline")
                ui.button("Deshacer (borrar la copia)", icon="undo", on_click=deshacer_export).props(
                    "color=negative outline")
                ui.button("Cerrar", on_click=dlg.close).props("flat")
        botones()
    dlg.open()


# ------------------------------------------------------------------ Proveedores

def hist_memoria() -> Path:
    return datos() / "memoria_proveedores.json"


def historiales_memoria() -> list[Path]:
    """Pasted by hand, plus one per batch (datos\\lotes\\<id>\\memoria_proveedores.json)."""
    return [hist_memoria()] + sorted((datos() / "lotes").glob(f"*/{ejecutor.HIST_MEMORIA}"))


@ui.refreshable
def panel_proveedores():
    """Paste a provider's details, check them and keep them in PROVEEDORES RUA.xlsx (the provider memory)."""
    from . import maestro
    F = {"vals": None, "avisos": []}

    def leer():
        if not (texto.value or "").strip():
            ui.notify("Pegá primero el texto con los datos", type="warning")
            return
        F["vals"], F["avisos"] = maestro.desde_texto(texto.value, area.value)
        ficha.refresh()

    @de_a_uno
    async def guardar():
        cfg = S.cfg
        if not puede_escribir(cfg):
            ui.notify("Para guardar en la memoria real elegí 'Cargar en la planilla real' en Configuración → Modo de trabajo (o practicá en la carpeta de prueba).",
                      type="warning", multi_line=True)
            return
        if not F["vals"][1] and not F["vals"][2]:
            ui.notify("Falta el nombre o el CUIT del proveedor", type="warning")
            return
        try:
            cambios = await con_espera("Guardando en la memoria…", maestro.guardar, cfg, [(F["vals"], False)], hist_memoria(), "pegado a mano")
        except Exception as ex:
            ui.notify(f"No pude guardar: {ex}", type="negative", multi_line=True)
            return
        ui.notify(f"{F['vals'][1]}: guardado en la memoria" if cambios else "No había nada nuevo para guardar",
                  type="positive" if cambios else "info")
        F["vals"] = None
        texto.value = ""
        ficha.refresh()
        memoria.refresh()

    @de_a_uno
    async def deshacer_memoria():
        u = maestro.ultimo_de(historiales_memoria())
        if not u:
            ui.notify("No hay cambios para deshacer", type="info")
            return
        try:
            problemas = await con_espera("Deshaciendo el cambio…", lambda: maestro.deshacer(S.cfg, u[0], id_=u[1]["id"]))
        except Exception as ex:
            ui.notify(f"No pude deshacer: {ex}", type="negative", multi_line=True)
            return
        ui.notify("Listo: el último cambio se deshizo" + (". Avisos: " + "; ".join(problemas) if problemas else ""),
                  type="warning" if problemas else "positive", multi_line=True)
        memoria.refresh()

    with ui.row().classes("w-full justify-center items-start gap-6"):
        with seccion("Pegar datos de un proveedor", "content_paste", "w-[480px]"):
            ui.label("Pegá el texto como venga (mail, WhatsApp, .txt). La app saca CUIT, CBU/CVU, alias, banco y "
                     "contacto, los chequea y te muestra la ficha para revisar antes de guardar.").classes(
                "text-sm text-gray-600 text-center")
            texto = ui.textarea(placeholder="Razón social: ...\nCUIT: ...\nCBU: ...\nAlias: ...").props(
                "outlined").classes("w-full").style("min-height: 220px")
            area = ui.select({"": "Sin área", **{a: a for a in config.AREAS}}, value="", label="Área").classes("w-60")
            ui.button("Leer los datos", icon="auto_fix_high", on_click=leer).props("color=secondary unelevated")

        with seccion("Ficha del proveedor", "badge", "w-[640px]"):
            @ui.refreshable
            def ficha():
                if F["vals"] is None:
                    ui.label("Pegá un texto y tocá 'Leer los datos'.").classes("text-gray-500")
                    return
                vals = F["vals"]

                def cambiar(i, v):
                    vals[i] = str(v or "").strip()
                    chequeo.refresh()
                with ui.grid(columns=2).classes("w-full gap-x-4 gap-y-1"):
                    for i, col in enumerate(maestro.COLS[1:], start=1):
                        ui.input(col, value=vals[i], on_change=lambda e, i=i: cambiar(i, e.value)).props(
                            "dense").classes("w-full")

                @ui.refreshable
                def chequeo():
                    for ok, t in maestro.chequeos(vals):
                        ui.label(("✔ " if ok else "⚠ ") + t).classes("text-sm " + ("chequeo-ok" if ok else "chequeo-aviso"))
                    for a in F["avisos"]:
                        ui.label("⚠ " + a).classes("text-sm chequeo-aviso")
                    ya = maestro.buscar(maestro.leer(S.cfg), maestro.solo_digitos(vals[2]), vals[1])
                    if ya:
                        difs = maestro.diferencias(ya["valores"], maestro.combinar(ya["valores"], vals))
                        ui.label(f"Ya está en la memoria (fila {ya['fila']})" + (": cambia" if difs else ": sin cambios")
                                 ).classes("text-sm font-bold")
                        for col, a, d in difs:
                            ui.label(f"{col}: {a or '(vacío)'} → {d}").classes("text-xs")
                    else:
                        ui.label("Proveedor nuevo en la memoria").classes("text-sm font-bold")
                with ui.column().classes("w-full gap-0"):
                    chequeo()
                mientras_libre(ui.button("Guardar en la memoria", icon="save", on_click=guardar).props("color=primary unelevated"))
            ficha()

    @ui.refreshable
    def memoria():
        cfg = S.cfg
        r = maestro.ruta(cfg)
        with seccion("Memoria de proveedores", "inventory_2", "w-[1150px] max-w-full"):
            try:
                mem = maestro.leer(cfg)
            except Exception as ex:
                ui.label(f"No pude leer {maestro.NOMBRE}: {ex}").classes("chequeo-falta")
                return
            ui.label(f"{len(mem)} proveedores en {r or 'la planilla de proveedores'}"
                     + ("" if r and r.is_file() else " (se crea con el primer guardado)")).classes("text-sm text-gray-600 ruta")
            if mem:
                filas = [{"area": m["valores"][0], "nombre": m["valores"][1], "cuit": m["valores"][2],
                          "banco": m["valores"][8], "cbu": m["valores"][11], "alias": m["valores"][12],
                          "contacto": m["valores"][5]} for m in mem.values()]
                ui.aggrid({"columnDefs": [{"field": "area", "headerName": "Área", "width": 100},
                                          {"field": "nombre", "headerName": "Proveedor", "flex": 1, "minWidth": 200},
                                          {"field": "cuit", "headerName": "CUIT", "width": 120},
                                          {"field": "banco", "headerName": "Banco", "width": 150},
                                          {"field": "cbu", "headerName": "CBU / CVU", "width": 190},
                                          {"field": "alias", "headerName": "Alias", "width": 150},
                                          {"field": "contacto", "headerName": "Contacto", "width": 150}],
                           "defaultColDef": {"filter": True, "sortable": True}, "rowData": filas},
                          theme="balham").classes("w-full").style(f"height: {min(500, 64 + 28 * len(filas))}px")
            with ui.row().classes("gap-3"):
                if r and r.is_file() and sys.platform == "win32":
                    import os
                    ui.button("Abrir la planilla de proveedores", icon="open_in_new",
                              on_click=lambda: os.startfile(r)).props("outline")
                u = maestro.ultimo_de(historiales_memoria())
                if u:
                    e = u[1]
                    mientras_libre(ui.button(f"Deshacer el último cambio ({e['motivo']}, {e['fecha'].replace('T', ' ')[:16]}, "
                                             f"en {Path(e['ruta']).parent.name})", icon="undo",
                                             on_click=deshacer_memoria).props("color=negative outline"))
    memoria()


# ------------------------------------------------------------------ Configuración

def panel_config(refrescar_modo):
    cfg_real = config.cargar()
    with ui.column().classes("w-full items-center gap-5"):
        with seccion("Setup", "construction", "w-[760px]"):
            ui.label("Crea las carpetas que falten (en la app y en tu OneDrive de RUA), una planilla vacía si no "
                     "tenés, y revisa que esté todo bien. Nunca pisa nada que ya exista.").classes("text-sm text-gray-600 text-center")
            mientras_libre(ui.button("Revisar y preparar todo", icon="construction", on_click=dialogo_setup)).props(
                "color=secondary unelevated size=md").classes("px-5")

        with seccion("Modo de trabajo", "tune", "w-[760px]"):
            from .carpeta_prueba import carpeta_prueba
            ui.label("Elegí uno. Se ve siempre arriba a la derecha.").classes("text-sm text-gray-600 text-center")

            def modo_actual() -> str:
                return "prueba" if S.usar_prueba else ("real" if config.cargar().modo_real else "simulacion")

            MODOS = {
                "simulacion": ("Solo simular",
                               "Lee tus facturas reales de _Entrada y te muestra qué haría, pero no mueve nada ni "
                               "escribe en ninguna planilla. Sirve para revisar sin riesgo."),
                "prueba": ("Practicar en la carpeta de prueba",
                           f"Trabaja de verdad, pero sobre una copia: {carpeta_prueba()}. Mueve archivos y escribe en "
                           "la COPIA de la planilla. Tus carpetas y la planilla real no se tocan, aunque la tengas abierta."),
                "real": ("Cargar en la planilla real",
                         f"Mueve las facturas en tu OneDrive y escribe en la planilla real: {cfg_real.planilla}. "
                         "Cada lote se puede deshacer."),
            }

            async def elegir(modo: str):
                if modo == modo_actual():
                    return
                if S.lote is not None:
                    with ui.dialog() as dlg0, ui.card().classes("items-center"):
                        ui.label("Hay un lote leído sin confirmar").classes("font-bold")
                        ui.label("Si cambiás de modo se descarta, con lo que hayas editado en la grilla. "
                                 "Los archivos no se tocan: lo podés volver a leer.").classes("text-sm text-center")
                        with ui.row().classes("justify-center"):
                            ui.button("Cambiar igual", on_click=lambda: dlg0.submit(True)).props("color=negative")
                            ui.button("Cancelar", on_click=lambda: dlg0.submit(False)).props("flat")
                    if not await dlg0:
                        return
                if modo == "real":
                    with ui.dialog() as dlg, ui.card().classes("items-center"):
                        ui.label("¿Paso a cargar en la planilla REAL?").classes("font-bold")
                        ui.label("Desde ahora Confirmar mueve archivos en tu OneDrive y escribe en la planilla real.").classes(
                            "text-sm text-center")
                        with ui.row().classes("justify-center"):
                            ui.button("Sí, usar lo real", on_click=lambda: dlg.submit(True)).props("color=negative")
                            ui.button("Cancelar", on_click=lambda: dlg.submit(False)).props("flat")
                    if not await dlg:
                        return
                config.guardar_opcion(cfg_real.archivo, "app", "modo_real", modo == "real")
                S.usar_prueba = modo == "prueba"
                S.cfg = cfg_activa()
                S.lote = None
                panel_revision.refresh()
                panel_proveedores.refresh()          # its memory file depends on the mode
                panel_historial.refresh()            # the round and its buttons depend on the workbook
                refrescar_modo()
                selector.refresh()
                ui.notify(f"Modo: {MODOS[modo][0]}")

            @ui.refreshable
            def selector():
                actual = modo_actual()
                with ui.column().classes("w-full gap-2"):
                    for clave_, (titulo, texto) in MODOS.items():
                        sel = clave_ == actual
                        with ui.card().classes("w-full cursor-pointer p-3 " + ("bg-indigo-50 border-2 border-indigo-800" if sel else "")
                                               ).on("click", lambda c=clave_: elegir(c)):
                            with ui.row().classes("items-center gap-3 no-wrap"):
                                ui.icon("radio_button_checked" if sel else "radio_button_unchecked").classes(
                                    "text-indigo-900 text-xl")
                                with ui.column().classes("gap-0"):
                                    ui.label(titulo).classes("font-bold")
                                    ui.label(texto).classes("text-sm text-gray-600").style("overflow-wrap: anywhere; word-break: normal")
            selector()

            @de_a_uno
            async def rearmar():
                from . import carpeta_prueba
                await con_espera("Rearmando la carpeta de prueba…", carpeta_prueba.rearmar)
                ui.notify("Carpeta de prueba rearmada con una copia de la planilla y el lote de prueba", type="positive")
            with ui.row().classes("justify-center items-center gap-3"):
                mientras_libre(ui.button("Rearmar la carpeta de prueba", icon="restart_alt", on_click=rearmar).props("outline"))
            ui.label(f"Facturas pendientes (real): {cfg_real.base} · Configuración: {cfg_real.archivo}").classes(
                "text-xs text-gray-500 text-center ruta")

        with seccion("Empezar la planilla de cero", "restart_alt", "w-[760px]"):
            ui.label("Deja vacías las filas 12 a 210 de las hojas de cada área: valores, colores, comentarios y links. "
                     "Vuelve a poner la fórmula del IVA en H, no toca la columna I ni el formato, y es el mismo "
                     "archivo (los links compartidos siguen andando). Antes guarda dos copias de respaldo.").classes(
                "text-sm text-gray-600 text-center")
            with ui.row().classes("justify-center gap-3"):
                ui.button("Reset de la planilla…", icon="delete_forever", on_click=dialogo_reset).props(
                    "color=negative outline")
                volver_reset_boton()

        with seccion("Ventana", "web_asset", "w-[760px]"):
            def cambiar_ventanas(e):
                config.guardar_opcion(cfg_real.archivo, "app", "varias_ventanas", bool(e.value))
                ui.notify("Listo: " + ("cada clic en el ícono abre otra ventana" if e.value else
                                       "un clic en el ícono trae al frente la ventana que ya está abierta"))
            ui.switch("Permitir abrir más de una ventana de la app", value=cfg_real.varias_ventanas,
                      on_change=cambiar_ventanas)
            ui.label("Apagado: si la app ya está abierta, tocar el ícono trae al frente esa ventana en vez de abrir "
                     "otra.").classes("text-xs text-gray-500 text-center")

        panel_clave_claude()

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
                        ui.label(l[len(pref):].strip()).classes("text-sm ruta")
                    break
            else:
                ui.label(l).classes("text-sm font-bold" if l.startswith(("Todo", "Hay")) else "text-xs text-gray-500")


F_setup = {"nueva": False}


@de_a_uno
async def dialogo_setup():
    from . import instalacion, preparar
    cfg_real = config.cargar()
    F_setup["nueva"] = False
    plan = await con_espera("Revisando la PC…", preparar.plan_setup, cfg_real)
    ok, lineas = await con_espera("Revisando la PC…", instalacion.chequeo)
    with ui.dialog() as dlg, ui.card().classes("seccion w-[760px] max-w-full gap-3 p-6"):
        ui.label("Setup").classes("titulo-seccion")

        @ui.refreshable
        def cuerpo():
            if plan["sin_onedrive"]:
                ui.label("No encuentro el OneDrive de RUA sincronizado en esta PC (carpeta 'OneDrive - Rua Asistencia'). "
                         "Abrí OneDrive, iniciá sesión con tu cuenta de RUA y volvé a tocar Setup.").classes("text-red-700 text-sm")
            pend = plan["dirs"] + plan["archivos"] + plan["hojas"]
            if plan.get("planilla_nueva"):
                with ui.card().classes("w-full bg-amber-50 gap-1"):
                    ui.label(f"No encuentro la planilla: {plan['planilla_nueva']}").classes("font-bold text-sm ruta")
                    if plan.get("onedrive_sin_bajar"):
                        ui.label(f"No la creo: {plan['onedrive_sin_bajar']}. Si la planilla ya existe en la nube, "
                                 "esperá a que OneDrive termine de sincronizar y volvé a tocar Setup.").classes(
                            "text-sm text-amber-800")
                    else:
                        ui.label("Si ya existe en la nube, esperá a que OneDrive termine de sincronizar: si creo una "
                                 "vacía ahora, OneDrive va a hacer una copia en conflicto.").classes("text-sm text-amber-800")
                        ui.checkbox("Sí, crear una planilla nueva vacía").bind_value(F_setup, "nueva")
            if pend:
                ui.label(f"Voy a crear {len(pend)} cosas:").classes("font-bold")
                with ui.column().classes("gap-0 w-full"):
                    for x in pend:
                        tipo = ("hoja de la planilla (copia vacía de otra área; antes guarda un respaldo)"
                                if x in plan["hojas"] else "archivo" if x in plan["archivos"] else "carpeta")
                        ui.label(f"• {tipo}: {x}").classes("text-xs ruta")
            else:
                ui.label("No falta crear nada.").classes("font-bold chequeo-ok")
            ui.separator()
            ui.label("Chequeo").classes("font-bold")
            mostrar_chequeo(lineas)
        cuerpo()

        @de_a_uno
        async def crear():
            nonlocal plan, ok, lineas
            if plan["hojas"] and not cfg_real.modo_real:
                ui.notify("Las hojas nuevas de la planilla real se agregan en el modo 'Cargar en la planilla real' "
                          "(Configuración). Ahora creo solo las carpetas y archivos.", type="warning", multi_line=True)
                plan = {**plan, "hojas": []}
            try:
                hecho = await con_espera("Creando lo que falta…", preparar.ejecutar_setup, cfg_real, plan,
                                           bool(F_setup["nueva"] and not plan.get("onedrive_sin_bajar")))
            except Exception as ex:
                ui.notify(f"No pude crear todo: {ex}", type="negative", multi_line=True)
                return
            ui.notify(f"Listo: {len(hecho)} cosas creadas", type="positive")
            plan = await con_espera("Revisando la PC…", preparar.plan_setup, config.cargar())
            ok, lineas = await con_espera("Revisando la PC…", instalacion.chequeo)
            S.cfg = cfg_activa()
            cuerpo.refresh()
            btn.set_visibility(bool(plan["dirs"] or plan["archivos"] or plan["hojas"] or plan.get("planilla_nueva")))

        with ui.row().classes("dialogo-botones w-full justify-center gap-3"):
            btn = ui.button("Crear lo que falta", icon="add_circle", on_click=crear).props("color=secondary unelevated")
            btn.set_visibility(bool(plan["dirs"] or plan["archivos"] or plan["hojas"] or plan.get("planilla_nueva")))
            ui.button("Cerrar", on_click=dlg.close).props("flat")
    dlg.open()


@ui.refreshable
def volver_reset_boton():
    from . import preparar
    r = preparar.reset_para_volver(S.cfg, datos() / "backups") if S.cfg else None
    if r:
        mientras_libre(ui.button(f"Volver al respaldo del reset ({r['fecha'].replace('T', ' ')[:16]})", icon="undo",
                                 on_click=volver_reset).props("outline"), puede_escribir(S.cfg), SOLO_SIMULAR)


@de_a_uno
async def volver_reset():
    from . import preparar
    if not puede_escribir(S.cfg):
        ui.notify(SOLO_SIMULAR, type="warning", multi_line=True)
        return
    with ui.dialog() as dlg, ui.card().classes("items-center"):
        ui.label("¿Vuelvo la planilla a como estaba antes del reset?").classes("font-bold")
        ui.label(f"{'COPIA DE PRUEBA' if S.usar_prueba else 'PLANILLA REAL'}: {S.cfg.planilla}").classes("text-xs ruta")
        ui.label("Nadie la cambió desde el reset. Tiene que estar cerrada en Excel.").classes("text-sm text-center")
        with ui.row().classes("justify-center"):
            ui.button("Sí, volver", on_click=lambda: dlg.submit(True)).props("color=negative")
            ui.button("Cancelar", on_click=lambda: dlg.submit(False)).props("flat")
    if not await dlg:
        return
    try:
        await con_espera("Volviendo al respaldo del reset…", preparar.volver_del_reset, S.cfg, datos() / "backups", lambda s: None)
    except Exception as ex:
        ui.notify(f"No pude volver (no se cambió nada): {ex}", type="negative", multi_line=True)
        return
    ui.notify("La planilla volvió a como estaba antes del reset.", type="positive")
    volver_reset_boton.refresh()


async def dialogo_reset():
    from . import preparar
    cfg = S.cfg
    if not puede_escribir(cfg):
        ui.notify("Para vaciar la planilla real elegí 'Cargar en la planilla real' en Configuración → Modo de trabajo (o practicá en la carpeta de prueba).",
                  type="warning", multi_line=True)
        return
    with ui.dialog() as dlg, ui.card().classes("seccion w-[620px] max-w-full items-center gap-3 p-6"):
        ui.label("Reset de la planilla").classes("titulo-seccion")
        ui.label(f"{'COPIA DE PRUEBA' if S.usar_prueba else 'PLANILLA REAL'}: {cfg.planilla}").classes(
            "text-xs text-center ruta " + ("chequeo-ok" if S.usar_prueba else "chequeo-falta"))
        ui.label(f"Se vacían las filas cargadas de {', '.join(config.AREAS[:-1])} y {config.AREAS[-1]}. Antes se guardan dos copias: una en "
                 "la carpeta Respaldos, al lado de la planilla, y otra en datos\\backups. Mientras nadie toque la "
                 "planilla, se puede volver atrás desde Configuración.").classes("text-sm text-center")
        provs = ui.checkbox("También vaciar la hoja Proveedores")
        conf = ui.input("Escribí BORRAR para confirmar").classes("w-64")
        with ui.row().classes("dialogo-botones w-full justify-center gap-3"):
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
    volver_reset_boton.refresh()


# ------------------------------------------------------------------ start

def _puerto_ocupado(p: int) -> bool:
    with socket.socket() as s:
        s.settimeout(0.3)
        return s.connect_ex(("127.0.0.1", p)) == 0


CERRADA_JS = """
document.body.innerHTML = '<div style="height:100vh;display:flex;align-items:center;justify-content:center;'
  + 'font-family:Segoe UI,sans-serif;color:#14147a;font-size:1.3rem;background:#f6f7fc">'
  + 'RUA EZ-Fac se cerró. Ya podés cerrar esta ventana.</div>';
setTimeout(() => window.close(), 200);
"""


async def _salir() -> None:
    """Close the app window (the browser only lets the page close itself), then stop the server.
    If the window can't close, it shows that the app is closed; if something keeps the server alive, force the exit."""
    import asyncio
    import os
    import threading
    ui.run_javascript(CERRADA_JS)
    await asyncio.sleep(0.6)           # let the browser get the script before the server goes away
    threading.Timer(1.5, lambda: os._exit(0)).start()
    app.shutdown()


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


def es_esta_app(linea: str) -> bool:
    """Only a python process running this app is closed (another program on the same port is left alone)."""
    linea = (linea or "").lower()
    return "python" in linea and "appfacturas" in linea


def _linea_de_comando(pid: str) -> str:
    import subprocess
    try:
        return subprocess.run(["powershell", "-NoProfile", "-Command",
                               f"(Get-CimInstance Win32_Process -Filter 'ProcessId={int(pid)}').CommandLine"],
                              capture_output=True, text=True, timeout=10,
                              creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0)).stdout
    except Exception:
        return ""


def _cerrar_viejo(puerto: int) -> bool:
    """Kill the old app still listening on the port (only if it is this app: python running appfacturas)."""
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
        if es_esta_app(_linea_de_comando(pid)):
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
            from . import ventana
            ventana.marcar_cuando_aparezca()     # its own icon in the taskbar, and pinnable as RUA EZ-Fac
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


def mostrar_app(url: str, cfg) -> None:
    """The app is already running: bring its window to the front, or open one if there is none
    (or always another one, with [app] varias_ventanas = true)."""
    from . import ventana
    if not cfg.varias_ventanas and ventana.mostrar_la_abierta():
        return
    abrir_ventana(url)


def main(abrir_navegador: bool = True) -> None:
    import time
    from . import ventana
    cfg = config.cargar()
    url = f"http://127.0.0.1:{cfg.puerto}"
    if not ventana.tomar_candado():
        # another click is already starting the app: wait for it and show its window
        for _ in range(150):
            if _puerto_ocupado(cfg.puerto):
                break
            time.sleep(0.2)
        if abrir_navegador:
            time.sleep(1.5)            # its window opens right after the server
            mostrar_app(url, cfg)
        return
    if _puerto_ocupado(cfg.puerto):
        if _version_corriendo(url) == __version__ or not _cerrar_viejo(cfg.puerto):
            if abrir_navegador:
                mostrar_app(url, cfg)  # same version already running (or could not close the old one)
            return
    if sys.stdout is None:            # pythonw: no console, log to a file
        log = datos() / "app.log"
        log.parent.mkdir(parents=True, exist_ok=True)
        sys.stdout = sys.stderr = open(log, "a", encoding="utf-8", buffering=1)
    if abrir_navegador:
        _abrir_cuando_este_lista(url, cfg.puerto)
    ui.run(host="127.0.0.1", port=cfg.puerto, title=NOMBRE, reload=False, show=False,
           language="es", show_welcome_message=False, favicon=STATIC / "ezfac.ico")
