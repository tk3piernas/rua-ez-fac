"""'Borrar todo': leaves the app as just started. Nothing is deleted: everything MOVES to '_Borrado <fecha>'
folders (same disk, so moving is a rename) and the workbook is emptied in place. It can be undone; emptying
the '_Borrado' folders for good ('Vaciar definitivamente') sends them to the Windows recycle bin.

Safety rules (see docs/DOCUMENTACION.md, "Borrar todo"):
- A CLOSED list of names. The OneDrive root is never listed to decide what goes; only fixed names and the
  app's own backup pattern are looked up.
- Every path is checked before the first move: inside its root, not a link or junction, not protected
  (config.toml, .venv, the code, the tour record, the workbook itself).
- In 'prueba' mode the root must be the test copy (prueba\\OneDrive - Rua Asistencia) and datos\\ is not touched.
- Each move is written to the record (borrado.json) before and after it, so a cut can always be undone.
- Undo and empty only read the record: they never scan folders. Undo never overwrites anything.
"""
from __future__ import annotations

import hashlib
import json
import os
import shutil
import zipfile
from datetime import datetime
from pathlib import Path

from . import config, preparar
from .carpeta_prueba import carpeta_prueba
from .excel import ExcelError, abrir
from .maestro import NOMBRE as NOMBRE_MAESTRO
from .util import guardar_json, sha256

FACTURAS = "Facturas pendientes"
PREFIJO = "_Borrado "
ACTIVOS = ("en curso", "hecho", "deshecho a medias")        # a record in these states blocks a new 'Borrar todo'
QUEDAN_EN_DATOS = ("recorrido.json", "borrado.json", "borrado.json.tmp")   # + app.log* (open while the app runs)
SOLO_NUBE = 0x00400000 | 0x00040000 | 0x00001000           # RECALL_ON_DATA_ACCESS, RECALL_ON_OPEN, OFFLINE
MAX_PAPELERA = 260                                          # MAX_PATH: longer paths do not fit in the recycle bin


class BorradoError(Exception):
    pass


# ------------------------------------------------------------------ where

def _registro(modo: str) -> Path:
    return config.carpeta_app() / "datos" / ("borrado_prueba.json" if modo == "prueba" else "borrado.json")


def carpeta_respaldo_zip() -> Path:
    """%USERPROFILE%\\Documents (APPFACTURAS_DOCUMENTOS only for a sandbox install)."""
    if os.environ.get("APPFACTURAS_DOCUMENTOS"):
        return Path(os.environ["APPFACTURAS_DOCUMENTOS"])
    return Path(os.environ.get("USERPROFILE") or Path.home()) / "Documents"


def _es_link(p: Path) -> bool:
    return p.is_symlink() or os.path.isjunction(p)


def _dentro(p: Path, raiz: Path) -> bool:
    r, b = p.resolve(), raiz.resolve()
    return r != b and b in r.parents


def _raices(cfg, modo: str) -> tuple[Path, Path, Path]:
    """(OneDrive root, workbook, app folder), after checking they are what 'Borrar todo' expects."""
    if modo not in ("prueba", "real"):
        raise BorradoError("modo desconocido")
    if modo == "real" and not cfg.modo_real:
        raise BorradoError("en 'Solo simular' no se puede: elegí 'Cargar en la planilla real' en Configuración")
    if not cfg.base or not cfg.planilla:
        raise BorradoError("no encuentro la carpeta 'Facturas pendientes' o la planilla")
    base, pl = Path(os.path.abspath(cfg.base)), Path(os.path.abspath(cfg.planilla))
    raiz = pl.parent
    if base.name != FACTURAS or base.parent != raiz or pl.name != preparar.NOMBRE_PLANILLA:
        raise BorradoError(f"la planilla y '{FACTURAS}' tienen que estar juntas en la carpeta del OneDrive "
                           f"({pl} / {base}); no sigo")
    if not raiz.name.startswith("OneDrive - ") or not raiz.is_dir() or _es_link(raiz):
        raise BorradoError(f"{raiz} no parece la carpeta del OneDrive; no sigo")
    app = config.carpeta_app().resolve()
    prueba = carpeta_prueba()
    if modo == "prueba" and raiz.resolve() != prueba.resolve():
        raise BorradoError(f"en 'Practicar' solo se borra la carpeta de prueba ({prueba}), no {raiz}")
    if modo == "real" and (raiz.resolve() == app or app in raiz.resolve().parents):
        raise BorradoError("en 'Cargar en la planilla real' la carpeta tiene que ser la del OneDrive, no la de prueba")
    return raiz, pl, app


def _protegidos(app: Path, pl: Path) -> list[Path]:
    return [app, app / "config.toml", app / ".venv", app / "app", app / "datos" / "recorrido.json", pl]


def _toca(p: Path, protegido: Path, app: Path) -> bool:
    """Moving `p` would move `protegido` (it is the same, or one of its folders) or `p` is inside it
    (everything in the app is inside the app folder, so that one only counts as itself)."""
    r, x = p.resolve(), protegido.resolve()
    return r == x or r in x.parents or (x != app and x in r.parents)


def _de_la_app(nombre: str, planilla: Path) -> bool:
    """A workbook backup the app (or the old script) leaves: 'DEUDA A PAGAR ARG - respaldo ....xlsx'."""
    return nombre.startswith(f"{planilla.stem} - respaldo ") and nombre.lower().endswith(".xlsx")


def _carpeta_o_archivos(carpeta: Path, propio) -> tuple[list[Path], list[Path]]:
    """The whole folder if everything in it is the app's; if not, only the app's files (the rest stays)."""
    hijos = sorted(carpeta.iterdir())
    mios = [h for h in hijos if h.is_file() and not _es_link(h) and propio(h.name)]
    if len(mios) == len(hijos):
        return [carpeta], []
    return mios, [h for h in hijos if h not in mios]


# ------------------------------------------------------------------ inventory (no side effects)

def _archivos(p: Path) -> list[Path]:
    if p.is_file():
        return [p]
    out = []
    for dirpath, dirnames, filenames in os.walk(p):
        for n in dirnames + filenames:
            if _es_link(Path(dirpath, n)):
                raise BorradoError(f"hay un acceso directo de carpeta o un link adentro ({Path(dirpath, n)}); no sigo")
        out += [Path(dirpath, f) for f in filenames]
    return out


def inventario(cfg, modo: str) -> dict:
    """What 'Borrar todo' would take, with counts. Raises BorradoError if any path fails a check."""
    raiz, pl, app = _raices(cfg, modo)
    base = raiz / FACTURAS
    items: list[tuple[str, Path]] = []                      # (onedrive / app, path)
    quedan: list[Path] = []
    if base.is_dir():
        items.append(("onedrive", base))
    for nombre, propio in ((preparar.CARPETA_EXPORTADAS, lambda n: n.lower().endswith(".xlsx")),
                           (preparar.CARPETA_RESPALDOS, lambda n: _de_la_app(n, pl))):
        if (raiz / nombre).is_dir() and not _es_link(raiz / nombre):
            mios, otros = _carpeta_o_archivos(raiz / nombre, propio)
            items += [("onedrive", m) for m in mios]
            quedan += otros
    if (raiz / NOMBRE_MAESTRO).is_file():
        items.append(("onedrive", raiz / NOMBRE_MAESTRO))
    items += [("onedrive", f) for f in sorted(raiz.glob(f"{pl.stem} - respaldo *"))
              if f.is_file() and _de_la_app(f.name, pl)]
    if modo == "real":
        datos = app / "datos"
        if datos.is_dir():
            items += [("app", h) for h in sorted(datos.iterdir())
                      if h.name not in QUEDAN_EN_DATOS and not h.name.startswith("app.log")]
        if (app / "prueba").is_dir():
            items.append(("app", app / "prueba"))

    prot = _protegidos(app, pl)
    for donde, p in items:
        if _es_link(p):
            raise BorradoError(f"{p} es un acceso directo o link; no sigo")
        if not _dentro(p, raiz if donde == "onedrive" else app):
            raise BorradoError(f"{p} está fuera de su carpeta; no sigo")
        if any(_toca(p, x, app) for x in prot):
            raise BorradoError(f"{p} está protegido; no sigo")
        if modo == "prueba" and not _dentro(p, carpeta_prueba()):
            raise BorradoError(f"{p} no es de la carpeta de prueba; no sigo")

    archivos = {str(p): _archivos(p) for _, p in items}
    todos = [f for fs in archivos.values() for f in fs]
    provs = [c for a in cfg.areas if (base / a).is_dir() for c in sorted((base / a).iterdir()) if c.is_dir()]
    partes = [f.relative_to(base).parts for f in archivos.get(str(base), [])]
    pdf = sum(1 for t in partes if len(t) >= 3 and t[0] in cfg.areas and t[-1].lower().endswith(".pdf"))
    entrada = sum(1 for t in partes if t[0] == "_Entrada")
    solo_nube = 0
    for f in todos:
        try:
            solo_nube += bool(getattr(f.stat(), "st_file_attributes", 0) & SOLO_NUBE)
        except OSError:
            pass
    filas, provs_planilla = {}, 0
    if pl.is_file():
        libro = abrir(cfg)
        try:
            filas, provs_planilla = preparar.contar_filas(libro)
        finally:
            libro.cerrar()
    return {"modo": modo, "raiz": raiz, "planilla": pl, "app": app,
            "items": [{"donde": d, "ruta": p, "carpeta": p.is_dir(), "archivos": len(archivos[str(p)]),
                       "bytes": sum(f.stat().st_size for f in archivos[str(p)])} for d, p in items],
            "archivos": todos, "quedan": quedan, "pdf": pdf, "carpetas_proveedores": len(provs),
            "en_entrada": entrada, "filas": filas, "proveedores": provs_planilla, "solo_nube": solo_nube,
            "bytes": sum(f.stat().st_size for f in todos)}


# ------------------------------------------------------------------ checks before moving anything

def _tomado(p: Path) -> str:
    """Why the file can not be moved now ("" if it is free): open somewhere (Excel, OneDrive, a viewer)."""
    if os.name != "nt":
        try:
            open(p, "rb").close()
            return ""
        except OSError as e:
            return str(e)
    import ctypes
    from ctypes import wintypes
    k32 = ctypes.WinDLL("kernel32", use_last_error=True)
    k32.CreateFileW.restype = wintypes.HANDLE
    k32.CreateFileW.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD, wintypes.LPVOID,
                                wintypes.DWORD, wintypes.DWORD, wintypes.HANDLE]
    # GENERIC_READ, no sharing (fails if anyone has it open), OPEN_EXISTING, OPEN_NO_RECALL (no download)
    h = k32.CreateFileW(str(p), 0x80000000, 0, None, 3, 0x80 | 0x00100000, None)
    if h in (None, wintypes.HANDLE(-1).value):
        err = ctypes.get_last_error()
        return "está abierto en otro programa" if err in (32, 33) else f"no lo puedo abrir (error {err})"
    k32.CloseHandle(h)
    return ""


def _planilla_sin_guardar(cfg, pl: Path) -> bool:
    if not preparar._abierta_en_excel(cfg, pl):
        return False
    libro = abrir(cfg)
    try:
        return not libro.saved
    finally:
        libro.cerrar()


def chequeo_previo(cfg, inv: dict, carpeta_zip: Path) -> list[str]:
    """Problems that stop 'Borrar todo' before it moves anything ([] = it can go)."""
    problemas = []
    reg = estado(inv["modo"])
    if reg and reg.get("estado") in ACTIVOS:
        problemas.append(f"ya hay un 'Borrar todo' del {reg['fecha']} sin deshacer ni vaciar: hacé eso primero")
    if inv["planilla"].is_file() and _planilla_sin_guardar(cfg, inv["planilla"]):
        problemas.append("la planilla está abierta en Excel con cambios sin guardar: guardala o cerrala")
    for f in inv["archivos"]:
        motivo = _tomado(f)
        if motivo:
            problemas.append(f"{f}: {motivo}")
            if len(problemas) > 20:
                problemas.append("… y más")
                break
    if not carpeta_zip.is_dir():
        problemas.append(f"no existe la carpeta para el respaldo ({carpeta_zip})")
    elif _dentro(carpeta_zip, inv["raiz"]) or carpeta_zip.resolve() == inv["raiz"].resolve() \
            or _dentro(carpeta_zip, inv["app"]) or carpeta_zip.resolve() == inv["app"]:
        problemas.append(f"el respaldo no puede quedar en el OneDrive ni en la app ({carpeta_zip})")
    elif shutil.disk_usage(carpeta_zip).free < inv["bytes"] * 1.1 + 50 * 2**20:
        problemas.append(f"no hay lugar en el disco para el respaldo ({inv['bytes'] // 2**20} MB)")
    return problemas


# ------------------------------------------------------------------ zip backup (outside OneDrive and the app)

def _respaldo_zip(inv: dict, carpeta_zip: Path, marca: str, log) -> Path:
    raiz, app = inv["raiz"], inv["app"]
    extra = [inv["planilla"]] if inv["planilla"].is_file() else []
    if inv["modo"] == "real" and (app / "config.toml").is_file():
        extra.append(app / "config.toml")
    nombre = f"Respaldo RUA EZ-Fac {marca} antes de Borrar todo{' (prueba)' if inv['modo'] == 'prueba' else ''}.zip"
    destino = carpeta_zip / nombre
    if destino.exists():
        raise BorradoError(f"ya existe {destino}")
    lista = []
    with zipfile.ZipFile(destino, "x", zipfile.ZIP_DEFLATED) as z:
        for f in inv["archivos"] + extra:
            arc = ("OneDrive/" + f.relative_to(raiz).as_posix()) if _dentro(f, raiz) else \
                ("app/" + f.relative_to(app).as_posix())
            h = hashlib.sha256()
            with open(f, "rb") as src, z.open(arc, "w", force_zip64=True) as dst:
                for b in iter(lambda: src.read(1 << 20), b""):
                    h.update(b)
                    dst.write(b)
            lista.append({"archivo": arc, "origen": str(f), "sha256": h.hexdigest().upper(), "bytes": f.stat().st_size})
        z.writestr("LISTA.json", json.dumps(lista, ensure_ascii=False, indent=1))
    verificar_zip(destino)
    log(f"Respaldo verificado: {destino} ({len(lista)} archivos)")
    return destino


def verificar_zip(destino: Path) -> int:
    """Each file of LISTA.json is in the zip with the same SHA256. Returns how many."""
    with zipfile.ZipFile(destino) as z:
        if z.testzip() is not None:
            raise BorradoError(f"el respaldo {destino} está dañado")
        lista = json.loads(z.read("LISTA.json"))
        for e in lista:
            h = hashlib.sha256()
            with z.open(e["archivo"]) as f:
                for b in iter(lambda: f.read(1 << 20), b""):
                    h.update(b)
            if h.hexdigest().upper() != e["sha256"]:
                raise BorradoError(f"el respaldo no coincide en {e['archivo']}")
    return len(lista)


# ------------------------------------------------------------------ record

def estado(modo: str) -> dict | None:
    try:
        return json.loads(_registro(modo).read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return None


def _guardar(modo: str, reg: dict) -> None:
    """Atomic and flushed to disk: the record must survive a cut right after it."""
    guardar_json(_registro(modo), reg)


def _mover(src: Path, dst: Path) -> None:
    """Only a rename on the same disk; never copy + delete, never over something."""
    if dst.exists():
        raise BorradoError(f"ya existe {dst}")
    dst.parent.mkdir(parents=True, exist_ok=True)
    if os.stat(src).st_dev != os.stat(dst.parent).st_dev:
        raise BorradoError(f"{src} y {dst.parent} no están en el mismo disco")
    os.rename(src, dst)


def _quitar_vacias(carpeta: Path) -> None:
    """Removes the empty folders left inside `carpeta` (and it, if empty). Never a non-empty one."""
    if not carpeta.is_dir() or _es_link(carpeta):
        return
    for dirpath, _, _ in sorted(os.walk(carpeta), key=lambda t: -len(t[0])):
        try:
            os.rmdir(dirpath)
        except OSError:
            pass


# ------------------------------------------------------------------ borrar

def borrar(cfg, modo: str, carpeta_zip: Path | None = None, log=print) -> dict:
    """The whole 'Borrar todo'. If anything fails before the end, it puts back what it had moved."""
    carpeta_zip = carpeta_zip or carpeta_respaldo_zip()
    inv = inventario(cfg, modo)
    problemas = chequeo_previo(cfg, inv, carpeta_zip)
    if problemas:
        raise BorradoError("No borré nada:\n- " + "\n- ".join(problemas))
    raiz, pl, app = inv["raiz"], inv["planilla"], inv["app"]
    marca = f"{datetime.now():%Y-%m-%d %H%M}"
    if (raiz / (PREFIJO + marca)).exists() or (app / (PREFIJO + marca)).exists() \
            or any(carpeta_zip.glob(f"Respaldo RUA EZ-Fac {marca} *")):
        marca = f"{datetime.now():%Y-%m-%d %H%M%S}"
    zip_ = _respaldo_zip(inv, carpeta_zip, marca, log)
    b_od, b_app = raiz / (PREFIJO + marca), app / (PREFIJO + marca)
    reg = {"fecha": datetime.now().isoformat(timespec="seconds"), "modo": modo, "raiz": str(raiz), "app": str(app),
           "carpetas": [str(b_od)] + ([str(b_app)] if modo == "real" else []), "zip": str(zip_),
           "movidos": [], "planilla": None, "creados": {}, "estado": "en curso", "avisos": []}
    _guardar(modo, reg)
    try:
        for it in inv["items"]:
            src = it["ruta"]
            dst = (b_od / src.relative_to(raiz)) if it["donde"] == "onedrive" else (b_app / src.relative_to(app))
            reg["movidos"].append({"desde": str(src), "hacia": str(dst), "hecho": False})
            _guardar(modo, reg)
            _mover(src, dst)
            reg["movidos"][-1]["hecho"] = True
            _guardar(modo, reg)
            log(f"MOVIDO {src} -> {dst}")
        if pl.is_file():
            copia = b_od / pl.name
            b_od.mkdir(parents=True, exist_ok=True)
            shutil.copy2(pl, copia)
            reg["planilla"] = {"ruta": str(pl), "copia": str(copia), "sha_antes": sha256(pl), "sha_vaciada": None}
            if sha256(copia) != reg["planilla"]["sha_antes"]:
                raise BorradoError("la copia de la planilla no salió igual")
            _guardar(modo, reg)
            _vaciar_planilla(cfg, log)
            reg["planilla"]["sha_vaciada"] = sha256(pl)
            _guardar(modo, reg)
    except Exception as e:
        log(f"FALLÓ el borrado: {e}. Devuelvo lo que ya moví.")
        try:
            _revertir(reg, log)
        except Exception as e2:
            raise BorradoError(f"No se pudo terminar ({e}) y no pude devolver todo ({e2}): "
                               "tocá 'Deshacer el borrado total'") from e2
        raise BorradoError(f"No se pudo terminar y devolví todo a su lugar: {e}") from e
    reg["creados"] = _rearmar(cfg, reg, log)
    reg["estado"] = "hecho"
    _guardar(modo, reg)
    log(f"BORRAR TODO listo: {len(reg['movidos'])} cosas en {', '.join(reg['carpetas'])}")
    return reg


def _vaciar_planilla(cfg, log) -> None:
    libro = abrir(cfg)
    try:
        if not libro.saved:
            raise ExcelError("la planilla está abierta con cambios sin guardar")
        preparar.vaciar_libro(libro, True, log)
        libro.guardar()
    finally:
        libro.cerrar()


def _revertir(reg: dict, log) -> None:
    """Back from a cut 'borrar': the workbook from its copy (if it changed) and the moves in reverse."""
    p = reg.get("planilla")
    if p and Path(p["copia"]).is_file() and sha256(Path(p["ruta"])) != p["sha_antes"]:
        shutil.copyfile(p["copia"], p["ruta"])
    if p and Path(p["copia"]).is_file() and sha256(Path(p["ruta"])) == p["sha_antes"]:
        Path(p["copia"]).unlink()
    for m in reversed(reg["movidos"]):
        src, dst = Path(m["hacia"]), Path(m["desde"])
        if src.exists() and not dst.exists():
            os.rename(src, dst)
            log(f"DEVUELTO {dst}")
    for c in reg["carpetas"]:
        _quitar_vacias(Path(c))
    reg["estado"] = "revertido" if not any(Path(c).exists() for c in reg["carpetas"]) else "en curso"
    _guardar(reg["modo"], reg)


def _rearmar(cfg, reg: dict, log) -> dict:
    """Setup again (empty folders and proveedores.csv), so the app keeps working. Recorded so undo can
    take them away. A failure here does not undo the 'Borrar todo'."""
    try:
        plan = preparar.plan_setup(cfg)
        plan["hojas"] = []
        preparar.ejecutar_setup(cfg, plan)
        return {"archivos": {str(f): sha256(f) for f in plan["archivos"] if f.is_file()}}
    except Exception as e:
        reg["avisos"].append(f"no pude rearmar las carpetas vacías ({e}): tocá Setup en Configuración")
        log(reg["avisos"][-1])
        return {"archivos": {}}


# ------------------------------------------------------------------ deshacer

def _devolver(src: Path, dst: Path, avisos: list[str]) -> None:
    """Puts `src` back at `dst`. If `dst` exists (a folder Setup made again) it goes inside; a file that
    appeared after the 'Borrar todo' is never overwritten: the old one stays in _Borrado."""
    if not src.exists():
        if not dst.exists():
            avisos.append(f"no encuentro {src}")
        return
    try:
        if not dst.exists():
            dst.parent.mkdir(parents=True, exist_ok=True)
            os.rename(src, dst)
        elif src.is_dir() and dst.is_dir() and not _es_link(src) and not _es_link(dst):
            for h in sorted(src.iterdir()):
                _devolver(h, dst / h.name, avisos)
            try:
                src.rmdir()
            except OSError:
                pass
        else:
            avisos.append(f"{dst} apareció después del borrado: no lo piso; el anterior quedó en {src}")
    except OSError as e:
        avisos.append(f"no pude devolver {dst} ({e}); quedó en {src}")


def _modo_permitido(cfg, modo: str) -> None:
    if modo == "real" and not cfg.modo_real:
        raise BorradoError("en 'Solo simular' no se puede: elegí 'Cargar en la planilla real' en Configuración")


def deshacer(cfg, modo: str, log=print) -> dict:
    """Undo of the last 'Borrar todo': everything back to its place (from the record only)."""
    _modo_permitido(cfg, modo)
    reg = estado(modo)
    if not reg or reg.get("estado") not in ACTIVOS:
        raise BorradoError("no hay un 'Borrar todo' para deshacer")
    avisos: list[str] = []
    p = reg.get("planilla")
    volver_planilla = False
    if p and not p.get("devuelta"):
        pl, copia = Path(p["ruta"]), Path(p["copia"])
        if not copia.is_file():
            avisos.append(f"no encuentro la copia de la planilla ({copia}); la planilla queda como está")
        elif p.get("sha_vaciada") is None or not pl.is_file() or sha256(pl) != p["sha_vaciada"]:
            avisos.append(f"la planilla cambió después del borrado: no la piso. La de antes está en {copia}")
        elif preparar._abierta_en_excel(cfg, pl):
            raise BorradoError("la planilla está abierta en Excel: cerrala y volvé a probar")
        else:
            volver_planilla = True
    if volver_planilla:
        shutil.copyfile(copia, pl)
        if sha256(pl) != p["sha_antes"]:
            raise BorradoError("la planilla no volvió igual; la copia sigue en " + str(copia))
        copia.unlink()
        p["devuelta"] = True
        _guardar(modo, reg)
        log(f"PLANILLA vuelve a como estaba antes del borrado ({pl})")
    for f, sha in reg.get("creados", {}).get("archivos", {}).items():
        f = Path(f)
        if f.is_file() and sha256(f) == sha:
            f.unlink()                                   # the empty proveedores.csv Setup made, unchanged
    for m in reversed(reg["movidos"]):
        _devolver(Path(m["hacia"]), Path(m["desde"]), avisos)
    for c in reg["carpetas"]:
        _quitar_vacias(Path(c))
    queda = [c for c in reg["carpetas"] if Path(c).exists()]
    reg["estado"] = "deshecho a medias" if queda else "deshecho"
    reg["avisos"] = avisos
    _guardar(modo, reg)
    for a in avisos:
        log(f"AVISO {a}")
    log("BORRAR TODO deshecho" + (f"; queda en {', '.join(queda)}" if queda else ""))
    return reg


# ------------------------------------------------------------------ vaciar definitivamente

def _a_la_papelera(p: Path) -> None:
    """Windows recycle bin (OneDrive sends it to its own recycle bin). Never rmtree."""
    import pythoncom
    from win32com.shell import shell, shellcon
    pythoncom.CoInitialize()                         # it runs in a worker thread
    flags = shellcon.FOF_ALLOWUNDO | shellcon.FOF_NOCONFIRMATION | shellcon.FOF_WANTNUKEWARNING
    res, abortado = shell.SHFileOperation((0, shellcon.FO_DELETE, str(p), None, flags, None, None))
    if res or abortado:
        raise BorradoError(f"no pude mandar {p} a la papelera (código {res})")


def vaciar(cfg, modo: str, confirmacion: str, log=print) -> dict:
    """Sends the '_Borrado' folders of the record to the recycle bin. Only those, and only if they are
    still the record's ('_Borrado ...' right inside the OneDrive or app folder)."""
    _modo_permitido(cfg, modo)
    if confirmacion != "VACIAR":
        raise BorradoError("para vaciar hay que escribir VACIAR")
    reg = estado(modo)
    if not reg or reg.get("estado") not in ("hecho", "deshecho a medias"):
        raise BorradoError("no hay carpetas de 'Borrar todo' para vaciar")
    padres = {str(Path(reg["raiz"])), str(Path(reg["app"]))}
    carpetas = [c for c in map(Path, reg["carpetas"]) if c.exists()]
    for c in carpetas:
        if not c.name.startswith(PREFIJO) or str(c.parent) not in padres or _es_link(c):
            raise BorradoError(f"{c} no es una carpeta de 'Borrar todo'; no la toco")
    # Windows can not recycle a path of MAX_PATH or more: it would offer to delete it for good instead
    largas = [f for c in carpetas for f in [c, *c.rglob("*")] if len(str(f)) >= MAX_PAPELERA]
    if largas:
        raise BorradoError(f"{len(largas)} archivos tienen la ruta demasiado larga para la papelera de Windows "
                           f"(por ejemplo {largas[0]}). No vacié nada: todo sigue en _Borrado.")
    for c in carpetas:
        _a_la_papelera(c)
        if c.exists():
            raise BorradoError(f"{c} sigue ahí después de mandarla a la papelera")
        log(f"A LA PAPELERA {c}")
    reg["estado"] = "vaciado"
    _guardar(modo, reg)
    return reg
