"""Runs a batch's file plan (stage 3) and undoes it.

Every step is written to diario.json as soon as it is done, so the journal is exact even if
the app closes halfway. Undo walks the journal backwards and only touches files that are
still exactly as the app left them (same SHA256)."""
from __future__ import annotations

import json
import shutil
import threading
import time
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path

from .model import Lote

HIST_MEMORIA = "memoria_proveedores.json"


from .util import guardar_json, relativa, ruta_libre, sha256, url_sharepoint


class YaDeshecho(Exception):
    """Undo asked twice for the same batch: nothing is touched."""


class YaConfirmado(Exception):
    """Confirm asked twice for the same batch (double click): nothing is done the second time."""


_ocupados: set[str] = set()              # batch folders with a confirm or undo running in this process
_candado = threading.Lock()


@contextmanager
def _uno_por_vez(carpeta: Path, error):
    clave = str(carpeta.resolve()).lower()
    with _candado:
        if clave in _ocupados:
            raise error(f"el lote {carpeta.name} ya se está procesando")
        _ocupados.add(clave)
    try:
        yield
    finally:
        with _candado:
            _ocupados.discard(clave)

ORDEN = {"proveedor_nuevo": 0, "mover": 1, "copiar": 1, "mover_zip": 2, "a_duplicados": 3, "txt_procesado": 4}


class Diario:
    def __init__(self, carpeta: Path):
        self.carpeta = carpeta
        self.ruta = carpeta / "diario.json"
        self.pasos: list[dict] = json.loads(self.ruta.read_text(encoding="utf-8")) if self.ruta.exists() else []

    def anotar(self, paso: dict) -> None:
        paso["hora"] = datetime.now().isoformat(timespec="seconds")
        self.pasos.append(paso)
        guardar_json(self.ruta, self.pasos)

    def completar_cortados(self) -> int:
        """A move or copy announced ('moviendo') but not confirmed: the app closed in between. If the file is at
        its destination with the same SHA256 (and, for a move, gone from where it was), it is journaled as done,
        so undo brings it back and continuing does not look for it in _Entrada. Returns how many."""
        hechos = {(x["op"], x.get("de"), x.get("a")) for x in self.pasos if x["op"] in ("mover", "copiar")}
        n = 0
        for x in list(self.pasos):
            if x["op"] != "moviendo" or (x["como"], x["de"], x["a"]) in hechos:
                continue
            de, a = Path(x["de"]), Path(x["a"])
            if a.is_file() and sha256(a) == x["sha256"] and (x["como"] == "copiar" or not de.exists()):
                self.anotar({"op": x["como"], "de": x["de"], "a": x["a"], "sha256": x["sha256"], "recuperado": True})
                n += 1
        return n


def estado(carpeta: Path) -> dict:
    p = carpeta / "estado.json"
    return json.loads(p.read_text(encoding="utf-8")) if p.exists() else {}


def _guardar_estado(carpeta: Path, **kw) -> None:
    d = estado(carpeta)
    d.update(kw)
    guardar_json(carpeta / "estado.json", d)


def _mover(de: Path, a: Path, intentos: int = 4) -> None:
    """Move with a few retries: OneDrive may hold a file for a moment while it syncs."""
    for i in range(intentos):
        try:
            shutil.move(str(de), str(a))
            return
        except PermissionError:
            if i == intentos - 1:
                raise
            time.sleep(1.5)


class Ejecutor:
    def __init__(self, cfg, lote: Lote, carpeta_lote: Path, log=print, con_excel: bool = True,
                 guardar_cambios_usuario: bool = False, modo: str = ""):
        self.cfg, self.lote, self.carpeta = cfg, lote, carpeta_lote
        self.modo = modo                  # "prueba" / "real": shown in Historial
        self.con_excel = con_excel
        self.guardar_cambios_usuario = guardar_cambios_usuario
        self.diario = Diario(carpeta_lote)
        self.log = log
        self.errores: list[str] = []
        self.hechos: list[str] = []

    def _carpeta(self, d: Path) -> None:
        faltan = []
        x = d
        while not x.exists():
            faltan.append(x)
            x = x.parent
        for c in reversed(faltan):
            c.mkdir()
            self.diario.anotar({"op": "carpeta_creada", "ruta": str(c)})

    def _rel(self, p: Path) -> str:
        return relativa(p, self.cfg.base)

    def ejecutar(self) -> dict:
        with _uno_por_vez(self.carpeta, YaConfirmado):
            ya = estado(self.carpeta).get("estado", "")
            if ya:
                raise YaConfirmado(f"el lote {self.carpeta.name} ya está {ya}; no lo vuelvo a confirmar")
            return self._ejecutar()

    def _ejecutar(self) -> dict:
        self.carpeta.mkdir(parents=True, exist_ok=True)
        _guardar_estado(self.carpeta, estado="en curso", inicio=datetime.now().isoformat(timespec="seconds"),
                        modo=self.modo, base=str(self.cfg.base or ""), planilla=str(self.cfg.planilla or ""))
        # the plan is final: saved before any step, so a batch cut halfway can still be undone and continued
        self._guardar_lote()
        self._pasos()
        return self._terminar()

    def _guardar_lote(self) -> None:
        guardar_json(self.carpeta / "lote_confirmado.json", self.lote, default=_a_json)

    def _pasos(self) -> None:
        """File steps of the plan that the journal does not have yet (all of them the first time)."""
        self.diario.completar_cortados()
        pasos = sorted((p for p in self.lote.plan if p["op"] in ORDEN and not _hecho(p, self.diario.pasos)),
                       key=lambda p: ORDEN[p["op"]])
        for p in pasos:
            try:
                getattr(self, "_" + p["op"])(p)
            except Exception as ex:
                msg = f"{p['op']} {Path(p.get('de', p.get('nombre', ''))).name}: {ex}"
                self.errores.append(msg)
                self.diario.anotar({"op": "error", "paso": p, "error": str(ex)})
                self.log(f"ERROR  {msg}")

    def _terminar(self) -> dict:
        """Rows (only those not written yet, and never those whose file did not move), memory and state."""
        from . import excel
        cortada = ""
        if self.con_excel and excel.tanda_cortada(self.diario.pasos):
            try:
                excel.reconocer_cortada(self.cfg, self.diario, self.log)
            except Exception as ex:                # decided later: never write those rows a second time
                cortada = str(ex)
                self.log(f"EXCEL pendiente: {ex}")
        sin_archivo = _sin_archivo(self.lote, self.diario.pasos)
        a_escribir = _sin_lo_escrito(self.lote, self.diario.pasos, sin_archivo)
        ya_escrito = bool(_escritas(self.diario.pasos)[0])
        excel_estado, excel_msg, excel_res = ("escrito" if ya_escrito else "no corresponde"), "", {"filas": 0, "proveedores": 0}
        if a_escribir.filas or a_escribir.proveedores:
            if cortada:
                excel_estado, excel_msg = "pendiente", cortada
            elif not self.con_excel:
                excel_estado = "pendiente"
            else:
                excel_estado, excel_msg, excel_res = escribir_excel(self.cfg, a_escribir, self.diario, self.carpeta,
                                                                    self.guardar_cambios_usuario, self.log)
        memoria = 0
        if not estado(self.carpeta).get("memoria_guardada"):
            memoria = self._guardar_memoria()
        pendientes = len(a_escribir.filas) + len(a_escribir.proveedores) if excel_estado == "pendiente" else 0
        filas_sin_archivo = [r for r in self.lote.filas if not r.excluida and r.factura_id in sin_archivo]
        if filas_sin_archivo:
            self.log(f"AVISO {len(filas_sin_archivo)} filas no se escribieron: su archivo no se movió")
        _guardar_estado(self.carpeta, estado="confirmado" if not self.errores and not sin_archivo else "confirmado con errores",
                        fin=datetime.now().isoformat(timespec="seconds"), errores=self.errores,
                        excel=excel_estado, excel_mensaje=excel_msg, excel_pendiente=pendientes, excel_escrito=excel_res,
                        archivos_pendientes=sorted(sin_archivo), filas_sin_archivo=len(filas_sin_archivo),
                        memoria_guardada=True)
        self._guardar_lote()             # destinations renamed on the way ("name (2).pdf") and their links
        return {"hechos": self.hechos, "errores": self.errores, "excel_pendiente": pendientes,
                "excel": excel_estado, "excel_mensaje": excel_msg, "excel_escrito": excel_res, "memoria": memoria,
                "sin_archivo": len(filas_sin_archivo)}

    def _guardar_memoria(self) -> int:
        """Provider details of this batch (.txt and invoices) into PROVEEDORES RUA.xlsx; never stops the batch."""
        from . import maestro
        fichas = [(p["valores"], p.get("solo_vacios", False)) for p in self.lote.proveedores if not p.get("de_memoria")]
        if not fichas:
            return 0
        try:
            return len(maestro.guardar(self.cfg, fichas, self.carpeta / HIST_MEMORIA, f"lote {self.carpeta.name}", self.log))
        except Exception as ex:
            self.log(f"MEMORIA no pude guardar los datos de proveedores: {ex}")
            self.errores.append(f"memoria de proveedores: {ex}")
            return 0

    # ---------------------------------------------------------------- steps
    def _proveedor_nuevo(self, p: dict) -> None:
        csv = self.cfg.csv_proveedores
        linea = f"{p['cuit']};{p['area']};{p['nombre']};{p['carpeta']}"
        if not csv.exists():
            self._carpeta(csv.parent)
            csv.write_text("CUIT;Area;Proveedor;Carpeta\r\n", encoding="utf-8-sig")
            self.diario.anotar({"op": "archivo_creado", "ruta": str(csv), "sha256": sha256(csv)})
        datos = csv.read_bytes()
        prefijo = b"" if not datos or datos.endswith(b"\n") else b"\r\n"
        with open(csv, "ab") as fh:
            fh.write(prefijo + (linea + "\r\n").encode("utf-8"))
        self.diario.anotar({"op": "csv_linea", "csv": str(csv), "linea": linea})
        if p["area"]:
            self._carpeta(self.cfg.base / p["area"] / p["carpeta"])
        self.hechos.append(f"Proveedor nuevo: {p['nombre']} ({p['area']}\\{p['carpeta']})")
        self.log(f"PROV {linea}")

    def _destino_libre(self, p: dict) -> Path:
        a = Path(p["a"])
        if a.exists():
            nuevo = ruta_libre(a)
            self._renombrar_destino(p.get("factura"), a, nuevo)
            a = nuevo
        self._carpeta(a.parent)
        return a

    def _renombrar_destino(self, fid, viejo: Path, nuevo: Path) -> None:
        """Someone saved a file with that name after the review: use 'name (2).pdf' and fix the link."""
        f = next((x for x in self.lote.facturas if x.id == fid), None)
        if f:
            f.destino = nuevo
        if self.cfg.url_base:
            url = url_sharepoint(self.cfg.url_base, nuevo.parent.parent.name, nuevo.parent.name, nuevo.name)
            if f:
                f.url = url
            for r in self.lote.filas:        # also when retrying from lote_confirmado.json (no invoices there)
                if r.factura_id == fid:
                    r.url = url
        self.log(f"AVISO {viejo.name} ya existía: lo guardo como {nuevo.name}")

    def _mover(self, p: dict) -> None:
        de = Path(p["de"])
        h = sha256(de)
        a = self._destino_libre(p)
        self.diario.anotar({"op": "moviendo", "como": "mover", "de": str(de), "a": str(a), "sha256": h})
        _mover(de, a)
        self.diario.anotar({"op": "mover", "de": str(de), "a": str(a), "sha256": h})
        self.hechos.append(f"{de.name} -> {self._rel(a)}")
        self.log(f"PDF  {de.name}  ->  {self._rel(a)}")

    def _copiar(self, p: dict) -> None:
        de = Path(p["de"])
        h = sha256(de)
        a = self._destino_libre(p)
        self.diario.anotar({"op": "moviendo", "como": "copiar", "de": str(de), "a": str(a), "sha256": h})
        shutil.copy2(de, a)
        self.diario.anotar({"op": "copiar", "de": str(de), "a": str(a), "sha256": sha256(a)})
        self.hechos.append(f"{de.name} (del zip) -> {self._rel(a)}")
        self.log(f"PDF  {de.name} (zip)  ->  {self._rel(a)}")

    def _mover_simple(self, de: Path, a: Path, op: str, etiqueta: str) -> None:
        h = sha256(de)
        a = ruta_libre(a)
        self._carpeta(a.parent)
        self.diario.anotar({"op": "moviendo", "como": op, "de": str(de), "a": str(a), "sha256": h})
        _mover(de, a)
        self.diario.anotar({"op": op, "de": str(de), "a": str(a), "sha256": h})
        self.hechos.append(f"{de.name} -> {self._rel(a)}")
        self.log(f"{etiqueta} {de.name}  ->  {self._rel(a)}")

    def _mover_zip(self, p: dict) -> None:
        self._mover_simple(Path(p["de"]), Path(p["a"]), "mover", "ZIP ")

    def _a_duplicados(self, p: dict) -> None:
        de = Path(p["de"])
        self._mover_simple(de, self.cfg.entrada / "_Duplicados" / de.name, "mover", "DUP ")

    def _txt_procesado(self, p: dict) -> None:
        self._mover_simple(Path(p["de"]), Path(p["a"]), "mover", "TXT ")


def escribir_excel(cfg, lote, diario, carpeta_lote: Path, guardar_cambios_usuario=False, log=print):
    """(estado, mensaje, resultado). If Excel can't be used the rows stay pending (retry or clipboard)."""
    from . import excel
    try:
        res = excel.escribir(cfg, lote, diario, carpeta_lote.parent.parent / "backups",
                             guardar_cambios_usuario=guardar_cambios_usuario, log=log)
        return "escrito", "", res
    except excel.LibroConCambios as e:
        log(f"EXCEL pendiente: {e}")
        return "pendiente", str(e), {"filas": 0, "proveedores": 0}
    except Exception as e:
        log(f"EXCEL pendiente: {e}")
        return "pendiente", str(e), {"filas": 0, "proveedores": 0}


def continuar(cfg, carpeta_lote: Path, guardar_cambios_usuario=False, log=print, con_archivos=True) -> dict:
    """Finishes a confirmed batch: the file steps that failed or never ran (batch cut halfway), then the rows
    not written yet. Never repeats what the journal already has, so it can be retried any number of times."""
    from .model import cargar_lote
    with _uno_por_vez(carpeta_lote, YaConfirmado):
        est = estado(carpeta_lote)
        if est.get("estado") not in ("en curso", "confirmado", "confirmado con errores"):
            raise YaConfirmado(f"el lote {carpeta_lote.name} está {est.get('estado') or 'sin confirmar'}")
        lote = cargar_lote(carpeta_lote / "lote_confirmado.json")
        ej = Ejecutor(cfg, lote, carpeta_lote, log=log, guardar_cambios_usuario=guardar_cambios_usuario,
                      modo=est.get("modo", ""))
        if con_archivos:
            ej._pasos()
        return ej._terminar()


def reintentar_excel(cfg, carpeta_lote: Path, guardar_cambios_usuario=False, log=print) -> tuple[str, str]:
    """Writes the pending rows of a confirmed batch (only those not written yet)."""
    r = continuar(cfg, carpeta_lote, guardar_cambios_usuario, log, con_archivos=False)
    return r["excel"], r["excel_mensaje"]


def _hecho(p: dict, pasos: list[dict]) -> bool:
    """This plan step is already in the journal."""
    if p["op"] == "proveedor_nuevo":
        linea = f"{p['cuit']};{p['area']};{p['nombre']};{p['carpeta']}"
        return any(x["op"] == "csv_linea" and x.get("linea") == linea for x in pasos)
    return any(x["op"] in ("mover", "copiar") and x.get("de") == p.get("de") for x in pasos)


def _sin_archivo(lote, pasos: list[dict]) -> set[str]:
    """Invoices whose PDF was not saved in its folder: their rows would point to nothing."""
    return {p["factura"] for p in lote.plan
            if p["op"] in ("mover", "copiar") and p.get("factura") and not _hecho(p, pasos)}


def _escritas(pasos: list[dict]) -> tuple[set[str], bool]:
    """(invoices with rows already saved in the workbook, provider rows already saved)."""
    from . import excel
    guardados = excel.pasos_guardados(pasos)
    return ({x.get("factura") for x in guardados if x["op"] == "fila_excel"},
            any(x["op"] == "proveedor_excel" for x in guardados))


def _sin_lo_escrito(lote, pasos: list[dict], sin_archivo: set[str]):
    from dataclasses import replace
    facturas, provs = _escritas(pasos)
    return replace(lote, filas=[r for r in lote.filas if not r.excluida and r.factura_id not in facturas
                                and r.factura_id not in sin_archivo],
                   proveedores=[] if provs else lote.proveedores)


def ocupado(carpeta: Path) -> bool:
    return str(carpeta.resolve()).lower() in _ocupados


def etiqueta(carpeta: Path) -> str:
    """State for Historial: a batch left "en curso" by a closed app shows as cut halfway."""
    e = estado(carpeta).get("estado", "")
    if e == "en curso" and not ocupado(carpeta):
        return "quedó a medias"
    return e or "solo leído / simulado"


def donde(carpeta: Path) -> tuple[str, str]:
    """("prueba" | "real" | "", workbook) where the batch was loaded. Old batches: guessed from the journal."""
    e = estado(carpeta)
    planilla = e.get("planilla") or next((x["libro"] for x in Diario(carpeta).pasos if x.get("libro")), "")
    modo = e.get("modo", "")
    if not modo and planilla:
        from .config import carpeta_app
        modo = "prueba" if str(planilla).lower().startswith(str(carpeta_app() / "prueba").lower()) else "real"
    return modo, planilla


def _a_json(o):
    from .model import a_json
    return a_json(o)


# ---------------------------------------------------------------------- undo

def lotes_confirmados(carpeta_lotes: Path) -> list[Path]:
    if not carpeta_lotes.exists():
        return []
    return sorted((d for d in carpeta_lotes.iterdir()
                   if d.is_dir() and estado(d).get("estado", "").startswith(("confirmado", "en curso"))
                   and not ocupado(d)), key=lambda d: d.name)


def ultimo_confirmado(carpeta_lotes: Path) -> Path | None:
    ls = lotes_confirmados(carpeta_lotes)
    return ls[-1] if ls else None


def deshacer(carpeta_lote: Path, log=print, cfg=None) -> list[str]:
    """Walks the journal backwards: first the Excel rows, then the files.
    Returns the problems (things that changed after the batch). If the workbook can't be
    opened, nothing is undone (so rows never point to files that moved back)."""
    with _uno_por_vez(carpeta_lote, YaDeshecho):
        e = estado(carpeta_lote)
        if e.get("estado") == "deshecho":
            if e.get("memoria_pendiente"):
                return _deshacer_memoria(carpeta_lote, log, cfg)
            raise YaDeshecho(f"el lote {carpeta_lote.name} ya estaba deshecho")
        return _deshacer(carpeta_lote, log, cfg)


def _deshacer_memoria(carpeta_lote: Path, log, cfg) -> list[str]:
    """Second try of the provider-memory part (the rest of the batch was already undone)."""
    from . import maestro
    if cfg is None:
        from .config import cargar
        cfg = cargar()
    problemas = maestro.deshacer(cfg, carpeta_lote / HIST_MEMORIA, todo=True, log=log)
    _guardar_estado(carpeta_lote, memoria_pendiente=False)
    return problemas


def _deshacer(carpeta_lote: Path, log, cfg) -> list[str]:
    from . import excel
    diario = Diario(carpeta_lote)
    diario.completar_cortados()                  # a file moved right before the app closed also goes back
    if excel.tanda_cortada(diario.pasos):        # rows saved right before the app closed also go
        if cfg is None:
            from .config import cargar
            cfg = cargar()
        excel.reconocer_cortada(cfg, diario, log)
    pasos = diario.pasos
    problemas: list[str] = []
    memoria_pendiente = False
    if excel.pasos_guardados(pasos):
        if cfg is None:
            from .config import cargar
            cfg = cargar()
        problemas += excel.deshacer_excel(cfg, pasos, log=log)
    if (carpeta_lote / HIST_MEMORIA).is_file():
        from . import maestro
        if cfg is None:
            from .config import cargar
            cfg = cargar()
        try:
            problemas += maestro.deshacer(cfg, carpeta_lote / HIST_MEMORIA, todo=True, log=log)
        except Exception as ex:
            # kept pending: touching Deshacer again retries only this part
            memoria_pendiente = True
            problemas.append(f"memoria de proveedores (quedó pendiente: volvé a tocar Deshacer): {ex}")
    for p in reversed(pasos):
        op = p["op"]
        try:
            if op == "mover":
                de, a = Path(p["de"]), Path(p["a"])
                if not a.exists():
                    problemas.append(f"{a.name}: ya no está en {a.parent}; no lo puedo devolver")
                elif de.exists():
                    problemas.append(f"{de.name}: ya hay un archivo con ese nombre en {de.parent}; dejo {a.name} donde está")
                elif p.get("sha256") and sha256(a) != p["sha256"]:
                    problemas.append(f"{a.name} cambió después del lote; no lo muevo")
                else:
                    de.parent.mkdir(parents=True, exist_ok=True)
                    _mover(a, de)
                    log(f"VUELVE {a.name} -> {de.parent.name}\\{de.name}")
            elif op in ("copiar", "archivo_creado"):
                a = Path(p.get("a") or p["ruta"])
                if a.exists() and sha256(a) == p["sha256"]:
                    a.unlink()
                    log(f"BORRO  {a.name} (lo había creado el lote)")
                elif a.exists():
                    problemas.append(f"{a.name} cambió después del lote; no lo borro")
            elif op == "carpeta_creada":
                c = Path(p["ruta"])
                if c.exists() and not any(c.iterdir()):
                    c.rmdir()
                    log(f"BORRO  carpeta vacía {c.name}")
                elif c.exists():
                    problemas.append(f"La carpeta {c} no está vacía; la dejo")
            elif op == "csv_linea":
                csv = Path(p["csv"])
                datos = csv.read_bytes().decode("utf-8-sig")
                lineas = datos.splitlines(keepends=True)
                idx = max((i for i, l in enumerate(lineas) if l.rstrip("\r\n") == p["linea"]), default=None)
                if idx is None:
                    problemas.append(f"proveedores.csv ya no tiene la línea '{p['linea']}'")
                else:
                    del lineas[idx]
                    bom = datos != csv.read_bytes().decode("utf-8")
                    csv.write_bytes(("﻿" if bom else "").encode("utf-8") + "".join(lineas).encode("utf-8"))
                    log(f"CSV    saco '{p['linea']}'")
        except Exception as ex:
            problemas.append(f"{op}: {ex}")
    _guardar_estado(carpeta_lote, estado="deshecho", deshecho=datetime.now().isoformat(timespec="seconds"),
                    problemas_deshacer=problemas, memoria_pendiente=memoria_pendiente)
    return problemas
