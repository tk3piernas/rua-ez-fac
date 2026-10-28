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


from .util import ruta_libre, sha256, url_sharepoint


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
        self.carpeta.mkdir(parents=True, exist_ok=True)
        self.ruta.write_text(json.dumps(self.pasos, ensure_ascii=False, indent=1), encoding="utf-8")


def estado(carpeta: Path) -> dict:
    p = carpeta / "estado.json"
    return json.loads(p.read_text(encoding="utf-8")) if p.exists() else {}


def _guardar_estado(carpeta: Path, **kw) -> None:
    d = estado(carpeta)
    d.update(kw)
    (carpeta / "estado.json").write_text(json.dumps(d, ensure_ascii=False, indent=1), encoding="utf-8")


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
                 guardar_cambios_usuario: bool = False):
        self.cfg, self.lote, self.carpeta = cfg, lote, carpeta_lote
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
        try:
            return str(p.relative_to(self.cfg.base))
        except ValueError:
            return str(p)

    def ejecutar(self) -> dict:
        with _uno_por_vez(self.carpeta, YaConfirmado):
            ya = estado(self.carpeta).get("estado", "")
            if ya:
                raise YaConfirmado(f"el lote {self.carpeta.name} ya está {ya}; no lo vuelvo a confirmar")
            return self._ejecutar()

    def _ejecutar(self) -> dict:
        _guardar_estado(self.carpeta, estado="en curso", inicio=datetime.now().isoformat(timespec="seconds"))
        pasos = sorted((p for p in self.lote.plan if p["op"] in ORDEN), key=lambda p: ORDEN[p["op"]])
        for p in pasos:
            try:
                getattr(self, "_" + p["op"])(p)
            except Exception as ex:
                msg = f"{p['op']} {Path(p.get('de', p.get('nombre', ''))).name}: {ex}"
                self.errores.append(msg)
                self.diario.anotar({"op": "error", "paso": p, "error": str(ex)})
                self.log(f"ERROR  {msg}")
        excel_estado, excel_msg, excel_res = "no corresponde", "", {"filas": 0, "proveedores": 0}
        if self.lote.filas or self.lote.proveedores:
            if not self.con_excel:
                excel_estado = "pendiente"
            else:
                excel_estado, excel_msg, excel_res = escribir_excel(self.cfg, self.lote, self.diario, self.carpeta,
                                                                    self.guardar_cambios_usuario, self.log)
        memoria = self._guardar_memoria()
        pendientes = [] if excel_estado != "pendiente" else [1] * (len(self.lote.filas) + len(self.lote.proveedores))
        _guardar_estado(self.carpeta, estado="confirmado" if not self.errores else "confirmado con errores",
                        fin=datetime.now().isoformat(timespec="seconds"), errores=self.errores,
                        excel=excel_estado, excel_mensaje=excel_msg, excel_pendiente=len(pendientes), excel_escrito=excel_res)
        (self.carpeta / "lote_confirmado.json").write_text(
            json.dumps(self.lote, default=_a_json, ensure_ascii=False, indent=1), encoding="utf-8")
        return {"hechos": self.hechos, "errores": self.errores, "excel_pendiente": len(pendientes),
                "excel": excel_estado, "excel_mensaje": excel_msg, "excel_escrito": excel_res, "memoria": memoria}

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
        if not f:
            return
        f.destino = nuevo
        if self.cfg.url_base:
            f.url = url_sharepoint(self.cfg.url_base, nuevo.parent.parent.name, nuevo.parent.name, nuevo.name)
            for r in self.lote.filas:
                if r.factura_id == fid:
                    r.url = f.url
        self.log(f"AVISO {viejo.name} ya existía: lo guardo como {nuevo.name}")

    def _mover(self, p: dict) -> None:
        de = Path(p["de"])
        h = sha256(de)
        a = self._destino_libre(p)
        _mover(de, a)
        self.diario.anotar({"op": "mover", "de": str(de), "a": str(a), "sha256": h})
        self.hechos.append(f"{de.name} -> {self._rel(a)}")
        self.log(f"PDF  {de.name}  ->  {self._rel(a)}")

    def _copiar(self, p: dict) -> None:
        de = Path(p["de"])
        a = self._destino_libre(p)
        shutil.copy2(de, a)
        self.diario.anotar({"op": "copiar", "de": str(de), "a": str(a), "sha256": sha256(a)})
        self.hechos.append(f"{de.name} (del zip) -> {self._rel(a)}")
        self.log(f"PDF  {de.name} (zip)  ->  {self._rel(a)}")

    def _mover_simple(self, de: Path, a: Path, op: str, etiqueta: str) -> None:
        h = sha256(de)
        a = ruta_libre(a)
        self._carpeta(a.parent)
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


def reintentar_excel(cfg, carpeta_lote: Path, guardar_cambios_usuario=False, log=print) -> tuple[str, str]:
    """Writes the pending rows of a confirmed batch."""
    from .model import cargar_lote
    lote = cargar_lote(carpeta_lote / "lote_confirmado.json")
    estado_, msg, res = escribir_excel(cfg, lote, Diario(carpeta_lote), carpeta_lote, guardar_cambios_usuario, log)
    _guardar_estado(carpeta_lote, excel=estado_, excel_mensaje=msg, excel_escrito=res,
                    excel_pendiente=0 if estado_ == "escrito" else len(lote.filas) + len(lote.proveedores))
    return estado_, msg


def _a_json(o):
    from .model import a_json
    return a_json(o)


# ---------------------------------------------------------------------- undo

def lotes_confirmados(carpeta_lotes: Path) -> list[Path]:
    if not carpeta_lotes.exists():
        return []
    return sorted((d for d in carpeta_lotes.iterdir()
                   if d.is_dir() and estado(d).get("estado", "").startswith("confirmado")), key=lambda d: d.name)


def ultimo_confirmado(carpeta_lotes: Path) -> Path | None:
    ls = lotes_confirmados(carpeta_lotes)
    return ls[-1] if ls else None


def deshacer(carpeta_lote: Path, log=print, cfg=None) -> list[str]:
    """Walks the journal backwards: first the Excel rows, then the files.
    Returns the problems (things that changed after the batch). If the workbook can't be
    opened, nothing is undone (so rows never point to files that moved back)."""
    with _uno_por_vez(carpeta_lote, YaDeshecho):
        if estado(carpeta_lote).get("estado") == "deshecho":
            raise YaDeshecho(f"el lote {carpeta_lote.name} ya estaba deshecho")
        return _deshacer(carpeta_lote, log, cfg)


def _deshacer(carpeta_lote: Path, log, cfg) -> list[str]:
    from . import excel
    pasos = Diario(carpeta_lote).pasos
    problemas: list[str] = []
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
            problemas.append(f"memoria de proveedores: {ex}")
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
                    problemas_deshacer=problemas)
    return problemas
