"""Core: builds a batch (lote) from input files. Reads only; the plan is executed elsewhere."""
from __future__ import annotations

import re
import shutil
import tempfile
from collections import defaultdict
from dataclasses import dataclass
from datetime import date
from decimal import ROUND_HALF_UP, Decimal
from pathlib import Path
from urllib.parse import unquote

from . import parsers, pdf
from . import planilla as pl_mod
from . import maestro, txt_proveedor
from .bancos import banco_de
from .carpetas import Indice, clave, clave_nf
from .config import Config
from .model import Archivo, Estado, Factura, Fila, Lote, Nivel
from .planilla import FilaPlanilla, num_de, tareas_de
from .proveedores import Registro, clave_sin_cuit, nombre_prolijo
from .util import ahora_id, cuit_valido, fmt_fecha, fmt_monto, norm, ruta_libre, sha256, url_sharepoint

CENT = Decimal("0.01")


@dataclass
class Entrada:
    ruta: Path
    area: str = ""              # "" = deduce from the provider
    area_origen: str = ""


def entradas_de_carpeta(cfg: Config) -> tuple[list[Entrada], list[tuple[Path, str]]]:
    """Files waiting in _Entrada\\<area>: (pdf/zip entries, txt files)."""
    out, txts = [], []
    for a in cfg.areas:
        d = cfg.entrada / a
        if not d.is_dir():
            continue
        for f in sorted(d.iterdir(), key=lambda p: p.name.lower()):
            if not f.is_file():
                continue
            s = f.suffix.lower()
            if s in (".pdf", ".zip"):
                out.append(Entrada(f, a, "carpeta"))
            elif s == ".txt":
                txts.append((f, a))
    # PS1 order: all PDFs of an area, then its zips
    out.sort(key=lambda e: (cfg.areas.index(e.area), e.ruta.suffix.lower() == ".zip", e.ruta.name.lower()))
    return out, txts


class Armador:
    def __init__(self, cfg: Config, simular: bool = True, leer_planilla: bool = True, hoy: date | None = None,
                 trabajo: Path | None = None):
        """trabajo: folder of this batch (datos\\lotes\\<id>); zips are extracted there and kept for commit/undo."""
        self.cfg = cfg
        self.hoy = hoy or date.today()
        self.lote = Lote(id=trabajo.name if trabajo else ahora_id(), simular=simular, modo_filas=cfg.filas)
        self.reg = Registro(cfg.csv_proveedores if cfg.base else None)
        self.pl = pl_mod.leer(cfg.planilla, cfg.areas) if leer_planilla else pl_mod.Planilla(None, error="no se leyó (opción)")
        self.idx = Indice(cfg.base, cfg.areas, self.reg.cuit_por_carpeta()) if cfg.base else None
        self.en_cola: dict[str, str] = {}
        self.por_clave: dict[str, Factura] = {}
        self.datos_txt: list = []
        self.hash_tanda: dict[str, str] = {}
        self.ocupadas: set[str] = set()
        self.trabajo = trabajo
        self.tmp = (trabajo / "extraidos") if trabajo else Path(tempfile.mkdtemp(prefix="appfact_"))
        self.tmp.mkdir(parents=True, exist_ok=True)
        self.n = 0
        c = self.lote.contexto
        c["proveedores_csv"] = len(self.reg.por_cuit)
        c["guardadas"] = self.idx.total if self.idx else 0
        c["planilla"] = (f"{len(self.pl.filas)} filas cargadas, {len(self.pl.proveedores)} proveedores"
                         if self.pl.leida else f"AVISO: {self.pl.error}; no controlo duplicados contra la planilla")

    # ------------------------------------------------------------ input
    def procesar(self, entradas: list[Entrada], txts: list[tuple[Path, str]] = ()) -> Lote:
        """entradas: PDFs and zips; txts: (path, area) of provider .txt files."""
        self.datos_txt = []
        for ruta, area in txts:
            try:
                self.datos_txt.append((txt_proveedor.leer(ruta), area))
            except Exception as ex:
                self.lote.pendientes.append(f"{area}\\{ruta.name} -> no pude leer el .txt ({ex})")
        for e in entradas:
            if e.ruta.suffix.lower() == ".zip":
                self._zip(e)
            else:
                self._pdf(e.ruta, e.area, e.area_origen, e.ruta.name)
        self._duplicados_planilla()
        self._proveedores_hoja()
        self._filas()
        self._plan()
        if not self.trabajo:
            shutil.rmtree(self.tmp, ignore_errors=True)
        return self.lote

    def _zip(self, e: Entrada) -> None:
        dest = self.tmp / f"zip{len(list(self.tmp.iterdir()))}"
        try:
            pdfs = pdf.expandir_zip(e.ruta, dest)
        except Exception as ex:
            self.lote.pendientes.append(f"{e.area}\\{e.ruta.name} -> no pude abrir el zip ({ex})")
            return
        res = [self._pdf(p, e.area, e.area_origen, f"{e.ruta.name} > {p.name}", zip_=e.ruta) for p in pdfs]
        # like the PS1: the zip moves only if every PDF inside found its place (a repeat of a file
        # already saved counts as placed; unreadable ones and repeats inside the batch do not)
        fallos = [f for f in res if f.estado == Estado.NO_LEIDA or (f.dup and not f.existente)]
        primero = next((f for f in res if f.carpeta and not f.dup), None)
        if primero and not fallos:
            destino = ruta_libre(self.cfg.base / primero.archivo.area / primero.carpeta / e.ruta.name, self.ocupadas)
            self.ocupadas.add(str(destino).lower())
            self.lote.plan.append({"op": "mover_zip", "de": str(e.ruta), "a": str(destino)})
        elif primero:
            self.lote.aviso(Nivel.INFO, "zip_queda", f"{e.ruta.name} queda en _Entrada: algunos PDF de adentro no se pudieron ubicar.")

    def _pdf(self, ruta: Path, area: str, area_origen: str, nombre: str, zip_: Path | None = None) -> Factura:
        self.n += 1
        arch = Archivo(ruta, nombre, area, area_origen, zip_, sha256(ruta), ruta.stat().st_size)
        f = Factura(id=f"f{self.n:02d}", archivo=arch)
        self.lote.facturas.append(f)
        doc = pdf.leer(ruta)
        parser, _ = parsers.elegir(doc, ruta.name) if doc.paginas else (None, 0)
        e = None
        if parser:
            try:
                e = parser.extraer(doc, ruta.name, self.cfg)
            except Exception as ex:   # a broken parser must not stop the batch
                self.lote.aviso(Nivel.REVISAR, "parser_error", f"{nombre}: error leyendo ({ex})", f.id)
        f.ext = e
        if e and e.cuit == self.cfg.cuit_rua:
            e.cuit = ""
        fiscal_ok = bool(e and cuit_valido(e.cuit) and e.pv is not None and e.nro is not None)
        nf_ok = bool(e and not fiscal_ok and (e.razon_social or e.cbu) and e.numero_txt and e.total is not None)
        if not (fiscal_ok or nf_ok):
            igual = self.hash_tanda.get(arch.sha256)
            if not igual and self.idx:
                x = self.idx.buscar(None, arch.tamano, arch.sha256, excluir=ruta)
                igual = self.idx.rel(x) if x else None
            if igual:
                f.estado, f.dup, f.a_duplicados = Estado.DUPLICADA, f"es el mismo archivo que {igual}", True
            else:
                f.estado = Estado.NO_LEIDA
                motivo = ("no tiene texto (¿escaneada?)" if not doc.tiene_texto else
                          "formato no reconocido todavía" if not parser else
                          "no pude leer CUIT / punto de venta / número")
                self.lote.pendientes.append(f"{area or '?'}\\{nombre}  -> {motivo}")
            return f
        self.hash_tanda.setdefault(arch.sha256, f"{area}\\{nombre}")
        prov = self._identificar(f, e, nombre, fiscal_ok)
        if self._duplicado(f, e, nombre):
            return f
        self._ubicar(f, prov, nombre, ruta)

        # ---- amounts (PS1 rules) and checks
        self._importes(f, e)
        self._validar(f, e)
        f.estado = Estado.REVISAR if any(a.ref == f.id and a.nivel != Nivel.INFO for a in self.lote.avisos) else Estado.LISTA
        return f

    def _identificar(self, f: Factura, e, nombre: str, fiscal_ok: bool):
        """Provider (registered or proposed as new), area, duplicate key and file / link names."""
        arch = f.archivo
        area = arch.area
        prov, sugerencia = self._buscar_proveedor(e)
        if not area:
            if prov and prov.area:
                arch.area, arch.area_origen = prov.area, "proveedor"
            else:
                self.lote.aviso(Nivel.BLOQUEA, "sin_area", f"{nombre}: no sé de qué área es (elegila)", f.id)
        area = arch.area
        if area and area in self.pl.faltan:
            self.lote.aviso(Nivel.BLOQUEA, "sin_hoja", f"{nombre}: la planilla no tiene la hoja {area}. "
                            "Agregala desde Configuración → Setup (crea una copia vacía de las otras).", f.id)
        if not prov:
            nombre_p, origen = sugerencia
            clave_p = e.cuit if fiscal_ok else clave_sin_cuit(nombre_p)
            prov = self.reg.proponer(clave_p, area, nombre_p)
            f.proveedor_nuevo = True
            quien = f"CUIT {e.cuit}" if fiscal_ok else "Este proveedor (sin CUIT)"
            self.lote.aviso(Nivel.REVISAR, "proveedor_nuevo",
                            f"{nombre}: {quien} no está en proveedores.csv. Se agregaría como '{prov.nombre}' "
                            f"(carpeta {area}\\{prov.carpeta}; nombre tomado {origen})", f.id)
        elif prov.area and area and prov.area != area:
            self.lote.aviso(Nivel.INFO, "otra_area", f"{nombre}: {prov.nombre} está registrado en {prov.area}", f.id)
        f.proveedor, f.carpeta, f.proveedor_clave = prov.nombre, prov.carpeta, prov.clave
        if not e.vto and (e.emision or e.desde):
            e.vto = e.emision or e.desde
            e.faltantes = [x for x in e.faltantes if x != "vto"]
            f.vto_supuesto = True
        f.clave = clave(e.cuit, e.tipo, e.pv, e.nro) if fiscal_ok else clave_nf(prov.clave, e.tipo, e.numero_txt)
        f.nombre_destino = f"{e.tipo} {e.numero}.pdf"
        f.texto_link = f"{e.tipo} {e.numero}"
        return prov

    def _duplicado(self, f: Factura, e, nombre: str) -> bool:
        """Certain duplicates in this batch, and the same number with another amount in the folders.
        True = this invoice stops here (not loaded)."""
        arch = f.archivo
        if f.clave in self.en_cola:
            otra = self.por_clave[f.clave]
            if otra.ext and e.total is not None and otra.ext.total is not None and abs(otra.ext.total - e.total) > CENT:
                f.estado = Estado.REVISAR
                f.dup = (f"tiene el mismo número que {self.en_cola[f.clave]} pero otro importe "
                         f"({fmt_monto(e.total)} contra {fmt_monto(otra.ext.total)}): revisalo antes de cargarlo")
                self.lote.aviso(Nivel.BLOQUEA, "mismo_numero", f"{nombre}: {f.dup}", f.id)
                return True
            f.estado, f.dup, f.a_duplicados = Estado.DUPLICADA, f"es la misma factura que {self.en_cola[f.clave]}", True
            return True
        existente = self.idx.buscar(f.clave, arch.tamano, arch.sha256, excluir=arch.ruta) if self.idx else None
        if existente and e.total is not None and self.idx.hash_de(existente) != arch.sha256:
            # same number already saved, but another file: if the amount differs it is not the same invoice
            otro_total = self._total_de(existente)
            if otro_total is not None and abs(otro_total - e.total) > CENT:
                f.estado = Estado.REVISAR
                f.dup = (f"tiene el mismo número que {self.idx.rel(existente)} pero otro importe "
                         f"({fmt_monto(e.total)} contra {fmt_monto(otro_total)}): revisalo antes de cargarlo")
                self.lote.aviso(Nivel.BLOQUEA, "mismo_numero", f"{nombre}: {f.dup}", f.id)
                return True
        f.existente = existente
        return False

    def _ubicar(self, f: Factura, prov, nombre: str, ruta: Path) -> None:
        """Where the PDF goes (or the already saved copy it points to) and its link."""
        arch, area, existente = f.archivo, f.archivo.area, f.existente
        if existente:
            f.existente, f.a_duplicados = existente, True
            carpeta, area_link, archivo = existente.parent.name, existente.parent.parent.name, existente.name
            self.lote.aviso(Nivel.INFO, "ya_guardada",
                            f"{nombre} ya estaba guardada como {self.idx.rel(existente)}; no la vuelvo a guardar "
                            f"(la copia va a _Entrada\\_Duplicados)", f.id)
        else:
            carpeta, area_link = prov.carpeta, area
            if self.cfg.base and area:
                destino = ruta_libre(self.cfg.base / area / carpeta / f.nombre_destino, self.ocupadas)
                self.ocupadas.add(str(destino).lower())
                f.destino = destino
                archivo = destino.name
                self.idx.agregar(f.clave, destino, arch.tamano, arch.sha256)
            else:
                archivo = f.nombre_destino
        self.en_cola[f.clave] = f"{area}\\{nombre}"
        self.por_clave[f.clave] = f
        f.url = url_sharepoint(self.cfg.url_base, area_link, carpeta, archivo) if self.cfg.url_base else ""

    def _total_de(self, ruta: Path):
        """Total of an invoice already saved in the folders (None if it can't be read)."""
        try:
            doc = pdf.leer(ruta)
            parser, _ = parsers.elegir(doc, ruta.name) if doc.paginas else (None, 0)
            return parser.extraer(doc, ruta.name, self.cfg).total if parser else None
        except Exception:
            return None

    # ------------------------------------------------------------ amounts
    def _importes(self, f: Factura, e) -> None:
        if e.letra in ("A", "M") and e.neto is not None:
            # G = everything paid without IVA: taxed net + not taxed + exempt (tolls, travel...)
            f.monto = e.neto + (e.no_gravado or 0)
            resto = (e.total - f.monto - (e.otros_tributos or 0)).quantize(CENT) if e.total is not None else None
            if e.iva:
                f.iva = e.iva
            elif resto is not None and resto > 0:
                f.iva = resto                    # IVA not read (or read as 0) but the total says there is some
            elif resto is not None and resto == 0:
                f.iva = Decimal(0)               # total = neto: no IVA (exempt); H must be 0, not the formula
            else:
                f.iva = None                     # nothing to check against: the sheet formula stays
        elif e.total is not None:
            f.monto = e.total
            f.iva = Decimal(0) if e.letra in ("B", "C") or not e.fiscal else None
        if e.es_nc:
            f.monto = -f.monto if f.monto is not None else None
            f.iva = -f.iva if f.iva else f.iva

    # ------------------------------------------------------------ checks
    def _validar(self, f: Factura, e) -> None:
        av = lambda nivel, cod, txt: self.lote.aviso(nivel, cod, f"{f.texto_link} ({f.proveedor}): {txt}", f.id)
        if e.letra in ("A", "M") and e.neto is None:
            # G would get the total and H the formula on top of it: paid 21 % more. Not loaded.
            av(Nivel.BLOQUEA, "a_sin_neto", f"es Factura {e.letra} pero no leí el neto gravado; no la cargo "
                                            "(cargala a mano en la planilla)")
        elif e.letra in ("A", "M") and f.iva == 0:
            av(Nivel.REVISAR, "factura_a_sin_iva", f"es Factura {e.letra} sin IVA (el total es igual al neto): H va en 0")
        elif e.letra == "A" or e.letra == "M":
            av(Nivel.REVISAR, "factura_a", f"es Factura {e.letra}: lleva IVA ({fmt_monto(f.iva)})")
        elif e.letra == "B":
            av(Nivel.REVISAR, "factura_b", "es Factura B: el monto ya incluye el IVA, H va en 0")
        if e.items:
            s = sum((i.monto or Decimal(0) for i in e.items), Decimal(0))
            ref = e.neto + (e.no_gravado or 0) if e.letra in ("A", "M") and e.neto is not None else e.total
            if ref is not None and abs(s - ref) > CENT:
                av(Nivel.REVISAR, "suma_items", f"la suma de los ítems ({fmt_monto(s)}) no coincide con {fmt_monto(ref)}")
        if e.letra in ("A", "M") and None not in (e.neto, e.total):
            calc = e.neto + (e.no_gravado or 0) + (e.iva or 0) + (e.otros_tributos or 0)
            if abs(calc - e.total) > CENT:
                av(Nivel.REVISAR, "total", f"neto + IVA + otros ({fmt_monto(calc)}) no da el total ({fmt_monto(e.total)})")
        if e.otros_tributos:
            av(Nivel.REVISAR, "otros_tributos", f"tiene otros tributos por {fmt_monto(e.otros_tributos)} (no van a la planilla)")
        if e.cuit_receptor and e.cuit_receptor != self.cfg.cuit_rua:
            av(Nivel.REVISAR, "receptor", f"no está hecha a RUA (CUIT receptor {e.cuit_receptor})")
        elif e.condicion_receptor and "inscripto" not in e.condicion_receptor.lower():
            av(Nivel.INFO, "consumidor_final", f"RUA figura como '{e.condicion_receptor}'")
        if f.vto_supuesto:
            av(Nivel.REVISAR, "sin_vto", f"no trae fecha de vencimiento: pongo la de emisión ({fmt_fecha(e.vto)}), como el script viejo")
        elif e.vto:
            dias = (e.vto - self.hoy).days
            if dias < 0:
                av(Nivel.REVISAR, "vencida", f"vencida el {fmt_fecha(e.vto)}")
            elif dias <= self.cfg.dias_vence_pronto:
                av(Nivel.REVISAR, "vence_pronto", f"vence el {fmt_fecha(e.vto)} (en {dias} días)")
        else:
            av(Nivel.REVISAR, "sin_vto", "no encontré la fecha de vencimiento")
        if not e.items:
            av(Nivel.REVISAR, "sin_items", "no pude leer el detalle")
        for i in e.items:
            if _sin_detalle(i.descripcion):
                av(Nivel.REVISAR, "sin_detalle", f"ítem sin detalle del servicio ('{i.descripcion}')")
            if not i.tareas:
                av(Nivel.INFO, "sin_tarea", f"ítem sin N° de tarea ('{i.descripcion[:60]}')")
        if e.fuente != "arca" and e.fiscal:
            nivel = Nivel.REVISAR if e.faltantes else Nivel.INFO
            av(nivel, "generico", "leída con el lector genérico: revisá los datos"
               + (f" (no encontré: {', '.join(e.faltantes)})" if e.faltantes else ""))
        elif not e.fiscal and e.faltantes:
            av(Nivel.REVISAR, "faltantes", f"no encontré: {', '.join(e.faltantes)}")
        if e.tareas_archivo and not any(i.tareas for i in e.items):
            av(Nivel.INFO, "tareas_archivo", f"el nombre del archivo menciona {', '.join(e.tareas_archivo)}; "
               "no las cargo en B (completalas en la revisión si corresponden)")
        for nota in e.notas:
            av(Nivel.REVISAR if re.search(r"no fiscal|vencimiento|no coincide|varios CUIT", nota) else Nivel.INFO, "nota", nota)

    # ------------------------------------------------------------ duplicates vs the sheet (PS1 Revisar-Factura)
    def _duplicados_planilla(self) -> None:
        nombre_cuit = self.reg.nombre_a_cuit(self.pl.proveedores)
        filas: list[FilaPlanilla] = list(self.pl.filas)
        grupos: dict[tuple, dict] = {}

        def agrupar(r: FilaPlanilla):
            k = (r.hoja, norm(r.proveedor), r.j_texto or r.j_link or f"fila{r.fila}")
            g = grupos.setdefault(k, {"monto": Decimal(0), "fecha": r.fecha, "fila": r, "hay_monto": False})
            if r.monto is not None:
                g["monto"] += r.monto
                g["hay_monto"] = True

        for r in filas:
            agrupar(r)
        for f in self.lote.facturas:
            if f.dup or not f.clave:
                continue
            e = f.ext
            num = f"{e.pv}-{e.nro}"
            tareas_f = [t for i in e.items for t in i.tareas]
            posibles = []
            for r in filas:
                rc = nombre_cuit.get(norm(r.proveedor))
                if not rc and r.j_link:
                    m = re.match(r"^(\d{11})_", re.split(r"[\\/]", unquote(r.j_link))[-1])
                    rc = m.group(1) if m else None
                mismo = (rc and rc == e.cuit) or norm(r.proveedor) == norm(f.proveedor)
                lugar = f"{r.hoja} fila {r.fila}" if isinstance(r.fila, int) else f"esta misma tanda ({r.j_texto})"
                if mismo and ((e.nro is not None and (num_de(r.j_texto) == num or num_de(r.j_link) == num))
                              or norm(r.j_texto) == norm(f.texto_link)):
                    f.dup = f"ya está cargada en {lugar}"
                    break
                if r.j_link and f.url and unquote(r.j_link) == unquote(f.url):
                    f.dup = f"su link ya está en {lugar}"
                    break
                for t in tareas_f:
                    if t in tareas_de(r.tarea):
                        posibles.append(f"la tarea {t} ya figura en {lugar} ({r.proveedor})")
            if f.dup:
                f.estado = Estado.DUPLICADA
                # it will not be loaded: its other warnings (due date, Factura A...) no longer matter
                self.lote.avisos = [a for a in self.lote.avisos if a.ref != f.id or a.nivel == Nivel.BLOQUEA]
                continue
            fecha = e.desde or e.emision
            for g in grupos.values():
                r = g["fila"]
                rc = nombre_cuit.get(norm(r.proveedor))
                mismo = (rc and rc == e.cuit) or norm(r.proveedor) == norm(f.proveedor)
                if (mismo and g["hay_monto"] and f.monto is not None and abs(g["monto"] - f.monto) < CENT
                        and fecha and g["fecha"] == fecha):
                    lugar = f"{r.hoja} fila {r.fila}" if isinstance(r.fila, int) else "esta misma tanda"
                    posibles.append(f"mismo proveedor, monto y fecha que {lugar} ({r.j_texto})")
            f.posibles = list(dict.fromkeys(posibles))
            if f.posibles:
                f.estado = Estado.REVISAR
                self.lote.aviso(Nivel.REVISAR, "posible_duplicado",
                                f"{f.texto_link} ({f.proveedor}): posible duplicado: {' | '.join(f.posibles)}", f.id)
            nueva = FilaPlanilla(f.archivo.area, "nueva", f.proveedor, " / ".join(dict.fromkeys(tareas_f)),
                                 fecha, f.monto, f.texto_link, f.url)
            filas.append(nueva)
            agrupar(nueva)


    # ------------------------------------------------------------ providers
    def _buscar_proveedor(self, e):
        """(registered provider, None) or (None, (suggested name, where it came from))."""
        if e.cuit:
            p = self.reg.buscar(e.cuit)
            if p:
                return p, None
        if e.cbu:
            r = next((r for r in self.pl.proveedores if r.get("cbu") == e.cbu), None)
            if r:
                p = (self.reg.buscar(r["cuit"]) if r["cuit"] else None) or self.reg.buscar_nombre(r["nombre"])
                if p:
                    return p, None
        for n in (e.nombre_fantasia, e.razon_social):
            if n:
                p = (self.reg.buscar_nombre(n) or self.reg.buscar_nombre(nombre_prolijo(n, self.cfg.siglas))
                     or self.reg.buscar(clave_sin_cuit(nombre_prolijo(n, self.cfg.siglas))))
                if p and (not e.cuit or not p.cuit or p.cuit == e.cuit):
                    return p, None
        t = next((d for d, _ in self.datos_txt if _txt_coincide(d, e)), None)
        if t and t.nombre:
            return None, (t.nombre, "del .txt")
        hoja = next((r for r in self.pl.proveedores if e.cuit and r["cuit"] == e.cuit), None)
        if hoja and hoja["nombre"]:
            return None, (hoja["nombre"], "de la hoja Proveedores")
        base = e.nombre_fantasia or e.razon_social
        return None, (nombre_prolijo(base, self.cfg.siglas) or e.cuit or "PROVEEDOR", "de la factura")

    def _proveedores_hoja(self) -> None:
        """Rows A:O for the Proveedores sheet: from each .txt, and from new providers' invoices."""
        hoja = self.pl.proveedores
        por_cuit = {r["cuit"]: r for r in hoja if r.get("cuit")}
        por_nombre = {norm(r["nombre"]): r for r in hoja if r.get("nombre")}
        por_cbu = {r["cbu"]: r for r in hoja if r.get("cbu")}
        filas: dict[str, dict] = {}
        validas = [f for f in self.lote.facturas if f.ext and f.proveedor_clave and f.estado != Estado.DUPLICADA]
        try:
            memoria = maestro.leer(self.cfg)
        except Exception as ex:
            memoria = {}
            self.lote.aviso(Nivel.INFO, "memoria", f"No pude leer {maestro.NOMBRE}: {ex}")

        def en_hoja(cuit, nombre):
            return por_cuit.get(cuit) if cuit else por_nombre.get(norm(nombre))

        for d, area in self.datos_txt:
            fact = next((f for f in validas if _txt_coincide(d, f.ext)), None)
            prov = self.reg.buscar(d.cuit) if d.cuit else None
            asociado = ""
            if not prov and fact:
                prov = self.reg.buscar(fact.proveedor_clave)
                asociado = f"por {'el CBU' if d.cuentas and fact.ext.cbu in [c.cbu for c in d.cuentas] else 'el titular' if not d.cuit else 'el CUIT'} de {fact.texto_link}"
            if not prov and d.cuentas and d.cuentas[0].cbu in por_cbu:
                r = por_cbu[d.cuentas[0].cbu]
                prov = self.reg.buscar(r["cuit"]) or self.reg.buscar_nombre(r["nombre"])
            cuit = d.cuit or (fact.ext.cuit if fact else "") or (prov.cuit if prov else "")
            nombre = (prov.nombre if prov else "") or d.nombre or nombre_prolijo(d.titular, self.cfg.siglas) or d.archivo.stem
            area = (prov.area if prov and prov.area else "") or area
            ref = f"{area}\\{d.archivo.name}"
            if not prov and cuit:
                prov = self.reg.proponer(cuit, area, nombre)
                self.lote.aviso(Nivel.REVISAR, "proveedor_nuevo", f"{ref}: CUIT {cuit} nuevo; se agregaría como '{nombre}'")
            elif not prov:
                self.lote.aviso(Nivel.REVISAR, "txt_sin_proveedor",
                                f"{ref}: no sé de qué proveedor es (no trae CUIT y no coincide con ninguna factura)")
            info = {"archivo": str(d.archivo), "nombre": d.archivo.name, "area": area, "proveedor": nombre,
                    "cuit": cuit, "asociado": asociado, "banco": d.cuentas[0].banco if d.cuentas else "",
                    "cbu": d.cuentas[0].cbu if d.cuentas else "", "avisos": list(d.avisos)}
            self.lote.txts.append(info)
            for a in d.avisos:
                self.lote.aviso(Nivel.REVISAR, "txt", f"{ref}: {a}")
            ya = en_hoja(cuit, nombre)
            if ya:
                txt_cbu = d.cuentas[0].cbu if d.cuentas else ""
                self.lote.aviso(Nivel.INFO, "ya_en_hoja", f"{ref}: {nombre} ya está en la hoja Proveedores (fila {ya['fila']}); "
                                "no copio otra fila. Si cambió algún dato, actualizalo ahí.")
                if txt_cbu and ya.get("cbu") and txt_cbu != ya["cbu"]:
                    self.lote.aviso(Nivel.REVISAR, "cbu_distinto", f"{ref}: el CBU del .txt ({txt_cbu}) no es el de la hoja (fila {ya['fila']})")
                continue
            obs = []
            c0 = d.cuentas[0] if d.cuentas else None
            if c0:
                if not c0.banco_escrito:
                    obs.append("Banco deducido del CBU")
                if not c0.valido:
                    obs.append("CBU/CVU con dígito verificador inválido: revisar")
                otro = por_cbu.get(c0.cbu)
                if otro and otro.get("cuit") != cuit:
                    obs.append(f"Mismo CBU que {otro['nombre']} (fila {otro['fila']}): revisar")
                    self.lote.aviso(Nivel.REVISAR, "cbu_ajeno", f"{ref}: el CBU es el mismo que tiene {otro['nombre']} (fila {otro['fila']})")
            else:
                obs.append("Sin CBU/CVU")
            for c in d.cuentas[1:]:
                obs.append(f"Otra cuenta: {c.banco}{' (' + c.tipo + ')' if c.tipo else ''} CBU/CVU {c.cbu}"
                           + (f" alias {c.alias}" if c.alias else ""))
            if not d.cuit and cuit:
                obs.append("CUIT tomado de la factura")
            if not cuit:
                obs.append("Sin CUIT")
            e = fact.ext if fact else None
            if e and e.razon_social and norm(e.razon_social) != norm(nombre):
                obs.append(f"Razón social: {nombre_prolijo(e.razon_social, self.cfg.siglas)}")
            vals = [area, nombre, cuit, d.condicion_iva or (e.condicion_emisor if e else ""),
                    d.domicilio or (e.domicilio if e else ""), d.contacto, d.telefono, d.email,
                    c0.banco if c0 else "", c0.tipo if c0 else "", c0.nro if c0 else "", c0.cbu if c0 else "",
                    c0.alias if c0 else "", (c0.titular if c0 else "") or d.titular or d.contacto, ". ".join(obs)]
            filas[prov.clave if prov else f"txt:{d.archivo.name}"] = {"origen": d.archivo.name, "valores": vals}

        # providers of this batch missing from the Proveedores sheet whose details are in the memory
        for f in validas:
            p = self.reg.buscar(f.proveedor_clave)
            if not p or p.clave in filas or en_hoja(p.cuit, p.nombre):
                continue
            m = maestro.buscar(memoria, p.cuit, p.nombre)
            if m:
                vals = list(m["valores"])
                vals[0] = vals[0] or p.area
                filas[p.clave] = {"origen": maestro.NOMBRE, "valores": vals, "de_memoria": True}

        for p in self.reg.nuevos.values():
            fs = [f for f in validas if f.proveedor_clave == p.clave]
            if not fs or p.clave in filas or en_hoja(p.cuit, p.nombre):
                continue
            e = fs[0].ext
            obs = []
            if e.razon_social and norm(e.razon_social) != norm(p.nombre):
                obs.append(f"Razón social: {nombre_prolijo(e.razon_social, self.cfg.siglas)}")
            if not p.cuit:
                obs.append("Sin CUIT (comprobante no fiscal)")
            if e.cbu:
                obs.append("Banco deducido del CBU")
            vals = [p.area, p.nombre, p.cuit, e.condicion_emisor, e.domicilio, "", "", "",
                    banco_de(e.cbu) if e.cbu else "", "", "", e.cbu, "", "", ". ".join(obs)]
            filas[p.clave] = {"origen": fs[0].texto_link, "valores": vals, "solo_vacios": True}
            if not e.cbu:
                self.lote.aviso(Nivel.INFO, "sin_datos", f"{p.nombre}: no tengo sus datos bancarios ni de contacto; "
                                "pegalos en la pestaña Proveedores y la app los recuerda")
        self.lote.proveedores = [{"clave": k, **v} for k, v in filas.items()]

    # ------------------------------------------------------------ rows
    def _filas(self) -> None:
        cfg = self.cfg
        for f in self.lote.facturas:
            if f.dup or not f.clave or not f.archivo.area:
                continue
            e = f.ext
            hoja = f.archivo.area
            alic = self.pl.alicuota.get(hoja, Decimal("0.21"))
            signo = -1 if e.es_nc else 1
            fecha = e.desde or e.emision
            base = dict(factura_id=f.id, hoja=hoja, proveedor=f.proveedor, fecha=fecha, vto=e.vto, url=f.url,
                        texto_link=f.texto_link, resaltar=bool(f.posibles),
                        comentario=("Posible duplicado:\n- " + "\n- ".join(f.posibles)) if f.posibles else "",
                        empresa=cfg.defaults.get(f"empresa_{hoja.lower()}", ""))
            if cfg.filas == "factura" or len(e.items) <= 1:
                tareas = list(dict.fromkeys(t for i in e.items for t in i.tareas))
                serv = " / ".join(_servicio_limpio(i.descripcion) for i in e.items) or "Servicio (sin detalle en la factura)"
                self.lote.filas.append(Fila(tarea=" / ".join(tareas), servicio=serv, monto=f.monto,
                                            iva=_iva_h(e.letra, f.monto, f.iva, alic, cfg.tolerancia_iva), **base))
            else:
                for i in e.items:
                    monto = (i.monto * signo) if i.monto is not None else None
                    iva = (i.iva * signo) if i.iva is not None else None
                    if e.letra in ("A", "M") and iva is None and f.iva is not None and e.neto and i.monto is not None:
                        iva = (f.iva * (i.monto / e.neto)).quantize(CENT, ROUND_HALF_UP)
                    self.lote.filas.append(Fila(tarea=" / ".join(i.tareas), servicio=_servicio_limpio(i.descripcion),
                                                monto=monto, iva=_iva_h(e.letra, monto, iva, alic, cfg.tolerancia_iva),
                                                **base))
        # an invoice with a blocking warning is shown in the grid but never loaded (one missing sheet
        # used to make the whole batch fail when writing)
        bloqueadas = {a.ref for a in self.lote.avisos if a.nivel == Nivel.BLOQUEA and a.ref}
        for r in self.lote.filas:
            if r.factura_id in bloqueadas:
                r.excluida = r.bloqueada = True

    # ------------------------------------------------------------ plan (what a real run would do)
    def _plan(self) -> None:
        usados = {f.proveedor_clave for f in self.lote.facturas if f.proveedor_clave and f.estado != Estado.DUPLICADA}
        usados |= {t["cuit"] for t in self.lote.txts if t["cuit"]}
        for p in self.reg.nuevos.values():
            if p.clave in usados:
                self.lote.plan.append({"op": "proveedor_nuevo", "cuit": p.cuit, "clave": p.clave, "area": p.area,
                                       "nombre": p.nombre, "carpeta": p.carpeta})
        for r in self.lote.proveedores:
            self.lote.plan.append({"op": "proveedor_hoja", "valores": r["valores"]})
        for t in self.lote.txts:
            destino = ruta_libre(self.cfg.entrada / "_Procesados" / f"{self.hoy:%Y-%m-%d} {t['nombre']}", self.ocupadas)
            self.ocupadas.add(str(destino).lower())
            self.lote.plan.append({"op": "txt_procesado", "de": t["archivo"], "a": str(destino)})
        bloqueadas = {a.ref for a in self.lote.avisos if a.nivel == Nivel.BLOQUEA and a.ref}
        for f in self.lote.facturas:
            if f.id in bloqueadas or f.estado in (Estado.NO_LEIDA, Estado.EXCLUIDA) or (f.dup and not f.a_duplicados and f.estado != Estado.DUPLICADA):
                continue      # stays in _Entrada
            if f.archivo.zip:
                if f.destino and not f.a_duplicados:   # PDFs inside a zip: their extracted copy is saved; the zip moves as a whole
                    self.lote.plan.append({"op": "copiar", "de": str(f.archivo.ruta), "a": str(f.destino), "factura": f.id})
                continue
            if f.a_duplicados:
                self.lote.plan.append({"op": "a_duplicados", "de": str(f.archivo.ruta), "factura": f.id})
            elif f.destino:
                self.lote.plan.append({"op": "mover", "de": str(f.archivo.ruta), "a": str(f.destino), "factura": f.id})
        for fila in self.lote.filas:
            self.lote.plan.append({"op": "fila", "hoja": fila.hoja, "factura": fila.factura_id})


# ------------------------------------------------------------ helpers
def _iva_h(letra: str, monto, iva, alic: Decimal, tol: float):
    """Value for column H, or None to keep the sheet formula."""
    if letra in ("B", "C", ""):
        return Decimal(0)
    if iva is None or monto is None:
        return None
    formula = (monto * alic).quantize(CENT, ROUND_HALF_UP)
    return iva if abs(iva - formula) > Decimal(str(tol)) else None


def _nombre_sugerido(razon: str) -> str:
    return re.sub(r"\s+", " ", razon or "").strip()


_TAREA_TXT = re.compile(r"(?i)\(?\s*(?:n\s*[°º.]*\s*de\s*tarea|tarea|exp(?:ediente)?\.?|servicio)\s*(?:n\s*[°º.*]*\s*)?:?\s*\d{6}\s*\)?")


def _servicio_limpio(desc: str) -> str:
    s = _TAREA_TXT.sub(" ", desc or "")
    s = re.sub(r"(?<!\d)\d{6}(?!\d)", " ", s)
    s = re.sub(r"\(\s*\)|\s+[,/-]\s*$|^\s*[,/-]\s+", " ", s)
    s = re.sub(r"\s+", " ", s).strip(" ,;-/")
    return s or "Servicio (sin detalle en la factura)"


def _sin_detalle(desc: str) -> bool:
    return re.sub(r"(?i)tarea:?|servicio|corresponde|[\d\s:.,/°º#*-]+", "", desc or "") == ""


def _txt_coincide(d, e) -> bool:
    """Is this .txt about the issuer of this invoice? (CUIT, CBU, account holder or name)."""
    if not e:
        return False
    if d.cuit and e.cuit:
        return d.cuit == e.cuit
    if e.cbu and any(c.cbu == e.cbu for c in d.cuentas):
        return True
    nombres = [n for n in (e.razon_social, e.nombre_fantasia) if n]
    tit = {w for w in norm(d.titular or d.contacto).split() if len(w) > 2}
    for n in nombres:
        toks = {w for w in norm(n).split() if len(w) > 2}
        if len(tit) >= 2 and tit <= toks:
            return True
    return bool(d.nombre) and norm(d.nombre) in [norm(n) for n in nombres]


def renombrar_proveedor(lote: Lote, cfg, factura_id: str, nombre: str) -> str | None:
    """Grid edit of "Proveedor". A NEW provider is renamed everywhere: folder, csv line, link and rows of all
    its invoices in the batch. A registered one is not (its name and folder come from proveedores.csv):
    returns why. None = done."""
    from .proveedores import limpiar_nombre
    nombre = re.sub(r"\s+", " ", nombre or "").strip()
    f = next((x for x in lote.facturas if x.id == factura_id), None)
    if not f or not nombre:
        return "Falta el nombre del proveedor"
    # new = it will be added to proveedores.csv by this batch (only its first invoice has proveedor_nuevo)
    if not any(p["op"] == "proveedor_nuevo" and p["clave"] == f.proveedor_clave for p in lote.plan):
        return (f"{f.proveedor} ya está registrado en proveedores.csv: no lo cambio acá (la factura iría a su carpeta "
                "de siempre con otro nombre en la fila). Si el nombre está mal, corregilo en proveedores.csv.")
    carpeta = limpiar_nombre(nombre)
    for g in lote.facturas:
        if g.proveedor_clave != f.proveedor_clave:
            continue
        g.proveedor, g.carpeta = nombre, carpeta
        if g.destino:
            nuevo = g.destino.parent.parent / carpeta / g.destino.name
            for p in lote.plan:
                if p.get("factura") == g.id and p.get("a") == str(g.destino):
                    p["a"] = str(nuevo)
            g.destino = nuevo
            if cfg.url_base:
                g.url = url_sharepoint(cfg.url_base, nuevo.parent.parent.name, carpeta, nuevo.name)
        for r in lote.filas:
            if r.factura_id == g.id:
                r.proveedor, r.url = nombre, g.url
    for p in lote.plan:
        if p["op"] == "proveedor_nuevo" and p["clave"] == f.proveedor_clave:
            p["nombre"], p["carpeta"] = nombre, carpeta
    for p in lote.proveedores:
        if p["clave"] == f.proveedor_clave:
            p["valores"][1] = nombre
    return None


def iva_al_editar_monto(lote: Lote, fila: Fila) -> Decimal | None:
    """Column H after the amount of a row is edited: 0 stays 0 (B, C, non-fiscal); a Factura A goes back to
    the sheet formula (the IVA read from the PDF was for the old amount)."""
    f = next((x for x in lote.facturas if x.id == fila.factura_id), None)
    letra = f.ext.letra if f and f.ext and f.ext.fiscal else ""
    return Decimal(0) if letra in ("B", "C", "") else None


def aplicar_exclusiones(lote: Lote) -> set[str]:
    """After the review: an invoice whose rows were all excluded is not moved nor loaded.
    Returns the excluded invoice ids."""
    filas: dict[str, list[Fila]] = defaultdict(list)
    for r in lote.filas:
        filas[r.factura_id].append(r)
    excluidas = {fid for fid, rs in filas.items() if rs and all(r.excluida for r in rs)}
    for f in lote.facturas:
        if f.id in excluidas:
            f.estado = Estado.EXCLUIDA
    zips = {str(f.archivo.zip) for f in lote.facturas if f.id in excluidas and f.archivo.zip}
    usados = {f.proveedor_clave for f in lote.facturas
              if f.proveedor_clave and f.id not in excluidas and f.estado != Estado.DUPLICADA}
    usados |= {t["cuit"] for t in lote.txts if t.get("cuit")}
    lote.proveedores = [p for p in lote.proveedores if p["clave"] in usados or p["clave"].startswith("txt:")]
    lote.plan = [p for p in lote.plan
                 if p.get("factura") not in excluidas
                 and not (p["op"] == "mover_zip" and p["de"] in zips)
                 and not (p["op"] == "proveedor_nuevo" and p["clave"] not in usados)
                 and p["op"] != "proveedor_hoja"]
    lote.plan += [{"op": "proveedor_hoja", "valores": p["valores"]} for p in lote.proveedores]
    return excluidas
