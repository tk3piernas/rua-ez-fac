"""Command line: python -m appfacturas simular [--entrada] [archivos...]"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

from . import __version__, config, reporte
from .lote import Armador, Entrada, entradas_de_carpeta


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(prog="appfacturas", description="RUA EZ-Fac")
    sub = ap.add_subparsers(dest="cmd", required=True)
    s = sub.add_parser("simular", help="lee las facturas y muestra qué haría, sin mover ni escribir nada")
    s.add_argument("archivos", nargs="*", type=Path, help="PDF o zip (si no hay, usa _Entrada)")
    s.add_argument("--entrada", action="store_true", help="procesar las carpetas _Entrada\\<área>")
    s.add_argument("--area", choices=list(config.AREAS), help="área para los archivos sueltos")
    s.add_argument("--filas", choices=["item", "factura"], help="una fila por ítem o por factura")
    s.add_argument("--base", type=Path, help="carpeta Facturas pendientes (para pruebas)")
    s.add_argument("--planilla", type=Path, help="planilla a leer (para pruebas)")
    s.add_argument("--sin-planilla", action="store_true", help="no leer la planilla")
    s.add_argument("--config", type=Path)
    s.add_argument("--salida", type=Path, help="también guarda el informe en este archivo")
    sub.add_parser("portapapeles", help="copia al portapapeles las filas del último lote (para pegarlas a mano)")
    sub.add_parser("verificar", help="revisa que la PC tenga todo lo necesario")
    sub.add_parser("revisar", help="auditoría de la planilla contra las carpetas (solo lee)")
    sub.add_parser("accesos", help="crea los íconos del Escritorio y del menú Inicio (anclables a la barra de tareas)")
    sub.add_parser("empaquetar", help="arma un zip para instalar la app en otra PC")
    u = sub.add_parser("ui", help="abre la app en el navegador")
    u.add_argument("--sin-navegador", action="store_true")
    a = ap.parse_args(argv or (sys.argv[1:] or ["ui"]))
    if a.cmd == "ui":
        from . import ui
        ui.main(abrir_navegador=not a.sin_navegador)
        return 0
    if a.cmd == "accesos":
        from . import ventana
        return 0 if ventana.crear_accesos(config.carpeta_app()) else 1
    if a.cmd == "revisar":
        from . import auditoria
        print(auditoria.informe(auditoria.revisar(config.cargar())))
        return 0
    if a.cmd in ("verificar", "empaquetar"):
        from . import instalacion
        if a.cmd == "empaquetar":
            instalacion.empaquetar()
            return 0
        return 0 if instalacion.verificar(config.carpeta_app() / "datos" / "verificacion.txt") else 1
    if a.cmd == "portapapeles":
        from . import ejecutor, excel
        from .model import cargar_lote
        d = ejecutor.ultimo_confirmado(config.carpeta_app() / "datos" / "lotes")
        if not d:
            print("No hay lotes confirmados.")
            return 1
        excel.copiar_portapapeles(cargar_lote(d / "lote_confirmado.json"))
        print("Filas copiadas. En Excel: clic en la columna A de la primera fila vacía y Ctrl+V.")
        return 0

    cfg = config.cargar(a.config, base=a.base)
    if a.planilla:
        cfg.planilla = a.planilla
    if a.filas:
        cfg.filas = a.filas
    if not cfg.base:
        print('No encuentro la carpeta "Facturas pendientes" del OneDrive de RUA. Revisá config.toml.')
        return 1
    txts = []
    if a.archivos and not a.entrada:
        ents = [Entrada(p, a.area or "", "elegida" if a.area else "") for p in a.archivos if p.suffix.lower() != ".txt"]
        txts = [(p, a.area or "") for p in a.archivos if p.suffix.lower() == ".txt"]
    else:
        ents, txts = entradas_de_carpeta(cfg)
    lote = Armador(cfg, simular=True, leer_planilla=not a.sin_planilla).procesar(ents, txts)
    inf = f"RUA EZ-Fac {__version__}\n" + reporte.texto(lote, cfg)
    print(inf)
    d = reporte.guardar(lote, config.carpeta_app() / "datos" / "lotes", inf)
    print(f"\nInforme guardado en: {d}")
    if a.salida:
        a.salida.write_text(inf, encoding="utf-8")
    return 0


if __name__ == "__main__":
    sys.exit(main())
