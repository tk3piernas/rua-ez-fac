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
    pr = sub.add_parser("paridad", help="test de paridad contra el script viejo")
    pr.add_argument("accion", choices=["preparar", "comparar"])
    pr.add_argument("--ps1", type=Path, help="transcript del PS1 (comparar)")
    pr.add_argument("--lote", type=Path, help="carpeta del lote de la app (por defecto, el último)")
    pc = sub.add_parser("procesar", help="mueve los archivos del lote (etapa 3: sin Excel)")
    pc.add_argument("--base", type=Path, help="carpeta Facturas pendientes")
    pc.add_argument("--real", action="store_true", help="permitir la carpeta real de OneDrive")
    pc.add_argument("--si", action="store_true", help="no preguntar")
    dh = sub.add_parser("deshacer", help="deshace el último lote confirmado")
    dh.add_argument("--si", action="store_true", help="no preguntar")
    sub.add_parser("prueba-mover", help="prueba la etapa 3 en la carpeta de prueba y la deshace")
    sub.add_parser("prueba-excel", help="prueba la etapa 4 sobre una copia de la planilla (cerrada y abierta)")
    ep = sub.add_parser("excel-pendiente", help="escribe en la planilla las filas que quedaron pendientes")
    ep.add_argument("--guardar-cambios", action="store_true", help="si la planilla tiene cambios sin guardar, guardarlos antes")
    sub.add_parser("portapapeles", help="copia al portapapeles las filas del último lote (para pegarlas a mano)")
    sub.add_parser("verificar", help="revisa que la PC tenga todo lo necesario")
    sub.add_parser("revisar", help="auditoría de la planilla contra las carpetas (solo lee)")
    sub.add_parser("empaquetar", help="arma un zip para instalar la app en otra PC")
    u = sub.add_parser("ui", help="abre la app en el navegador")
    u.add_argument("--sin-navegador", action="store_true")
    a = ap.parse_args(argv or (sys.argv[1:] or ["ui"]))
    if a.cmd == "ui":
        from . import ui
        ui.main(abrir_navegador=not a.sin_navegador)
        return 0
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
    if a.cmd == "prueba-mover":
        from . import prueba
        return 0 if prueba.prueba_mover(config.carpeta_app() / "datos" / "prueba_mover.txt") else 1
    if a.cmd == "prueba-excel":
        from . import prueba
        return 0 if prueba.prueba_excel(config.carpeta_app() / "datos" / "prueba_excel.txt") else 1
    if a.cmd == "excel-pendiente":
        from . import ejecutor
        d = ejecutor.ultimo_confirmado(config.carpeta_app() / "datos" / "lotes")
        if not d or ejecutor.estado(d).get("excel") != "pendiente":
            print("No hay filas pendientes de escribir.")
            return 0
        est, msg = ejecutor.reintentar_excel(config.cargar(), d, a.guardar_cambios)
        print("Listo, filas escritas." if est == "escrito" else f"Siguen pendientes: {msg}")
        return 0 if est == "escrito" else 1
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
    if a.cmd == "procesar":
        return _procesar(a)
    if a.cmd == "deshacer":
        return _deshacer(a)
    if a.cmd == "paridad":
        from . import paridad
        if a.accion == "preparar":
            paridad.preparar()
            return 0
        datos = config.carpeta_app() / "datos"
        sin = paridad.comparar(a.ps1 or datos / "ps1_paridad.txt", a.lote or paridad.ultimo_lote(),
                               datos / "paridad_resultado.txt")
        return 1 if sin else 0

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


def _procesar(a) -> int:
    from . import ejecutor, prueba
    cfg = config.cargar(base=a.base)
    if not cfg.base:
        print('No encuentro la carpeta "Facturas pendientes".')
        return 1
    en_prueba = "prueba" in [p.lower() for p in cfg.base.parts]
    if not en_prueba and not a.real:
        print("Por ahora (etapa 3) solo se mueven archivos en la carpeta de prueba. Usá --real para la carpeta real.")
        return 1
    lote, trabajo = prueba.correr_lote(cfg, simular=False)
    print(reporte.texto(lote, cfg))
    if not a.si and input("\n¿Confirmo y muevo los archivos? (s/n) ").strip().lower() != "s":
        print("No se hizo nada.")
        return 0
    res = ejecutor.Ejecutor(cfg, lote, trabajo).ejecutar()
    print(f"\nListo: {len(res['hechos'])} pasos, {len(res['errores'])} errores. Excel: {res['excel']} {res['excel_mensaje']}")
    print(f"Lote: {trabajo}")
    return 1 if res["errores"] else 0


def _deshacer(a) -> int:
    from . import ejecutor
    d = ejecutor.ultimo_confirmado(config.carpeta_app() / "datos" / "lotes")
    if not d:
        print("No hay ningún lote confirmado para deshacer.")
        return 1
    if not a.si and input(f"¿Deshago el lote {d.name}? (s/n) ").strip().lower() != "s":
        return 0
    problemas = ejecutor.deshacer(d, cfg=config.cargar())
    print("Listo." if not problemas else "Listo, con problemas:\n  " + "\n  ".join(problemas))
    return 1 if problemas else 0


if __name__ == "__main__":
    sys.exit(main())
