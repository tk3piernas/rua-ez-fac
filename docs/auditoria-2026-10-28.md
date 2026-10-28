# Auditoría de RUA EZ-Fac 0.19.3 (28/10/2026)

Solo lectura: no se cambió código ni se tocó nada real. Se corrieron las dos tandas de tests (152 públicos, 33 privados) y la prueba con Excel de verdad (`-m excel`, sobre carpetas temporales): todas pasan.

**Estado:** el lote 1 (A1, A2, A4, M1 y el test de la CLI) quedó hecho en la 0.19.4.

Esfuerzo: **chico** = menos de una hora con test; **mediano** = media jornada; **grande** = varios días.

## Lo que está bien

- **Repo público:** no hay CUIT, CBU ni alias reales, ni en los archivos de hoy ni en el historial de Git (revisé cada número de 11 y 22 dígitos que pasó alguna vez por el repo). No se subieron PDF, xlsx, csv, `config.toml`, `datos\` ni `prueba\`. Los nombres de proveedores reales de la revisión anterior ya no están. `.gitignore` tiene las redes de seguridad (`*.xlsx`, `*.log`, PDF y txt en `app\`).
- **Reglas de la planilla:**
  - La columna I nunca se escribe.
  - H se pisa solo con 0 (B, C, no fiscal) o si el IVA real difiere de la fórmula en más de la tolerancia; una Factura A sin neto se frena.
  - CUIT, teléfono, cuenta y CBU van como texto, y un valor con "=" pegado a mano no se vuelve fórmula.
  - `modo_real` arranca apagado.
  - Excel siempre por COM; openpyxl solo en tests.
- **Escritura en Excel:** cada fila se anota en el diario antes de escribirla. Si falla a mitad, revierte sin guardar; deshacer solo toca filas con A y J como las dejó la app.
- **Archivos:** deshacer solo devuelve archivos con el mismo SHA256. Lo que borra la app está acotado a sus propias copias, con la excepción del punto M1.
- **De la revisión anterior** (0.9.0) está resuelto casi todo lo que miré: A1 (planilla real en modo prueba), A5 (IVA de Factura A), D1 a D5 (privacidad), y la tanda privada pasó de 3,5 minutos a 8 segundos. Sigue abierto B3 (ver M4).

## Alta: que no se pierdan datos

### A1. La línea de comandos salta el modo de trabajo
- **Dónde:** `app/appfacturas/cli.py:104-139` (`procesar` y `deshacer`).
- **Qué pasa:**
  - `python -m appfacturas procesar --real` mueve archivos del OneDrive real y escribe en la planilla real sin mirar `modo_real`.
  - Además no guarda el modo del lote (Historial no sabe si fue real o de prueba), y decide "prueba" buscando la palabra en la ruta.
  - `deshacer` deshace el último lote sea del modo que sea.
  - El mensaje dice "etapa 3", que es de cuando no había pantalla.
- **Cambio:** exigir `modo_real` para lo real (igual que la pantalla) o, más simple, sacar `procesar`, `deshacer` y `excel-pendiente` de la CLI: hoy todo eso se hace desde la app. **Chico.**

### A2. Los registros JSON se escriben sin protección contra un corte
- **Dónde:** 13 lugares con `write_text(json.dumps(...))`. Los importantes:
  - `ejecutor.py:62` (diario.json) y `:73` (estado.json);
  - `maestro.py:188` y `:251` (historial de la memoria de proveedores);
  - `tandas.py:25`;
  - `preparar.py:330` y `:367` (ultimo_reset.json).
- **Qué pasa:** si la PC se apaga o la app se cierra justo mientras escribe, el archivo queda cortado. Ese lote no se puede deshacer ni terminar (la app no puede leer su diario).
- **Cambio:** una función `guardar_json` (archivo temporal, `fsync` y `os.replace`, como ya hace `borrado.py`) usada en todos lados. **Chico.**

### A3. Una ventana chica para filas duplicadas al terminar un lote cortado
- **Dónde:** `excel.py:492-493` y `pasos_guardados` (`:599`).
- **Qué pasa:** Excel guarda la planilla y recién después se anota `excel_guardado`. Si la app se corta entre las dos cosas, el diario cree que esas filas no se guardaron. **Terminar este lote** las vuelve a escribir: quedan duplicadas.
- **Cambio:** al continuar, en los pasos sin `excel_guardado`, mirar la celda: si A y J son las que escribió la app, darla por escrita. **Mediano.**

### A4. Mover un PDF se anota después de moverlo
- **Dónde:** `ejecutor.py:229-236` y `_mover_simple` (`:246`).
- **Qué pasa:** si la app se corta entre el movimiento y la anotación, el PDF queda en la carpeta del proveedor sin rastro en el diario. **Deshacer** no lo devuelve, y **Terminar este lote** falla porque el archivo ya no está en `_Entrada`.
- **Cambio:** anotar la intención antes de mover y marcarla hecha después (como `borrado.py`). Al terminar, si el archivo ya está en destino con el mismo SHA256, contarlo como movido. **Chico.**

## Media

### M1. Borrar un pedido a Claude usa `rmtree` dentro del OneDrive
- **Dónde:** `claude_cowork.py:198-202`.
- **Qué pasa:** borra entera una carpeta de `_Revisar con Claude\`, y solo se fija en el nombre de la carpeta padre. Lo que haya adentro (también algo que alguien puso a mano) se pierde sin papelera.
- **Cambio:** mandarla a la papelera (`borrado._a_la_papelera`) o borrar solo los archivos que creó la app. **Chico.**

### M2. Un Excel colgado traba a todos los que vienen después
- **Dónde:** `excel.py`, `LibroCOM.cerrar` (`:269-282`).
- **Qué pasa:**
  - Si Excel queda colgado (por ejemplo con un cartel escondido), `Quit` falla en silencio y queda un EXCEL.EXE oculto. Los Excel que se abren después rechazan las llamadas. Pasó hoy.
  - Si Excel se cuelga sin rechazar (esperando un cartel), la tarea no termina nunca y los botones quedan grises hasta reiniciar la app.
- **Cambio:** guardar el número de proceso del Excel que abrió la app y cerrarlo si `Quit` falla, solo ese. Para el cuelgue sin rechazo, un límite de tiempo que avise y libere los botones. **Mediano.**

### M3. Se pinta la columna I (sigue abierto B3 de la revisión anterior)
- **Dónde:** `excel.py:535-537` (`for c in COLS: h.pintar(...)`, COLS incluye la I).
- **Qué pasa:** un posible duplicado pinta de amarillo A:J, I incluida, y el reset le saca el relleno. No es escribir el valor, pero es tocar el formato de la I.
- **Cambio:** decidir si la I se pinta o no. Si no, excluirla del pintado y del reset. **Chico.**

## Baja: limpieza

### L1. Código que no usa nadie
- **Funciones:** `excel._col` (`:60`), `lote._nombre_sugerido` (`:635`), `pdf.imagen_pagina` (`:89`), `util.fmt_cuit` (`:109`), `ventana.id_de` (`:48`).
- **Imports sin uso:** `re` y `Decimal` en `excel.py`, `solo_digitos` en `parsers/generico.py`.
- **Cambio:** sacarlos. **Chico.**

### L2. Herramientas y archivos del script viejo
- `app/herramientas/paridad.py` compara contra el PS1, que se retiró en la 0.13.1: ya no se puede usar.
- Hay `.pyc` huérfanos en `app/herramientas/__pycache__`.
- Revisar si siguen sirviendo `Simular (CLI).bat` (va con A1), `Probar Excel (prueba).bat` y `Probar mover archivos (prueba).bat`. Los dos `Probar` son locales: no están en el repo.
- **Cambio:** retirar `paridad.py` y lo que no se use. **Chico.**

### L3. La misma secuencia de Excel repetida en 8 lugares
- **Dónde:** "abrir la planilla → frenar si tiene cambios sin guardar → hacer algo → guardar → cerrar", en `excel.py` (2), `maestro.py` (2), `preparar.py` (2) y `borrado.py` (2).
- **Cambio:** un `with excel.libro_listo(cfg, ruta) as libro:` que haga el chequeo y el cierre. Menos lugares donde olvidarse del chequeo. **Mediano** (toca mucho; con las dos tandas de tests).

### L4. Detalles
- `ejecutor.escribir_excel` tiene dos `except` idénticos.
- `excel.escribir` renombra `anotar` dos veces sin necesidad.
- En `ejecutor.py:18-21`, un import quedó entre constantes.
- **Cambio:** limpiar. **Chico.**

### L5. `ui.py` tiene 2.100 líneas
- Siete funciones de 90 a 130 líneas: `panel_revision`, `panel_proveedores`, `panel_config`, `seccion_claude` y otras.
- **Cambio:** pasar Configuración (con la Zona de peligro) a `ui_config.py`. Solo ordena; no cambia nada que se vea. **Grande** y opcional: no lo haría antes de CONTROL AMCA, salvo que AMCA agregue otra pantalla grande.

## Tests

- **T1. Sin tests:** `cli.py` (justo donde está A1), `reporte.py`, `bancos.py` y `carpeta_prueba.py`. `parsers/generico.py` solo se prueba con facturas reales. Agregar al menos uno de la CLI con A1 y uno del reporte. **Chico.**
- **T2. La prueba con Excel de verdad** (`-m excel`) pasa, pero nunca corre sola. Al cerrar Excel tira un "Windows fatal exception" que es ruido de COM. Conviene correrla antes de cada versión que toque `excel.py`. **Chico** (anotarlo en CLAUDE.md).
- **T3.** El reset con Excel de verdad sigue sin probar (está en Pendiente). Lo cubriría un test `-m excel` del reset sobre una copia temporal. **Chico.**

## Documentación

- **Doc1.**
  - `docs/DOCUMENTACION.md:3` dice "Versión 0.18.2".
  - `CLAUDE.md:3` dice "áreas Vehicular y Hogar" (son cuatro) y que se reemplaza al PS1, que ya se retiró.
  - `CLAUDE.md` dice "cualquier .ps1", pero `instalar.ps1` se sube a propósito: aclararlo. **Chico.**
- **Doc2. Los prompts viejos están mezclados con la documentación:**
  - subidos: `prompt-borrar-todo.md` y `prompt-revision-fable.md`;
  - sin subir: `prompt-opus-plan-revision.md` y `revision-2026-10-28.md` (casi todo resuelto).
  - **Cambio:** pasarlos a `docs/historial/` (o sacarlos del repo) y dejar en `docs\` solo lo vigente. **Chico.**

## Orden sugerido para limpiar

1. **0.19.4, datos:** A1, A2, A4, M1 y T1 de la CLI. Todo chico.
2. **0.19.5:** A3, M2 y M3 (con lo que decidas sobre la I).
3. **0.19.6, limpieza:** L1, L2, L4, Doc1, Doc2 y T2/T3.
4. **Opcional:** L3, y L5 si hace falta para AMCA.
