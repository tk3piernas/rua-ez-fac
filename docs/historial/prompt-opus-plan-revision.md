# Prompt: implementar el informe de revisión del 28/10/2026 (para Claude Opus)

Pegalo en una sesión nueva de Claude Code abierta en la carpeta de la app (`Documents\App Facturas`), con el modelo Opus.

---

Leé `CLAUDE.md` y respetalo en todo (español rioplatense, nunca la palabra "envergadura", repo público, qué no se sube, cómo probar, versión, commit y push). Después leé `docs/revision-2026-10-28.md`: es un informe de revisión con hallazgos numerados (A1…A12, B2…B4, C1…C9, D1…D5, E1…E6, F1…F5, G1…G5), cada uno con archivo y línea, caso concreto, cambio propuesto y esfuerzo. Tu trabajo es implementarlo entero, en el orden de abajo, en varias versiones y commits separados.

## Reglas de esta sesión

- Antes de tocar nada, verificá cada hallazgo contra el código actual: los números de línea pueden haber corrido. Si un hallazgo no se reproduce o el cambio propuesto es peor que otro, decímelo en una línea y hacé lo mejor; no lo saltees en silencio.
- Nunca escribas en la planilla real, en `OneDrive - Rua Asistencia` ni en `datos\` para probar. Todo se prueba con los tests (`app\tests` y `app\tests_privados`, cada tanda con su comando) y, si hace falta Excel de verdad, con la carpeta de prueba (`prueba\`) y el modo "carpeta de prueba" de la app.
- Cada hallazgo corregido lleva su test (o la ampliación de uno existente). Los tests públicos de `app\tests` no pueden tener CUIT, CBU, alias, nombres de proveedores reales ni rutas personales: usá números inventados con dígito verificador válido y nombres inventados. Los que necesiten facturas reales van en `app\tests_privados`.
- No cambies el comportamiento documentado que no esté en el informe. Si un cambio toca algo que ve el usuario, actualizá `docs/DOCUMENTACION.md` y `README.md` (y decime si la guía PDF necesita cambios; no la regeneres sin avisar).
- Después de cada versión: correr las dos tandas, subir la versión en `app/appfacturas/__init__.py`, `git status` y `git diff --cached --stat` antes de commitear, commit con mensaje claro en español, `git push origin main`, y decime qué cambió y la ruta completa en Windows de cada archivo tocado.
- Al terminar cada versión frená y esperá mi OK antes de seguir con la siguiente.

## Orden de trabajo

### Versión 0.9.1: que no se pierdan datos (todos chicos)

1. **A1** `excel.py`: pasar `url_onedrive` solo cuando la ruta es la planilla real de `config.cargar()`; después de encontrar el libro en la ROT verificar `wb.FullName` contra la ruta local o la URL completa esperada. Test unitario de `_buscar_abierto` con nombre igual y URL distinta (simulando la ROT).
2. **A6** `excel.py`: tras `_revertir` con libro ya abierto que estaba guardado, dejar `wb.Saved = True`.
3. **A2** `ui.py` y `ejecutor.py`: flag de ocupado que deshabilite Confirmar, Deshacer y Reintentar mientras corre algo; `Ejecutor.ejecutar` se niega si `estado.json` ya dice "en curso" o "confirmado". Test del rechazo.
4. **A4** `lote.py`: las facturas con aviso BLOQUEA no generan filas (o quedan excluidas en `aplicar_exclusiones`) y en la grilla aparecen con "Cargar" destildado y no editable. Test con planilla sin la hoja Persona.
5. **A5** `lote.py`: Factura A/M sin neto → aviso BLOQUEA y no se carga; IVA 0 leído se distingue de IVA no leído (H = 0). Dos tests con PDF sintético (como `_pdf_santiago` en `test_generico.py`).
6. **C1** `parsers/generico.py`: el vencimiento ignora líneas con "CAE". **C2** `util.py`: fechas ISO `aaaa-mm-dd` antes que `dd-mm-aa`. **C8** `lote.py`: ítem sin monto en Factura A no rompe la lectura. Tests de los tres.
7. **D1** `app/tests/test_util.py`: reemplazar el nombre real por uno inventado. **D4** `.gitignore`: agregar `*.xlsx`, `*.log`, `/*.csv`, `/*.pdf`, `app/**/*.pdf`, `app/**/*.txt`, `.claude/`; y un test público que recorra los archivos trackeados (`git ls-files`) y falle si encuentra un CBU válido de 22 dígitos o un CUIT válido distinto del de RUA.
8. **F4** `app/tests_privados/test_ejecutor.py`: el fixture usa `excel_motor = "openpyxl"`; un test aparte con marca `excel` que corre solo si hay Excel. En `LibroCOM.cerrar` soltar las hojas antes de `Quit`.

### Versión 0.9.2: deshacer y lotes a medias (medianos)

9. **A3** `ejecutor.py`: escribir `lote_confirmado.json` antes de ejecutar los pasos; los lotes "en curso" aparecen en Historial como "quedó a medias" y se pueden deshacer y reintentar.
10. **A7** `ejecutor.py`: si falla el paso de archivos de una factura, sus filas no se escriben; quedan pendientes en Resultado con "Reintentar archivos" (o, como mínimo, un aviso claro).
11. **A8** `maestro.py`, `ejecutor.py`, `ui.py`: el deshacer de la memoria arma el cfg desde la ruta guardada en el historial; un lote no se marca "deshecho" si una parte tiró excepción (se guarda qué quedó pendiente); un solo historial de memoria, o que la pestaña Proveedores muestre los dos.
12. **A11** `estado.json` guarda `modo`, `base` y `planilla`; Historial muestra una columna "Dónde" y el botón de deshacer dice sobre qué planilla actúa.
13. **A9, A10, A12** `preparar.py`, `excel.py`: confirmación aparte para crear una planilla nueva (aviso de OneDrive sin sincronizar); botón "Volver al respaldo del reset" si la planilla no cambió desde entonces; conservar las últimas 20 copias en `datos\backups` y poner las del OneDrive en una subcarpeta `Respaldos`.

### Versión 0.9.3: lectura y reglas de la planilla

14. **C5** `parsers/arca.py`, `lote.py`: leer "Importe Neto No Gravado" e "Importe Exento", sumarlos a G sin IVA y pisar H con el IVA real. Test sintético.
15. **C3** `parsers/generico.py`: CUIT emisor preferido por cercanía a la etiqueta "CUIT" en el tercio superior; aviso REVISAR si hay más de un candidato.
16. **C4** `parsers/base.py`: no tomar como tarea números precedidos por "$" o por "precio/importe/total".
17. **C6** `parsers/base.py`, `carpetas.py`: tipos "Recibo" y "Nota de Venta".
18. **B2** `excel.py`: solo J y H pueden ser fórmula; el resto con formato texto. **B4** `planilla.py`: leer CUIT y CBU cargados como número con `maestro._texto`. **B3** (opcional, preguntame antes): sacar la columna I del pintado y del reset.
19. **C7 y C9**: agregar los tests que faltan (Factura B real a partir de la imagen de ejemplo si alcanza, Nota de Débito, FCE MiPyME, PDF escaneado, PDF con contraseña, IVA 10,5 %, multipágina, zip con subcarpetas, .txt con dos CBU, mismo número con otro importe contra las carpetas). Los que necesiten facturas reales van en `tests_privados`.

### Versión 0.10.0: interfaz

20. **E1** panel Proveedores refrescable desde los switches de prueba y escritura real. **E2** confirmación antes de tirar un lote leído al cambiar de modo. **E3** helper `con_espera` con spinner y botón deshabilitado para rearmar, guardar memoria, deshacer, exportar, reintentar y Setup. **E4** lista de archivos subidos con quitar individual, aviso de archivos de otra sesión y área obligatoria para los .txt. **E5** editar Proveedor propaga a carpeta, csv y link para proveedores nuevos (y avisa para registrados); editar Monto recalcula H. **E6** textos del diálogo de deshacer, cartel ALPHA al header, saltar el "¿Confirmo?" cuando no hay avisos REVISAR ni BLOQUEA, fecha legible en Historial.
21. **D2, D3, D5**: descripciones genéricas en `generico.py`, `proveedores.py` y la documentación; `TENANT` como valor por defecto en `config.toml` y `%USERPROFILE%` en los docs; `_cerrar_viejo` solo mata procesos cuya línea de comando contenga `appfacturas`.
22. **F1, F2, F3, F5**: partir `_pdf` en tres funciones y pasar `vto_supuesto` a `Factura`; mover `YaDeshecho` debajo de los imports; un solo `_rel` y una sola función de respaldo; `instalacion.chequeo` usa `preparar.SUBCARPETAS`; `generico.py` usa `cfg.cuit_rua`; tests de deshacer con A/J cambiadas, reintentar con hoja faltante, exclusión parcial de un zip, `url_desde_onedrive` con registro simulado.

### Después, de a una y con mi OK para cada una

23. **G1** pestaña Revisar (auditoría solo lectura). 24. **G2** cierre de tanda. 25. **G5** badge "N para leer" en el header. 26. **G4** retirar el PS1 (`paridad.py` y `prueba.py` a `app\herramientas\`, borrar los .bat de paridad). 27. **G3** modos con Claude: primero el B (carpeta `_Revisar con Claude\<lote>\`), después el A (API con la clave en el Administrador de credenciales de Windows; revisá la documentación actual del SDK `anthropic` antes de escribir código; nunca mandar CBU ni datos bancarios).

Empezá por la 0.9.1. Antes de escribir código, decime en diez líneas cómo vas a resolver A1 y A5, que son los dos que más me preocupan.
