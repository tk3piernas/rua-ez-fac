# RUA EZ-Fac: documentación técnica

Versión 0.19.6 (alpha, 28/10/2026). La app reemplaza a `OrdenarFacturas.ps1` y a la carga manual con Claude de las facturas de proveedores de RUA Asistencia (áreas Vehicular, Hogar, Persona y ServiciosYa; la lista está en `config.AREAS`).

## Qué hace

1. Lee PDF, zip y .txt de proveedores, ya sea desde `Facturas pendientes\_Entrada\<Área>` o subidos desde la pantalla Cargar. **Quitar lo subido** borra solo esas copias (nunca los originales), avisa cuántas quitó y queda deshabilitado si no hay nada subido.
2. Saca los datos de cada factura: CUIT, tipo, punto de venta y número, fechas, montos, IVA, ítems y N° de tarea.
3. Detecta duplicados contra la planilla, contra las carpetas de cada proveedor y dentro del mismo lote.
4. Arma las filas para `DEUDA A PAGAR ARG.xlsx` y las de la hoja Proveedores para los proveedores nuevos o que no estén en la hoja (con los datos de la memoria de proveedores si los tiene).
5. Muestra todo en Revisión. Ahí se puede editar, excluir y confirmar.
6. Al confirmar:
   - renombra y mueve los PDF a la carpeta del proveedor, dentro del OneDrive sincronizado;
   - registra los proveedores nuevos;
   - escribe en Excel por COM, con la planilla abierta o cerrada.
7. Guarda un diario de cada lote, que permite deshacer el último lote completo (Excel y archivos).

## Dónde vive cada cosa

| Qué | Dónde |
|---|---|
| App, entorno de Python, accesos directos y configuración | `Documentos\<carpeta de la app>\`. **Nunca en OneDrive**: el OneDrive de RUA bloquea .exe, .bat, .ps1 y .py. |
| Facturas | `OneDrive - Rua Asistencia\Facturas pendientes\` (carpeta sincronizada; no se usa Graph ni Azure) |
| Planilla | `OneDrive - Rua Asistencia\DEUDA A PAGAR ARG.xlsx` (una hoja por área y Proveedores) |
| Memoria de proveedores | `OneDrive - Rua Asistencia\PROVEEDORES RUA.xlsx` (hoja Proveedores, columnas A–O como la hoja de la planilla + P Actualizado) |
| Registro de proveedores | `Facturas pendientes\_Script\proveedores.csv` (`CUIT;Area;Nombre;Carpeta`, UTF-8) |
| Lotes, informes, diario, respaldos y log | `<app>\datos\lotes\<id>\`, `<app>\datos\backups\`, `<app>\datos\app.log` |
| Carpeta de prueba | `<app>\prueba\OneDrive - Rua Asistencia\` (copia de lo real y de la planilla) |

Estructura de la carpeta de la app:

```
<app>\
  Instalar o actualizar RUA EZ-Fac.bat   instala, actualiza, crea los íconos y revisa la PC
  Abrir RUA EZ-Fac (con consola).bat     igual que el ícono, pero muestra lo que pasa
  Armar instalador para otra PC.bat      arma "RUA EZ-Fac <versión> para instalar.zip"
  config.toml                            lo crea la app la primera vez (se edita con el Bloc de notas)
  requirements*.txt                      librerías con versión fija
  app\appfacturas\                       el código (paquete Python)
  app\tests\                             pruebas automáticas (tienen facturas reales: no se reparten)
  .venv\  datos\  prueba\                los crea la app
```

## Instalación y actualización

- **Requisitos:**
  - Windows con Excel de escritorio.
  - OneDrive de RUA sincronizado con "Facturas pendientes".
  - Python 3.15. El instalador lo baja con el Python Install Manager (`py install 3.15`) si hace falta.
- **Instalar o actualizar con una línea** (`instalar.ps1`, en la raíz del repo): en PowerShell, `irm https://raw.githubusercontent.com/tk3piernas/rua-ez-fac/main/instalar.ps1 | iex`.
  - Elige la carpeta: `Documentos\RUA EZ-Fac`; si ya existe `Documentos\App Facturas` con el instalador, usa esa; si Documentos está adentro de OneDrive, `%LOCALAPPDATA%\RUA EZ-Fac`. Nunca instala en OneDrive ni encima de una copia de Git.
  - Baja `main.zip` de GitHub, lo descomprime en una carpeta temporal y lo copia con `robocopy /E`: actualiza lo que cambió y no toca `config.toml`, `datos\` ni `.venv\` (no vienen en el zip). Le saca la marca de "bajado de internet" (`Unblock-File`) y corre el `.bat`.
  - Como se ejecuta con `iex` y no como archivo `.ps1`, no lo frena la política de ejecución de scripts.
  - Para probarlo sin tocar nada real: `RUA_EZFAC_DESTINO` (otra carpeta), `RUA_EZFAC_ZIP` (un zip local) y `RUA_EZFAC_SIN_INSTALAR=1`.
- **Qué hace `Instalar o actualizar RUA EZ-Fac.bat`:**
  1. Frena si la carpeta está dentro de OneDrive.
  2. Busca Python 3.15. Si está el lanzador `py` (Python Install Manager), lo instala con `py install 3.15`; si no, abre la página de descarga. Si no lo consigue, avisa y termina (antes, sin `py` se cortaba y con `py` pero sin 3.15 no instalaba nada).
  3. Crea `.venv`.
  4. Instala las librerías:
     - `requirements.txt` y `requirements-ui.txt`;
     - después, `requirements-nodeps.txt` con `--no-deps`. PyMuPDF 2.0 pide onnxruntime y NiceGUI pide uvicorn[standard], que trae pyyaml; ninguno de los dos tiene versión para Python 3.15 y la app no los usa.
     - Sin internet, usa la carpeta `wheels\` si existe.
  5. Escribe `appfacturas.pth` en el venv.
  6. Crea los íconos "RUA EZ-Fac" en el Escritorio y en el menú Inicio (`python -m appfacturas accesos`: `pythonw -m appfacturas ui`, ícono `static\ezfac.ico`, AppUserModelID `RUA.EZFac`) y borra los viejos "App Facturas". Si eso falla, los crea como antes con PowerShell (sin la identidad).
  7. Corre las pruebas, si están, y `appfacturas verificar` (resultado en `datos\verificacion.txt`).
- **Actualizar:** se copian los `.py` nuevos en `app\appfacturas\` y se abre la app. Si estaba abierta una versión anterior, la app la cierra sola: consulta `/api/version` y, si no coincide, hace `taskkill` del python que escucha en el puerto. Solo hace falta reinstalar cuando cambian los `requirements`.
- **Instalar en otra PC:**
  1. En la PC que ya tiene la app, correr `Armar instalador para otra PC.bat`. El zip no lleva pruebas, datos, configuración ni facturas.
  2. En la otra PC, descomprimirlo en Documentos y abrir el instalador.
  3. El link de SharePoint de la columna J sale de la cuenta de OneDrive de esa PC: registro `HKCU\Software\Microsoft\OneDrive\Accounts\BusinessN`, `ServiceEndpointUri` o, si no está, `UserEmail` más la dirección de SharePoint de la empresa (`[rutas] tenant` en `config.toml`; no está escrita en el código). Si en `config.toml` se deja `url_base` vacío, se arma solo.

## Cómo arranca

- El ícono abre un servidor local (NiceGUI) en `127.0.0.1:8765`. No es accesible desde otra PC.
- Cuando el servidor está listo, abre la app en una ventana propia de Edge, o de Chrome si no está Edge, con `--app=`: sin pestañas ni barra de direcciones. Si no encuentra ninguno de los dos, usa el navegador predeterminado.
- Con `pythonw` no hay consola: lo que pasa queda en `datos\app.log`.
- **Recorrido para empezar** (`recorrido.py` + `static/recorrido.js`): la primera vez que se abre la app, un recorrido de 22 pasos oscurece la pantalla y resalta cada parte (modo de trabajo, "N para leer", planilla, Salir, cada pestaña y las secciones de Configuración) con una tarjeta que explica qué hace, y Anterior / Siguiente / Saltar (también Esc y las flechas). Cambia solo de pestaña. Si no hay un lote leído, Revisión y Resultado muestran durante el recorrido un **lote de ejemplo** armado en memoria (`recorrido.lote_demo`: cuatro facturas inventadas, con una fila amarilla, una roja y una violeta, avisos y un resultado), así esos pasos se ven como de verdad. No apunta a ningún archivo: Confirmar, Deshacer y Terminar lote se niegan mientras está, y desaparece al terminar o saltar el recorrido (o al volver a abrir la ventana). Si hay un lote leído de verdad, el recorrido usa ese. Lo que igual no existe ("N para leer" sin nada esperando) se explica con la tarjeta en el medio. Se anota en `datos\recorrido.json` y no vuelve a salir solo; el botón **?** del encabezado lo repite. Los pasos apuntan a elementos marcados con `data-tour` en `ui.py` (un test controla que existan). El script lleva la fecha del archivo en la dirección, así el navegador no usa uno viejo después de actualizar.
- Arriba, al lado del nombre de la app:
  - **N para leer**: aparece cuando hay archivos esperando en `_Entrada` o subidos en la app (se actualiza cada 5 segundos), y agrega "· M nuevos" si llegaron después de la última lectura (se recuerda en `datos\ultima_lectura.json`). Un clic lee desde cualquier pestaña y abre Revisión.
  - **Abrir planilla**: abre en Excel la planilla en la que escribe la app (la real o la copia de prueba, según el modo).
  - **Exportar planilla**: guarda una copia con el nombre que elijas (por defecto `DEUDA A PAGAR ARG dd-mm-aaaa`) en la carpeta `Planillas exportadas`, al lado de la planilla. Usa `SaveCopyAs` de Excel, así que si está abierta incluye lo que tenga sin guardar; la planilla no cambia. Nunca pisa una exportada con el mismo nombre. Después de exportar se puede abrir la copia o deshacer (la borra si nadie la tocó). Para empezar de cero después, Reset.
- Deshacer un lote que ya estaba deshecho no hace nada (antes, un segundo clic informaba como "cambió después" todo lo que ya había vuelto).
- **Barra de tareas:** la ventana es Edge en modo app, pero apenas se abre la app le pone la identidad `RUA.EZFac` (AppUserModelID), con su ícono, su nombre y el comando para volver a abrirla (`ventana.py`). Los accesos directos llevan la misma identidad. Así la ventana se agrupa con el ícono de RUA EZ-Fac y no con Edge, y *Anclar a la barra de tareas* (clic derecho en el ícono de la barra, o en el acceso directo) deja el ícono de la app y abre la app.
- **Una sola ventana:** con la app abierta, tocar el ícono trae al frente la ventana que ya está, en vez de abrir otra. Varios clics rápidos mientras arranca no levantan varias copias: el primero arranca la app (con un candado de Windows, `Local\RUA.EZFac.inicio.<puerto>`, uno por puerto) y los demás esperan y muestran su ventana. Configuración → *Ventana* → "Permitir abrir más de una ventana de la app" (`[app] varias_ventanas`, apagado por defecto) vuelve a abrir una ventana por clic.
- El botón Salir, en la punta derecha del encabezado (al lado del modo de trabajo), cierra la ventana de la app (si el navegador no la deja cerrarse, muestra "RUA EZ-Fac se cerró") y apaga el servidor; lo fuerza a los 1,5 segundos si algo lo retiene.

## Pantallas

1. **Cargar:**
   - Área de lo que subís: Vehicular, Hogar, Persona, ServiciosYa o detectar por proveedor.
   - Subir archivos: se copian a `datos\entrada_ui`; los originales quedan donde estaban. Un .txt necesita el área elegida (no se detecta por proveedor): si no, no se sube y lo avisa.
   - Cuánto hay para leer, y la lista de lo subido con una X para quitar cada archivo. Si quedaron archivos subidos en otra sesión, lo avisa en naranja (se suman al próximo lote).
   - Botón Leer facturas: no mueve ni escribe nada.
2. **Revisión:**
   - Resumen del lote y avisos.
   - Grilla editable: Cargar, Proveedor, fechas, N° de tarea, Servicio y Monto. Colores: amarillo = posible duplicado, rojo = frenada, violeta = leída por Claude.
     - Las columnas tienen ancho fijo y no se achican: los encabezados siempre se leen (y muestran el nombre al pasar el mouse); si la ventana es angosta, la grilla se desplaza hacia el costado. Proveedor y Servicio usan el espacio que sobra y muestran el texto completo al pasar el mouse.
     - Editar **Proveedor** de un proveedor nuevo lo cambia en todo: carpeta, línea de `proveedores.csv`, links y filas de todas sus facturas del lote. Uno ya registrado no se cambia acá (avisa que se corrige en `proveedores.csv`).
     - Editar **Monto**: en Factura B, C o no fiscal H sigue en 0; en Factura A, H vuelve a la fórmula (el IVA leído era del monto anterior) y lo avisa.
   - Secciones de lo que no se carga, los proveedores nuevos y los archivos que se mueven.
   - **Pasarle a Claude lo que no pude leer** (modo B, Cowork): si hay PDF no leídos (escaneados o con un formato desconocido), *Preparar el pedido para Claude* copia los elegidos a `Facturas pendientes\_Revisar con Claude\<lote>\` con `pedido.md` (instrucciones y formato) y `pedido.json` (archivos y SHA256). Los originales quedan en `_Entrada`. Nunca van los .txt; un PDF con un CBU impreso queda destildado, y el pedido le pide a Claude no copiar datos bancarios ni personales. En Cowork se abre esa carpeta y se le pide "Leé pedido.md y hacé lo que dice"; Claude escribe `respuesta.json`. Al leer de nuevo, cada PDF no leído con respuesta (se busca por SHA256) se toma de ahí: fuente "claude", fila violeta y aviso REVISAR; sigue el camino normal (duplicados, grilla, confirmar, deshacer). Una respuesta con `no_pude` se ignora. Si Claude la leyó pero no se puede cargar (por ejemplo un CUIT con el dígito verificador mal, como el de la factura de ejemplo de ARCA), queda en `_Entrada` con un aviso que dice por qué, y no se vuelve a ofrecer para mandar. La lista "Pedidos a Claude" muestra cuántos se respondieron, con Abrir la carpeta y Borrar (solo tiene copias; va a la papelera, porque está en el OneDrive). En Solo simular el botón está deshabilitado.
   - **Leer ahora con la API de Claude** (modo A, `claude_api.py`): con una clave guardada (Configuración → *Claude para lo que no se puede leer*), manda los PDF no leídos elegidos a la API (modelo `claude-opus-5-5`, effort `medium`, respuesta en JSON con esquema fijo, `fallbacks: "default"` por si una política la rechaza) y pide confirmación antes, porque cada factura cuesta unos centavos de dólar. La respuesta se guarda como un pedido más (`_Revisar con Claude\<lote> API\` con `pedido.json` y `respuesta.json`) y el lote se vuelve a leer solo: las facturas leídas quedan violeta con aviso REVISAR, igual que en el modo B. Un error de una factura (conexión, límite, rechazo) no frena las demás y se muestra. El esquema no tiene campos para datos bancarios, nunca van los .txt y los PDF con CBU impreso quedan destildados.
   - **Clave de la API:** se guarda en el Administrador de credenciales de Windows (`keyring`, servicio "RUA EZ-Fac"), nunca en un archivo; también sirve la variable de entorno `ANTHROPIC_API_KEY`. Se puede borrar desde Configuración.
   - Botones Confirmar y Copiar al portapapeles. Si no hay avisos para revisar ni que frenen y nada excluido, Confirmar carga directo (sin el "¿Confirmo?"); si hay, pregunta.
3. **Resultado:** lo que se hizo. Acciones: reintentar la planilla, copiar al portapapeles, deshacer.
4. **Historial:**
   - **Tanda de pagos:** dice en qué tanda está la planilla activa y cuántos lotes se cargaron desde el último cierre. **Cerrar la tanda N…** guarda una copia como `DEUDA A PAGAR ARG tanda N` (o el nombre que elijas) en `Planillas exportadas`, anota en `datos	andas.json` qué lotes entraron y, si tildás "dejar la planilla vacía" y escribís BORRAR, hace el reset con un solo respaldo en `datosackups` (la copia exportada ya queda en OneDrive). **Deshacer el cierre** vuelve la planilla a como estaba (si nadie la tocó y está cerrada en Excel) y borra la copia si nadie la cambió. Cada planilla (real o copia de prueba) lleva su propia numeración.
   - Los lotes anteriores, con la columna **Tanda**, la fecha legible (28/10/2026 10:41), la columna **Dónde** (PRUEBA o REAL) y el estado. Un lote que quedó "en curso" porque se cerró la app a mitad aparece como **quedó a medias**. Para el último lote: **Terminar este lote** (si quedó a medias o con cosas pendientes) y **Deshacer**, que dice sobre qué planilla actúa.
5. **Revisar** (auditoría, solo lee; también `python -m appfacturas revisar`): compara la planilla activa con las carpetas y lista, con un OK o la cantidad:
   - links de J que apuntan a un PDF que ya no está;
   - filas repetidas (mismo link, o mismo proveedor con el mismo texto en J);
   - filas sin link;
   - PDF guardados en las carpetas de los proveedores que no están en ninguna fila (puede ser algo ya pagado o algo que falta cargar);
   - proveedores de la planilla sin datos de pago (ni en la hoja Proveedores ni en la memoria);
   - filas con H escrito como número (Factura B o C, o IVA real distinto);
   - archivos que esperan hace más de 7 días en `_Entrada`.
   Botón **Copiar informe** para pegarlo en un mail o chat.
5. **Proveedores** (`maestro.py`): memoria de datos de proveedores en `PROVEEDORES RUA.xlsx`.
   - Cuadro para pegar texto (mail, WhatsApp, .txt): lo lee el mismo lector de los .txt y arma una ficha editable con chequeos (dígito verificador de CUIT y CBU, banco del CBU). Si el proveedor ya está, muestra qué cambia.
   - Guardar escribe por Excel (COM), una fila por proveedor (clave CUIT, o el nombre si no tiene CUIT); los campos vacíos nunca borran. CUIT, teléfono, N° de cuenta y CBU van como texto. Necesita el modo "Cargar en la planilla real" o "Practicar en la carpeta de prueba".
   - Al leer un lote, los proveedores que no están en la hoja Proveedores de la planilla toman sus datos de la memoria. Al confirmar, los datos de los .txt (y lo que se sacó de las facturas, solo para completar vacíos) se guardan en la memoria; deshacer el lote los vuelve atrás.
   - Historial de cambios en `datos\memoria_proveedores.json` (los de cada lote, en la carpeta del lote). El botón para deshacer el último cambio mira los dos historiales y dice de dónde vino el cambio (pegado a mano o un lote). Trabaja sobre el archivo que quedó anotado, aunque la app esté ahora en el otro modo (prueba o real). Solo vuelve atrás filas que nadie tocó después; si la app había creado el archivo y queda vacío, lo borra.
6. **Configuración:**
   - **Setup** (ver abajo): crea lo que falte y muestra el chequeo.
   - **Modo de trabajo**, uno solo de tres (también se ve arriba a la derecha):
     - **Solo simular** (por defecto): lee las facturas reales y muestra qué haría; no mueve nada ni escribe en ninguna planilla. Quedan deshabilitados (y frenados también por dentro) Confirmar, cerrar o deshacer una tanda, volver del reset, y deshacer, terminar o reintentar un lote real. Un lote de la carpeta de prueba se puede deshacer siempre. Exportar planilla sí funciona: guarda una copia y la planilla no cambia.
     - **Practicar en la carpeta de prueba**: trabaja de verdad sobre `<app>\prueba\` (copia de la planilla y de las carpetas). Nunca toca lo real, aunque la planilla real esté abierta. Botón para rearmar la carpeta de prueba.
     - **Cargar en la planilla real**: mueve en el OneDrive y escribe en la planilla real. Pide confirmación al elegirlo.
     - Se guarda en `[app] modo_real` (true solo en el modo real). La carpeta de prueba se elige por sesión.
   - **Empezar la planilla de cero** (reset, ver abajo).
   - Una fila por ítem o una por factura; escribir el servicio en la columna C.
- Al abrir, si hay otra versión de la app escuchando en el puerto, la cierra; solo si la línea de comando del proceso es Python con `appfacturas` (otro programa en el mismo puerto no se toca).
- Diseño: colores de RUA (azul #14147a y naranja #ff542c), cada parte en una tarjeta blanca con línea de color arriba. En el encabezado, al lado del nombre, un cartel ALPHA con la foto de Simba ("¡Esto es una alpha! Mandá bugs y sugerencias a Matheo" al pasar el mouse).
- Las tareas largas (rearmar la carpeta de prueba, guardar la memoria, deshacer, exportar, reintentar, Setup, volver del reset) muestran un aviso con ruedita mientras corren, y los botones de acción quedan deshabilitados hasta que terminan.
- Cambiar de modo de trabajo con un lote leído sin confirmar pide confirmación (se descarta lo editado en la grilla). La pestaña Proveedores se actualiza sola con el modo.

## Setup (preparar.py)

- Botón **Configuración → Setup → Revisar y preparar todo**.
- Busca el OneDrive de RUA (`OneDriveCommercial` o `%USERPROFILE%\OneDrive - *`, prefiriendo el que dice Rua) y muestra **qué va a crear** antes de hacerlo. Nunca pisa nada que exista.
  - En la app: `datos\`, `datos\lotes`, `datos\backups`, `datos\entrada_ui`.
  - En el OneDrive: `Facturas pendientes` con `_Entrada\<área>` para cada área, `_Entrada\_Duplicados`, `_Entrada\_Procesados`, una carpeta por área, `_Script` y `_Script\proveedores.csv` (solo el encabezado `CUIT;Area;Proveedor;Carpeta`).
  - Si no hay planilla, la puede crear **vacía**, pero solo si tildás aparte "Sí, crear una planilla nueva vacía" (si ya existe en la nube y OneDrive no la bajó todavía, una vacía genera una copia en conflicto). No lo ofrece si la carpeta del OneDrive está vacía o tiene archivos que todavía están solo en la nube. La planilla vacía tiene la misma estructura: una hoja por área (B4 alícuota, B5 `=TODAY()`, resumen H4:I9, encabezados en la fila 11, filas 12–210 con las fórmulas de H e I, formato de fechas y montos, colores de vencido / vence pronto, lista desplegable de Empresa desde la hoja Listas, una columna por área), Proveedores (encabezados A–O en la fila 4) y Listas.
  - Si la planilla existe pero le faltan hojas de área (por ejemplo una de antes de Persona y ServiciosYa), las agrega como copia vacía de la última hoja de área: mismo formato, fórmulas y resumen, título y lista de Empresa propios. Antes guarda dos respaldos (en la carpeta `Respaldos`, al lado de la planilla, y en `datos\backups`). En la planilla real necesita el modo "Cargar en la planilla real". Mientras falte la hoja, las facturas de esa área quedan frenadas con un aviso.
- Después muestra el **chequeo** (lo mismo que `appfacturas verificar`): librerías, carpetas, `proveedores.csv`, planilla y hojas, link de SharePoint, Excel y Edge, con OK / FALTA.
- Si no encuentra el OneDrive de RUA sincronizado, no crea nada en la nube y lo avisa.

## Reset de la planilla (preparar.reset_planilla)

- Botón **Configuración → Empezar la planilla de cero → Reset de la planilla…**. Para confirmar hay que escribir BORRAR.
- Trabaja sobre la planilla activa: la copia de prueba, o la real solo en el modo "Cargar en la planilla real".
- No crea un archivo nuevo: vacía **el mismo** archivo, así los links compartidos de OneDrive siguen andando.
  - Las hojas de cada área, filas 12–210: borra valores de A–G y J, links, comentarios y rellenos.
  - Vuelve a poner la fórmula de H, tomada de una fila que la conserve (o `=IF(G="","",ROUND(G*$B$4,2))`).
  - La columna I no se toca. El encabezado, el resumen, el formato y las listas quedan.
  - Opcional: vaciar también Proveedores A5:O300.
- Antes guarda dos respaldos: `Respaldos\DEUDA A PAGAR ARG - respaldo antes del reset <fecha>.xlsx` al lado de la planilla, y otra copia en `datos\backups`.
- **Volver al respaldo del reset:** mientras nadie haya cambiado la planilla desde el reset (se compara el SHA256 guardado en `datos\ultimo_reset.json`), aparece este botón al lado del Reset. Copia el respaldo encima del mismo archivo (los links compartidos siguen andando). La planilla tiene que estar cerrada en Excel. Se puede hacer una sola vez.
- Si la planilla está abierta con cambios sin guardar, no hace nada.
- Después de un reset, deshacer un lote anterior ya no borra filas de la planilla (no las encuentra), pero sí devuelve los archivos.

## Borrar todo (borrado.py)

Deja la app como recién empezada. **Nunca borra: mueve** a carpetas `_Borrado <fecha>` y se puede deshacer.

- **Pantalla:** Configuración → **Zona de peligro** (sección roja con ☢, al final).
  - En "Solo simular" el botón está deshabilitado. En "Practicar" actúa solo sobre la copia de prueba.
  - Primera confirmación: qué se va, con cantidades (PDF y carpetas de proveedores, `_Entrada`, filas de la planilla, cada carpeta o archivo), qué queda, el tamaño del zip y el aviso de que el OneDrive se sincroniza con otras personas.
  - Segunda confirmación: escribir BORRAR TODO exacto; el botón se habilita recién a los 5 segundos.
  - Después aparecen **Deshacer el borrado total** y **Vaciar definitivamente…**, con la fecha, el estado, las carpetas `_Borrado` y el zip.
  - Todo corre con `con_espera` y `de_a_uno`, y los botones usan `mientras_libre`. Las funciones de `borrado.py` vuelven a revisar el modo.
- `APPFACTURAS_DOCUMENTOS` cambia la carpeta del zip. Es solo para una instalación de juguete: no cambies `USERPROFILE`, porque Excel arranca con un perfil vacío y se cuelga en la pantalla de primer inicio.

- **Lista cerrada.** Nunca se recorre la raíz del OneDrive para decidir qué se lleva. Solo se llevan estos nombres fijos:
  - en la raíz del OneDrive: `Facturas pendientes\` entero, `PROVEEDORES RUA.xlsx` y los `DEUDA A PAGAR ARG - respaldo *.xlsx`;
  - de `Planillas exportadas\`, los `.xlsx`; de `Respaldos\`, los `DEUDA A PAGAR ARG - respaldo *.xlsx`. Si en esas carpetas hay otra cosa, esa queda y se avisa;
  - en la app (solo en "Cargar en la planilla real"): lo de `datos\` salvo `recorrido.json`, `app.log` y `borrado.json`, y `prueba\`.
- **Nunca toca** `config.toml`, `.venv\`, el código, la clave de Claude ni nada más del OneDrive.
- **Controles de cada ruta, antes del primer movimiento.** Si una falla, no mueve nada:
  - está dentro de su carpeta;
  - no es un link ni una junction, ni tiene uno adentro;
  - no es nada de lo protegido;
  - la planilla y `Facturas pendientes` están juntas en una carpeta `OneDrive - …`.
- **Practicar** trabaja solo sobre `prueba\OneDrive - Rua Asistencia` (si la raíz es otra, frena) y no toca `datos\`. **Cargar en la planilla real** necesita `modo_real`.
- **Pasos:**
  1. Chequeo previo:
     - que no haya otro borrado sin deshacer;
     - que la planilla no tenga cambios sin guardar;
     - que ningún archivo esté abierto en otro programa (se abre sin compartir y sin bajarlo de la nube);
     - que haya lugar para el zip.
  2. Zip de respaldo en `%USERPROFILE%\Documents` (nunca en el OneDrive ni en la app), con `LISTA.json` y el SHA256 de cada archivo, verificado.
  3. Mueve cada cosa con `os.rename`: mismo disco y nunca encima de algo. Anota cada movimiento en `datos\borrado.json` (`borrado_prueba.json` en Practicar) antes y después de hacerlo.
  4. Guarda una copia de la planilla en `_Borrado` y la vacía en el mismo archivo, con el mismo código que el reset (`preparar.vaciar_libro`): hojas de área, filas 12–210, sin tocar la I; y Proveedores A5:O300.
  5. Setup de nuevo: carpetas vacías y `proveedores.csv` con el encabezado.
  - Si algo falla en los pasos 3 o 4, devuelve todo lo que ya movió y la planilla desde la copia.
- **Deshacer** (`borrado.deshacer`) lee solo el registro:
  - Devuelve cada cosa a su lugar. Si una carpeta ya existe (porque la rearmó el Setup), mete adentro lo que vuelve.
  - Un archivo que apareció después del borrado no se pisa: el viejo queda en `_Borrado` y se avisa (estado "deshecho a medias", se puede reintentar).
  - La planilla vuelve de la copia solo si nadie la cambió desde el borrado (SHA256) y está cerrada en Excel.
  - El `proveedores.csv` vacío del Setup se saca si sigue igual.
- **Vaciar definitivamente** (`borrado.vaciar`): hay que escribir VACIAR. Manda a la papelera de Windows (`SHFileOperation` con deshacer) solo las carpetas `_Borrado …` del registro que están justo en la raíz del OneDrive o de la app. Lo del OneDrive va además a la papelera de OneDrive. Nunca usa `rmtree`.
  - Si alguna ruta tiene 260 caracteres o más, no vacía nada y avisa. La papelera no acepta esas rutas, y Windows ofrecería borrarlas para siempre.
  - Si igual aparece ese cartel de Windows, contestá No: la app avisa "No pude vaciar" y todo sigue en `_Borrado`.

## Reglas de negocio

- **Lectura de PDF:**
  - PyMuPDF; los PDF de ARCA (RC4) se abren con contraseña vacía.
  - Se lee solo la copia ORIGINAL. Las marcas DUPLICADO y TRIPLICADO se buscan arriba y al pie, también con letras separadas.
- **Parsers** (`parsers\`, se registran solos):
  - `arca.py` lee por posición: el valor está a la derecha de su etiqueta.
  - `generico.py` arma líneas por posición. Cubre facturas fiscales hechas con otros sistemas de facturación ("Comprobante N°:", "Nº: 0001-00000001", "Neto:") y los comprobantes no fiscales, sin CUIT ni CAE, con número propio ("FAC0001").
- **Nombre de un proveedor nuevo:** si viene todo en mayúsculas se pasa a "Tipo Título", salvo SRL, SA, SAS, las siglas sin vocales ("MG", "STM") y las de `[reglas] siglas` en `config.toml`.
- **N° de tarea:** 6 cifras entre 300000 y 999999. Se toma solo del detalle de la factura; el nombre del archivo es apenas una pista. No se toma un número precedido por "$" o por precio, importe, total, monto o unitario (salvo que diga tarea, exp. o servicio en el medio).
- **CUIT del emisor (lector genérico):** entre los CUIT válidos que no son de RUA, se prefiere el que está en una línea con la etiqueta "CUIT" en el tercio de arriba de la página. Si hay más de uno, aviso REVISAR "hay varios CUIT".
- **Tipos de comprobante:** Factura, Nota de Crédito, Nota de Débito, Recibo (códigos 4, 9, 15 y 54) y Nota de Venta (5 y 10). El tipo va en el nombre del archivo y en la clave de duplicado, así un Recibo no choca con la Factura del mismo número.
- **ARCA sin "Período Facturado"** (facturas de productos): el corte entre emisor y receptor se toma justo arriba del "CUIT:" del receptor.
- **Claves de duplicado:**
  - Factura fiscal: `CUIT|tipo|PV|nro`.
  - No fiscal: `SC:<nombre normalizado>|tipo|FAC…`.
  - Mismo número con otro importe: se frena (BLOQUEA) y queda en `_Entrada`, tanto dentro del lote como contra un PDF ya guardado en la carpeta del proveedor.
- **Avisos BLOQUEA:** la factura se ve en la grilla en rojo, con "Cargar" destildado y sin poder tildarlo. No se mueve ni genera filas, así que el resto del lote se carga igual (por ejemplo, si a la planilla le falta la hoja de un área).
- **Columna H (IVA):**
  - Factura B, Factura C y no fiscal: 0.
  - Factura A: se escribe el IVA real solo si difiere de la fórmula en más de `tolerancia_iva` ($1).
  - Factura A sin IVA (el total es igual al neto, por ejemplo un servicio exento): 0, con aviso REVISAR.
  - Factura A sin neto leído: BLOQUEA ("no leí el neto gravado; no la cargo"). Se carga a mano.
  - Si no, queda la fórmula.
  - La columna I nunca se toca.
- **Factura A:** en G va el neto, no el total. El script viejo cargaba el total y contaba el IVA dos veces.
  - "Importe Neto No Gravado" e "Importe Exento" (peajes, viáticos) se suman a G sin IVA, y H lleva el IVA real (difiere de la fórmula, así que se pisa). Un ítem "Exento" o "No Gravado" va con H en 0.
- **Fechas:**
  - E: fecha del servicio ("Desde"); si no hay, la emisión.
  - F: vencimiento; si no hay, la emisión, con aviso. Nunca el vencimiento del CAE.
  - Se leen dd/mm/aaaa (también con - o .), aaaa-mm-dd y fechas con el mes en letras.
- **CUIT, teléfono, número de cuenta y CBU:** siempre como texto. Un CUIT cargado a mano como número en la hoja Proveedores se reconoce igual; un CBU cargado como número perdió dígitos en Excel y se ignora.
- **Fórmulas:** solo H y J se escriben como fórmula. Cualquier otro valor que empiece con "=" (un alias, una observación) va como texto.
- **.txt de proveedores:**
  - Varias cuentas: la primera va a la fila y las demás a Observaciones.
  - El banco se saca del CBU (384 = Ualá Bank, ex Wilobank).
  - El .txt se asocia a la factura por CUIT, CBU o titular.

## Escritura en Excel (excel.py)

- **Si la planilla está abierta:** la busca en la Running Object Table y escribe en esa misma ventana. Solo toma el libro cuya ruta local es exactamente la de la planilla, o cuya URL de SharePoint completa es la de la planilla real de `config.toml` (también lo verifica con `FullName`). En la carpeta de prueba no se usa ninguna URL, así que nunca escribe en la planilla real aunque esté abierta, ni en una exportada con el mismo nombre.
- **Si algo falla a mitad con la planilla abierta:** revierte las celdas y deja el libro como guardado, para que el reintento no pida guardar cambios que no hizo nadie.
- **Si está cerrada:** abre un Excel oculto, guarda y cierra.
- **Excel colgado:** la app guarda el número de proceso del Excel oculto que abre. Si al cerrar no responde, termina ese proceso (solo si es EXCEL.EXE), así no queda en segundo plano trabando a los siguientes. Nunca toca un Excel abierto por el usuario.
- **Si Excel no responde** ("Call was rejected by callee", porque está arrancando o tiene un cartel escondido): reintenta abrir la planilla durante 20 segundos. Si sigue sin responder, no toca nada y pide cerrar los "Microsoft Excel" del Administrador de tareas y abrir Excel una vez a mano.
- **Fila libre:** la primera fila entre la 12 y la 210 con A–G y J vacías.
- **Columna J:** `=HYPERLINK("url","Factura 0002-00000189")`.
- **Posible duplicado:** A:J en amarillo (color 13431551) y un comentario en A. La I se pinta a propósito: solo cambia el relleno, nunca la fórmula ni el valor.
- **Si la planilla tiene cambios sin guardar:** no escribe y las filas quedan pendientes. Se pueden reintentar o copiar al portapapeles.
- **Antes de escribir:** hace un respaldo en `datos\backups\` (se guardan los últimos 20). Cada paso se anota en el diario antes de escribir; si algo falla a mitad, revierte y no guarda.

## Confirmar y deshacer, de a uno

- Mientras corre una confirmación, un deshacer o un reintento, esos botones quedan deshabilitados.
- `Ejecutor.ejecutar` no confirma un lote que ya tiene estado (en curso, confirmado o deshecho), y ni confirmar ni deshacer corren dos veces a la vez sobre el mismo lote. Antes, un doble clic duplicaba las filas.

## Lotes cortados a medias y archivos que no se movieron

- `lote_confirmado.json` se escribe **antes** de mover nada (el plan ya es final), y otra vez al terminar.
- El diario, los estados de los lotes, las tandas, el historial de la memoria de proveedores y el registro del reset se guardan con `util.guardar_json`. Escribe un archivo temporal, lo baja al disco y recién después lo cambia por el anterior, así que un corte (app cerrada, PC apagada) nunca deja un registro a medio escribir.
- Si la app se cortó mientras escribía en Excel (el diario tiene filas sin `excel_guardado` ni `excel_revertido`), Terminar este lote y Deshacer miran la planilla antes de seguir (`excel.reconocer_cortada`). Excel guarda todo o nada: si alguna fila sigue como la escribió la app, esa escritura se había guardado; Terminar no la repite y Deshacer la saca. Si la planilla tiene cambios sin guardar, espera a que la guardes o la cierres.
- Cada movimiento de un PDF se anota en el diario **antes** de hacerlo (`moviendo`) y otra vez después (`mover`). Si la app se cortó entre las dos anotaciones y el archivo está en destino con el mismo SHA256, Terminar este lote y Deshacer lo dan por movido (`Diario.completar_cortados`): Deshacer lo devuelve, y Terminar no lo busca en `_Entrada`.
- Si un archivo no se pudo mover (por ejemplo OneDrive lo tenía tomado), las filas de esa factura **no** se escriben, para que nunca haya un link a un archivo que sigue en `_Entrada`. En Resultado aparece cuántas quedaron y el botón **Reintentar archivos**.
- **Terminar este lote / Reintentar** (`ejecutor.continuar`): hace solo los pasos de archivos que no están en el diario y después escribe solo las filas que todavía no se guardaron en la planilla. Se puede repetir sin duplicar nada. Sirve también para un lote que quedó a medias.
- `estado.json` guarda `modo` (prueba / real), `base` y `planilla`. Para los lotes viejos, Historial lo deduce de la ruta de la planilla en el diario.

## Deshacer (ejecutor.py)

- Deshace el último lote confirmado (o el que quedó a medias), recorriendo el diario al revés: primero Excel, después los archivos.
- Si la parte de la memoria de proveedores no se pudo deshacer (por ejemplo `PROVEEDORES RUA.xlsx` abierta con cambios sin guardar), el lote queda deshecho pero con esa parte pendiente: tocar Deshacer de nuevo reintenta solo eso.
- Solo toca:
  - filas que siguen como las dejó la app (compara A y J);
  - archivos con el mismo SHA256 que tenían al moverlos.
- Borra las carpetas que creó si quedaron vacías y saca las líneas que agregó a `proveedores.csv`.

## Comandos (python -m appfacturas …)

`ui` (por defecto) · `verificar` · `revisar` · `accesos` · `simular [--entrada]` · `portapapeles` · `empaquetar`

Ninguno mueve archivos ni escribe en la planilla. `procesar`, `deshacer` y `excel-pendiente` se sacaron en la 0.19.4: no miraban el modo de trabajo y podían tocar lo real en "Solo simular". Todo eso se hace desde la app.

Herramienta aparte (fuera del paquete, en `app\herramientas\`): `prueba.py mover|excel` (pruebas automáticas sobre la carpeta de prueba). La herramienta de paridad contra el script viejo y el `paridad.csv` de cada lote se retiraron en la 0.19.6.

## Pruebas

- `app\tests` (en el repo): pruebas que no necesitan facturas reales: montos, fechas, CUIT/CBU (con números inventados), links, empaquetado, Setup y reset de la planilla, búsqueda del libro abierto, doble confirmación, lotes cortados a medias, Recibo y Nota de Venta, tarea vs precio, CUIT del emisor, No Gravado, ARCA sin período, Factura B, Nota de Débito, IVA 10,5 %, PDF escaneado y con contraseña, original y duplicado en un PDF, zip con subcarpetas, .txt con dos CBU, mismo número contra las carpetas, archivos que no se movieron, memoria y reset que se deshacen, edición de Proveedor y Monto, deshacer con A o J cambiadas, reintentar sin la hoja, zip excluido a medias, link de SharePoint con el registro simulado, y facturas sintéticas (Factura A sin neto o exenta, hoja faltante, vencimiento del CAE). `test_repo_sin_datos.py` falla si algún archivo del repo tiene un CBU o un CUIT válido que no sea el de RUA ni uno de los inventados.
- `app\tests_privados` (solo en la PC de desarrollo, en `.gitignore`): pruebas con facturas reales como fixtures (`app\tests_privados\fixtures\`, tienen CUIT y CBU de proveedores). Cubren ARCA A/C/NC, otros sistemas de facturación, no fiscales, duplicados, columna H, .txt, mover + deshacer, Excel y exclusiones.
- Se corren por separado (cada carpeta tiene su `conftest.py`): `python -m pytest -q app\tests` y `python -m pytest -q app\tests_privados`. Las dos usan openpyxl y tardan segundos. La prueba que abre Excel de verdad (COM) corre solo si se pide: `python -m pytest -q -m excel app\tests_privados` (un minuto y medio).
- También solo en esa PC: `Probar mover archivos (prueba).bat`, `Probar Excel (prueba).bat` (Excel real, planilla cerrada y abierta; deshacer la dejó idéntica celda por celda) (los dos llaman a `app\herramientas\prueba.py`). La paridad contra el script viejo dio 0 diferencias sin explicar; el script y sus .bat se retiraron en la 0.13.1.

## Desarrollo con Claude Code y Git

- La carpeta de la app en la PC de desarrollo es a la vez una copia de trabajo de Git del repo `tk3piernas/rua-ez-fac`. Lo que es local (entorno, datos, configuración, facturas de prueba, el script viejo) está en `.gitignore`.
- `CLAUDE.md` (en la raíz) tiene las reglas para Claude Code: qué nunca se sube, cómo probar, cómo versionar y que después de cada cambio hace commit y `git push` a `main`.
- Un compañero que solo usa la app no necesita Git: baja el ZIP del repo o usa `Armar instalador para otra PC.bat`.

## Pendiente

- Probar el reset con Excel real (COM) en la PC, primero sobre la carpeta de prueba.

- Primera corrida real: Matheo elige "Cargar en la planilla real".
