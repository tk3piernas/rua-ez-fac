# RUA EZ-Fac: documentación técnica

Versión 0.6.0 (27/10/2026). La app reemplaza a `OrdenarFacturas.ps1` y a la carga manual con Claude de las facturas de proveedores de RUA Asistencia (áreas Vehicular y Hogar).

## Qué hace

1. Lee PDF, zip y .txt de proveedores, ya sea desde `Facturas pendientes\_Entrada\<Área>` o subidos desde la pantalla Cargar.
2. Saca los datos de cada factura: CUIT, tipo, punto de venta y número, fechas, montos, IVA, ítems y N° de tarea.
3. Detecta duplicados contra la planilla, contra las carpetas de cada proveedor y dentro del mismo lote.
4. Arma las filas para `DEUDA A PAGAR ARG.xlsx` y las de la hoja Proveedores para los proveedores nuevos.
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
| Planilla | `OneDrive - Rua Asistencia\DEUDA A PAGAR ARG.xlsx` (hojas Vehicular, Hogar y Proveedores) |
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
- **Qué hace `Instalar o actualizar RUA EZ-Fac.bat`:**
  1. Frena si la carpeta está dentro de OneDrive.
  2. Crea `.venv`.
  3. Instala las librerías:
     - `requirements.txt` y `requirements-ui.txt`;
     - después, `requirements-nodeps.txt` con `--no-deps`. PyMuPDF 2.0 pide onnxruntime y NiceGUI pide uvicorn[standard], que trae pyyaml; ninguno de los dos tiene versión para Python 3.15 y la app no los usa.
     - Sin internet, usa la carpeta `wheels\` si existe.
  4. Escribe `appfacturas.pth` en el venv.
  5. Crea los íconos "RUA EZ-Fac" en el Escritorio y en el menú Inicio (`pythonw -m appfacturas ui`, ícono `static\ezfac.ico`) y borra los viejos "App Facturas".
  6. Corre las pruebas, si están, y `appfacturas verificar` (resultado en `datos\verificacion.txt`).
- **Actualizar:** se copian los `.py` nuevos en `app\appfacturas\` y se abre la app. Si estaba abierta una versión anterior, la app la cierra sola: consulta `/api/version` y, si no coincide, hace `taskkill` del python que escucha en el puerto. Solo hace falta reinstalar cuando cambian los `requirements`.
- **Instalar en otra PC:**
  1. En la PC que ya tiene la app, correr `Armar instalador para otra PC.bat`. El zip no lleva pruebas, datos, configuración ni facturas.
  2. En la otra PC, descomprimirlo en Documentos y abrir el instalador.
  3. El link de SharePoint de la columna J sale de la cuenta de OneDrive de esa PC: registro `HKCU\Software\Microsoft\OneDrive\Accounts\BusinessN`, `ServiceEndpointUri` o `UserEmail`. Si en `config.toml` se deja `url_base` vacío, se arma solo.

## Cómo arranca

- El ícono abre un servidor local (NiceGUI) en `127.0.0.1:8765`. No es accesible desde otra PC.
- Cuando el servidor está listo, abre la app en una ventana propia de Edge, o de Chrome si no está Edge, con `--app=`: sin pestañas ni barra de direcciones. Si no encuentra ninguno de los dos, usa el navegador predeterminado.
- Con `pythonw` no hay consola: lo que pasa queda en `datos\app.log`.
- El botón Salir apaga el servidor, y lo fuerza a los 1,5 segundos si algo lo retiene.

## Pantallas

1. **Cargar:**
   - Área de lo que subís: Vehicular, Hogar o detectar por proveedor.
   - Subir archivos: se copian a `datos\entrada_ui`; los originales quedan donde estaban.
   - Cuánto hay para leer.
   - Botón Leer facturas: no mueve ni escribe nada.
2. **Revisión:**
   - Resumen del lote y avisos.
   - Grilla editable: Cargar, Proveedor, fechas, N° de tarea, Servicio y Monto. Colores: amarillo = posible duplicado, rojo = frenada, violeta = leída por Claude.
   - Secciones de lo que no se carga, los proveedores nuevos y los archivos que se mueven.
   - Botones Confirmar y Copiar al portapapeles.
3. **Resultado:** lo que se hizo. Acciones: reintentar la planilla, copiar al portapapeles, deshacer.
4. **Historial:** los lotes anteriores y la opción de deshacer el último confirmado.
5. **Configuración:**
   - Carpeta de prueba sí o no, y rearmarla.
   - **Habilitar escritura real** (`[app] modo_real`, apagado por defecto). Mientras está apagado, lo real solo se simula.
   - Una fila por ítem o una por factura; escribir el servicio en la columna C.

## Reglas de negocio

- **Lectura de PDF:**
  - PyMuPDF; los PDF de ARCA (RC4) se abren con contraseña vacía.
  - Se lee solo la copia ORIGINAL. Las marcas DUPLICADO y TRIPLICADO se buscan arriba y al pie, también con letras separadas.
- **Parsers** (`parsers\`, se registran solos):
  - `arca.py` lee por posición: el valor está a la derecha de su etiqueta.
  - `generico.py` arma líneas por posición. Cubre MDQ, el estilo Santiago SOS y los comprobantes no fiscales, como los de ABC con número "FAC0297".
- **N° de tarea:** 6 cifras entre 300000 y 999999. Se toma solo del detalle de la factura; el nombre del archivo es apenas una pista.
- **Claves de duplicado:**
  - Factura fiscal: `CUIT|tipo|PV|nro`.
  - No fiscal: `SC:<nombre normalizado>|tipo|FAC…`.
  - Mismo número con otro importe: se frena (BLOQUEA) y queda en `_Entrada`.
- **Columna H (IVA):**
  - Factura B, Factura C y no fiscal: 0.
  - Factura A: se escribe el IVA real solo si difiere de la fórmula en más de `tolerancia_iva` ($1).
  - Si no, queda la fórmula.
  - La columna I nunca se toca.
- **Factura A:** en G va el neto, no el total. El script viejo cargaba el total y contaba el IVA dos veces.
- **Fechas:**
  - E: fecha del servicio ("Desde"); si no hay, la emisión.
  - F: vencimiento; si no hay, la emisión, con aviso.
- **CUIT, teléfono, número de cuenta y CBU:** siempre como texto.
- **.txt de proveedores:**
  - Varias cuentas: la primera va a la fila y las demás a Observaciones.
  - El banco se saca del CBU (384 = Ualá Bank, ex Wilobank).
  - El .txt se asocia a la factura por CUIT, CBU o titular.

## Escritura en Excel (excel.py)

- **Si la planilla está abierta:** la busca en la Running Object Table, por ruta o por su URL de SharePoint, y escribe en esa misma ventana.
- **Si está cerrada:** abre un Excel oculto, guarda y cierra.
- **Fila libre:** la primera fila entre la 12 y la 210 con A–G y J vacías.
- **Columna J:** `=HYPERLINK("url","Factura 0002-00000189")`.
- **Posible duplicado:** A:J en amarillo (color 13431551) y un comentario en A.
- **Si la planilla tiene cambios sin guardar:** no escribe y las filas quedan pendientes. Se pueden reintentar o copiar al portapapeles.
- **Antes de escribir:** hace un respaldo en `datos\backups\`. Cada paso se anota en el diario antes de escribir; si algo falla a mitad, revierte y no guarda.

## Deshacer (ejecutor.py)

- Deshace el último lote confirmado, recorriendo el diario al revés: primero Excel, después los archivos.
- Solo toca:
  - filas que siguen como las dejó la app (compara A y J);
  - archivos con el mismo SHA256 que tenían al moverlos.
- Borra las carpetas que creó si quedaron vacías y saca las líneas que agregó a `proveedores.csv`.

## Comandos (python -m appfacturas …)

`ui` (por defecto) · `simular [--entrada]` · `procesar [--real]` · `deshacer` · `excel-pendiente` · `portapapeles` · `verificar` · `empaquetar` · `prueba-mover` · `prueba-excel` · `paridad preparar|comparar`

## Pruebas

- En este repo: `app\tests` trae las pruebas que no necesitan facturas reales (montos, CUIT/CBU, links, empaquetado).
- En la PC de desarrollo hay además 48 pruebas con facturas reales como fixtures (`app\tests\fixtures\`, fuera del repo por tener CUIT y CBU de proveedores). Cubren ARCA A/C/NC, MDQ, no fiscales, duplicados, columna H, .txt, mover + deshacer, Excel y exclusiones.
- También fuera del repo, en esa PC: `Probar mover archivos (prueba).bat`, `Probar Excel (prueba).bat` (Excel real, planilla cerrada y abierta; deshacer la dejó idéntica celda por celda) y `Prueba de paridad.bat` contra el script viejo (0 diferencias sin explicar).

## Pendiente

- Pestaña Revisar: auditoría como `-Revisar` del script viejo.
- Primera corrida real: Matheo habilita la escritura real.
- Modos con Claude: A (API, clave en el Administrador de credenciales de Windows) y B (Cowork, `_Revisar con Claude\<lote>\`).
- Retirar el PS1.
