# Prompt: revisión de código de RUA EZ-Fac (para Claude Fable)

Pegalo en una sesión nueva de Claude Code abierta en `%USERPROFILE%\Documents\App Facturas`, con el modelo Fable.

---

Leé `CLAUDE.md`, `README.md` y `docs/DOCUMENTACION.md`. Después revisá todo el código de `app/appfacturas/` y los tests de `app/tests/` (y `app/tests_privados/` solo para entender qué se prueba; no copies nada de ahí). Es una app de escritorio (Python + NiceGUI) que carga facturas de proveedores de RUA Asistencia en `DEUDA A PAGAR ARG.xlsx` por Excel COM, mueve los PDF a las carpetas de cada proveedor en OneDrive y guarda los datos de los proveedores en `PROVEEDORES RUA.xlsx`. La usa una persona sola en su PC, con Windows y Excel.

**En esta sesión no cambies código, no hagas commits y no escribas en la planilla, en OneDrive ni en `datos/`.** Podés leer y correr las dos tandas de tests (cada una con su comando, como dice `CLAUDE.md`). Solo para ver la app, abrila con la carpeta de prueba; ojo, si ya hay una abierta de otra versión, la cierra.

Quiero un informe con sugerencias, ordenado por prioridad, en español rioplatense y sin la palabra "envergadura". Para cada punto: archivo y línea, qué pasa hoy (con un caso concreto: qué entrada produce qué resultado), por qué importa y el cambio que proponés, con un esfuerzo estimado (chico, mediano o grande).

Revisá en este orden:

1. **Que no se pierdan ni se corrompan datos.** Todo cambio en la planilla, en los archivos y en la memoria de proveedores tiene que poder deshacerse (diario + `deshacer`). Buscá caminos donde no: errores a mitad de un lote, Excel abierto con cambios sin guardar, OneDrive sincronizando, dos clics seguidos, deshacer después de que alguien tocó algo, la planilla de prueba contra la real.
2. **Reglas de la planilla** (`CLAUDE.md`): nunca escribir en la columna I; pisar H solo en los casos que dice; CUIT, teléfono, N° de cuenta y CBU como texto; las hojas por área (Vehicular, Hogar, Persona, ServiciosYa).
3. **Lectura de facturas:** dónde puede leer mal un monto, una fecha, un CUIT, un N° de tarea o el tipo de comprobante, y qué casos faltan en los tests.
4. **Seguridad y privacidad:** el repo `github.com/tk3piernas/rua-ez-fac` es público. Revisá que el código y los tests no tengan datos reales (CUIT, CBU, alias, nombres de proveedores reales, rutas o URLs personales) y que `.gitignore` cubra todo lo privado.
5. **Interfaz (`ui.py`):** errores que el usuario no ve (avisos que se pierden al refrescar un panel, tareas largas sin indicador), textos confusos y pasos que se podrían ahorrar.
6. **Código:** duplicación, funciones demasiado largas, cosas difíciles de testear y tests que faltan. Solo lo que valga la pena, no estilo.
7. **Funciones nuevas que sumarían**, pensando en el uso diario y en lo pendiente de `CLAUDE.md` (pestaña "Revisar", modos con Claude A y B, retirar el PS1, exportar planillas por tanda). Como mucho cinco, cada una con cómo la harías.

Cerrá con un resumen de 5 líneas: lo más urgente y lo que conviene hacer primero.
