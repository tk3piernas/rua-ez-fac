# RUA EZ-Fac: instrucciones para Claude Code

App de escritorio (Python 3.15 + NiceGUI) que carga las facturas de proveedores de RUA Asistencia (áreas Vehicular y Hogar) en `DEUDA A PAGAR ARG.xlsx`. Reemplaza al script `OrdenarFacturas.ps1`. Detalle técnico: `docs/DOCUMENTACION.md`. Guía de uso: `docs/Guia de uso RUA EZ-Fac.pdf`.

## Con quién hablás

- Matheo (Asesor de Alianzas en RUA). Respondé en español rioplatense, breve. Nunca uses la palabra "envergadura".
- Tiene un límite de uso ajustado: preferí leer texto y logs antes que capturas, y no releas archivos grandes sin necesidad.
- Cada vez que cambies algo, decile **dónde quedó**: la ruta completa en Windows de cada archivo nuevo o cambiado (por ejemplo `%USERPROFILE%\Documents\App Facturas\app\appfacturas\ui.py`, con la ruta real escrita).

## El repo es PÚBLICO: qué nunca se sube

`github.com/tk3piernas/rua-ez-fac` es público. Antes de cada commit revisá `git status` y `git diff --cached --stat`. Nunca commitees:

- Facturas, PDFs o zips de proveedores, ni CUIT, CBU, CVU, alias o números de cuenta reales (en tests usá números inventados con dígito verificador válido).
- `app/tests_privados/` (tests con facturas reales y sus fixtures), `datos/`, `prueba/`, `config.toml`, `.venv/`, `paridad/` y cualquier `.ps1`.
- Rutas o URLs personales nuevas (por ejemplo la URL de SharePoint de Matheo) dentro del código.

Todo eso ya está en `.gitignore`. Si algo privado aparece como "untracked", agregalo a `.gitignore` en vez de commitearlo. Si ves que ya se subió algo privado, frená y avisale a Matheo.

## Reglas de la app (no romper)

- La UI va en español rioplatense; el código y los comentarios pueden ir en inglés.
- La app, su `.venv` y su configuración viven en `Documents\App Facturas`, nunca dentro de OneDrive (el OneDrive de RUA bloquea .exe, .bat, .ps1 y .py). Las facturas y la planilla sí están en el OneDrive sincronizado (`OneDrive - Rua Asistencia`). No usar Microsoft Graph ni Azure.
- Excel por COM (pywin32), con la planilla abierta o cerrada. Nunca escribir sobre la columna I. Pisar H solo con 0 en Factura B, Factura C y no fiscal, o si el IVA real difiere de la fórmula en más de `tolerancia_iva`. CUIT, teléfono, N° de cuenta y CBU van como texto.
- `[app] modo_real` (Configuración → Modo de trabajo → "Cargar en la planilla real") arranca apagado. No escribas en la planilla ni en las carpetas reales para probar: usá la carpeta de prueba (`prueba\`) o los tests.
- Todo cambio de archivos o de la planilla tiene que poder deshacerse (diario + `deshacer`).
- Claude como respaldo: el modo B (Cowork, `claude_cowork.py`) ya está; el A (API) está pendiente; si se hace: la API key va en el Administrador de credenciales de Windows (keyring) o en una variable de entorno, nunca en un archivo; no mandar CBU ni datos bancarios; revisar la documentación actual del SDK `anthropic` antes de escribir código.

## Cómo probar

Desde la carpeta de la app (usa el Python del `.venv`):

```
.venv\Scripts\python.exe -m pytest -q app\tests
.venv\Scripts\python.exe -m pytest -q app\tests_privados
```

Se corren por separado porque cada carpeta tiene su `conftest.py`. Las dos tienen que pasar antes de commitear. Para ver la app: `.venv\Scripts\python.exe -m appfacturas ui` (si ya hay una abierta de otra versión, la cierra sola).

## Después de cada cambio

1. Correr las dos tandas de tests.
2. Subir la versión en `app/appfacturas/__init__.py` (0.x.y: y para arreglos, x para funciones nuevas).
3. Si cambia algo que ve el usuario, actualizar `docs/DOCUMENTACION.md`, `README.md` y, si hace falta, la guía PDF.
4. `git add` solo de lo que corresponde, revisar `git diff --cached --stat`, commit con mensaje claro en español y **`git push origin main`** (Matheo pidió que los cambios se suban solos).
5. Decirle a Matheo qué cambió y dónde quedaron los archivos. Si cambió algún `requirements*.txt`, recordarle correr `Instalar o actualizar RUA EZ-Fac.bat`; si no, alcanza con tocar Salir y abrir la app de nuevo.

## Pendiente

- Probar el reset de la planilla con Excel real, primero en la carpeta de prueba.
- Primera corrida real (Matheo elige "Cargar en la planilla real").
- Modo A con Claude (API).
