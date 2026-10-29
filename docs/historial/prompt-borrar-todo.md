# Prompt: construir "Borrar todo" en RUA EZ-Fac

Pegalo en una sesión nueva de Claude Code abierta en la carpeta de la app (`%USERPROFILE%\Documents\App Facturas`).

---

Leé `CLAUDE.md` y respetalo en todo (español rioplatense, breve, nunca la palabra "envergadura"; repo público y qué nunca se sube; cómo probar; versión, commit y push; decirme la ruta completa de cada archivo que toques). Después leé `docs/DOCUMENTACION.md` (secciones de Setup, reset, tandas, deshacer y modo de trabajo) y `app/appfacturas/preparar.py`, `tandas.py`, `ejecutor.py` y `ui.py` en lo que haga falta, para reusar lo que ya existe (respaldos, reset, `vaciar_rango`, `con_espera`, `de_a_uno`, `mientras_libre`, `puede_escribir`, `SOLO_SIMULAR`).

Vamos a construir el botón **"Borrar todo"**: deja la app como recién empezada. Hay que ser **extremadamente cuidadoso**: toca el OneDrive de RUA, que se sincroniza con otras personas. Un error pierde datos de la empresa.

## Estado de partida

- Hay un respaldo completo y verificado en `%USERPROFILE%\Documents\Respaldo RUA EZ-Fac 2026-10-28 1931.zip` (lo de la app en el OneDrive, `datos\`, `prueba\` y `config.toml`, con `LISTA.json` y el SHA256 de cada archivo). No lo toques.
- El código antes de empezar está marcado en Git con la etiqueta `respaldo-antes-de-borrar-todo` (versión 0.18.1).
- Nunca pruebes contra lo real. Todo se prueba en carpetas temporales (tests), en la carpeta de prueba (modo "Practicar en la carpeta de prueba") y en una instalación de juguete con `APPFACTURAS_HOME` apuntando a una carpeta temporal y otro puerto.

## Decisiones ya tomadas (no las vuelvas a preguntar)

1. Se lleva también las facturas ya guardadas en las carpetas de proveedores, y vacía todas las hojas de la planilla, incluida Proveedores.
2. No borra: **mueve** todo a carpetas `_Borrado <fecha>` y se puede deshacer. "Vaciar definitivamente" va aparte, con su propia confirmación.
3. No toca `config.toml`, la clave de Claude (Administrador de credenciales) ni el registro de que se vio el recorrido.

## Qué toca (lista cerrada; nunca "todo lo que haya en la carpeta")

- En el OneDrive (la carpeta padre de `Facturas pendientes`): `Facturas pendientes\` entero (carpetas de área y de proveedores con sus PDF, `_Entrada` con `_Duplicados` y `_Procesados`, `_Script\proveedores.csv`, `_Revisar con Claude`), `Planillas exportadas\`, `Respaldos\`, `PROVEEDORES RUA.xlsx` y los `DEUDA A PAGAR ARG - respaldo *.xlsx` que la app dejó en la raíz.
- La planilla `DEUDA A PAGAR ARG.xlsx` **no se mueve** (los links compartidos tienen que seguir andando): se guarda una copia en `_Borrado` y se vacía en el mismo archivo (las hojas de área, filas 12 a 210 como el reset, sin tocar la columna I ni el formato, y Proveedores A5:O300).
- En la app: `datos\` (salvo el registro del recorrido) y `prueba\`.

## Qué no toca nunca

`config.toml`, la clave, `.venv\`, el código, y **todo lo demás del OneDrive**: en la raíz del OneDrive hay carpetas y planillas personales que no son de la app (escritorio, imágenes, adjuntos, chats, otras planillas). Tiene que haber un test con archivos "señuelo" en esos lugares (carpetas y `.xlsx` sueltos en la raíz del OneDrive con otros nombres) que falle si alguno se mueve, cambia o desaparece.

## Cómo funciona

1. **Antes de mover nada**: revisa que la planilla no esté abierta en Excel con cambios sin guardar y que ningún archivo esté tomado (intentar abrir cada uno). Si algo falla, no hace nada y dice qué.
2. Hace un zip de respaldo como el de partida en `%USERPROFILE%\Documents` (fuera del OneDrive y fuera de la app), y lo verifica.
3. **Mueve** cada cosa a `OneDrive\_Borrado <fecha>\` (lo del OneDrive) y a `<app>\_Borrado <fecha>\` (lo de la app). Es el mismo disco: mover es renombrar. Anota cada movimiento en un registro (`borrado.json`) **antes** del siguiente. Si algo falla a mitad, frena y devuelve lo que ya movió.
4. Copia la planilla a `_Borrado` y la vacía por Excel (COM, como el reset).
5. Vuelve a crear la estructura vacía con Setup (carpetas y `proveedores.csv` con el encabezado), así la app sigue usable.
6. **Deshacer el borrado total**: devuelve cada cosa a su lugar. Si después del borrado apareció algo con el mismo nombre, no lo pisa: lo deja en `_Borrado` y lo avisa. La planilla vuelve de la copia solo si nadie la cambió desde el borrado (SHA256).
7. **Vaciar definitivamente**: confirmación aparte (escribir VACIAR); manda las carpetas `_Borrado` a la papelera (la de Windows; en OneDrive va a la papelera de OneDrive). Nunca `rmtree` directo.

## Pantalla

- En Configuración, al final, una sección roja **"Zona de peligro"** con el ícono ☢ y el botón **Borrar todo**.
- Deshabilitado en "Solo simular". En **"Practicar en la carpeta de prueba"** actúa **solo** sobre la copia de prueba (así se prueba entero sin riesgo); en "Cargar en la planilla real", sobre lo real.
- Primera confirmación: qué se va exactamente, con cantidades ("N PDF en M carpetas de proveedores, X filas en la planilla, …"), qué queda, y el aviso de que el OneDrive se sincroniza con otras personas.
- Segunda confirmación: escribir **BORRAR TODO** exacto; el botón se habilita recién a los 5 segundos.
- Después, los botones **Deshacer el borrado total** y **Vaciar definitivamente** con su estado.
- Usá `con_espera`, `de_a_uno` y `mientras_libre`, y frená también dentro de cada función (no solo en el botón), como en el resto de la app.

## Paso a paso (frená al final de cada paso y esperá mi OK)

- **Paso 2**: `app/appfacturas/borrado.py` (inventario, chequeo previo, borrar, deshacer, vaciar) y sus tests en `app/tests` con datos inventados: borrar y deshacer deja todo idéntico (comparar SHA256 de cada archivo); los señuelos no se tocan; un fallo a mitad revierte; deshacer no pisa archivos nuevos; la planilla queda vacía y vuelve. Sin pantalla todavía.
- **Paso 3**: la pantalla, probada en la instalación de juguete (nunca en la real).
- **Paso 4**: yo lo pruebo en la carpeta de prueba.
- **Paso 5**: usarlo en lo real, solo si yo lo decido.

En cada versión: las dos tandas de tests, subir la versión (es una función nueva: 0.19.0), actualizar `docs/DOCUMENTACION.md`, `README.md` y, si hace falta, la guía PDF, commit en español y `git push origin main`. Antes de escribir código, decime en diez líneas cómo vas a garantizar que nunca toque nada fuera de la lista cerrada.
