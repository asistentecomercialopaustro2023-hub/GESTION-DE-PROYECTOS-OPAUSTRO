# Gestión de Proyectos — OPAUSTRO

App interna de gestión de proyectos (estilo Monday.com), independiente de la app principal APP OPAUSTRO.

## Estructura

- **`index.html`** — Frontend (una sola página, sin build). Se sirve tal cual, por ejemplo con GitHub Pages.
- **`Code.gs`** — Backend en Google Apps Script (Web App). Se publica y actualiza directamente desde el editor de Apps Script, **no** desde este repositorio.
- **`CHANGELOG.md`** — Historial de versiones de la app y cómo revertir cada parte si algo sale mal.

## Cómo funciona

- Cada proyecto vive en una carpeta de Google Drive elegida por quien lo crea, con un Google Sheet ("Control") como base de datos.
- El backend (`Code.gs`) se despliega como Web App en Apps Script y expone una API (`doGet`/`doPost`) que `index.html` consume vía `fetch`.
- No hay build ni dependencias de npm: `index.html` es HTML/CSS/JS plano.

## Desplegar cambios

**Frontend (`index.html`):**
Con GitHub Pages activado sobre este repo, cualquier cambio empujado a la rama `main` se publica solo con esperar unos segundos — no requiere ningún paso manual.

**Backend (`Code.gs`):**
1. Copia el contenido de `Code.gs` al editor de Apps Script del proyecto.
2. `Implementar → Administrar implementaciones → ✏️ → Versión: Nueva versión → Implementar`.
3. Si se agregó un permiso nuevo (Drive, Docs, Sheets, Slides), ejecuta una función una vez desde el editor para volver a autorizar.

## Historial y rollback

Ver [`CHANGELOG.md`](CHANGELOG.md). En resumen:
- `index.html`: `git log` para ver versiones, `git checkout <commit> -- index.html`.
- `Code.gs`: desde el desplegable de versiones en Apps Script (se guardan automáticamente en cada "Nueva versión").

Los datos de los proyectos (tareas, archivos, comentarios) viven en Google Sheets/Drive y nunca se ven afectados por revertir una versión del código.
