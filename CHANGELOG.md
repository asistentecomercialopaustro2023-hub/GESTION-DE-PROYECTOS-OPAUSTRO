# Historial de versiones — Gestión de Proyectos

Este archivo lleva el control de versiones de la **app** (`index.html` y `Code.gs`).

**Importante:** los proyectos, tareas y avances que ya creaste **no dependen de estos archivos** —
viven en Google Sheets y Drive, cada proyecto en su propia carpeta. Restaurar una versión
anterior de la app **nunca borra ni afecta** tus proyectos existentes.

Cómo restaurar una versión anterior:
- **Frontend (`index.html`):** `git log` para ver versiones, `git checkout <commit> -- index.html`
  para traer esa versión de vuelta (o pide que se haga, dando el número o descripción de versión).
- **Backend (`Code.gs`):** en Apps Script, Implementar → Administrar implementaciones → editar →
  en el desplegable "Versión" elige una versión anterior (Apps Script las guarda automáticamente
  cada vez que se implementa como "Nueva versión") → Implementar. No requiere tocar el código.

## v1.0 — 2026-09-15
Punto de referencia inicial con todo lo construido hasta ahora:
- Home con logo, listado de proyectos, creación de proyecto (validación completa,
  enlace de carpeta de Drive propia, sin plantilla automática).
- Acceso por proyecto: selección de usuario, contraseña solo si es Admin, Enter para
  ingresar, sin necesidad de cuenta de Google.
- Tabla principal con subtareas anidadas sin límite, agrupación Completado solo por
  el estado del módulo principal (no por subtareas individuales).
- Kanban, Gantt, Sprints/Backlog, Reportes (Excel/PDF).
- Enlaces y documentos en línea por módulo (Word/Excel/PowerPoint vía Google Docs/
  Sheets/Slides), compartidos sin login, con logos por tipo y registro propio de
  accesos (pestaña "Accesos" del Sheet).
- Persistencia de sesión y pestaña activa entre refrescos (localStorage), con cierre
  automático por 1 hora de inactividad.
- Tabla responsiva sin scroll horizontal.
