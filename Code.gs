// =========================================================
// GESTIÓN DE PROYECTOS — Backend (Google Apps Script)
// App independiente (no forma parte de APP OPAUSTRO).
//
// Estructura en Drive:
//   RAIZ (ROOT_FOLDER_ID)
//     └─ <Nombre del Proyecto>/          (una carpeta por proyecto)
//          ├─ Control                    (Google Sheet: Config/Usuarios/Tareas/Sprints/Archivos)
//          ├─ <Usuario 1>/                (carpeta por participante, para sus documentos)
//          ├─ <Usuario 2>/
//          └─ General/                   (archivos sin responsable único)
//
// Seguridad: es una herramienta interna de confianza. Los
// participantes solo eligen su nombre (sin contraseña). Los
// admins requieren contraseña, guardada como hash SHA-256
// (no reversible, pero sin salt — suficiente para uso interno,
// no para datos sensibles).
// =========================================================

var ROOT_FOLDER_ID = '1-xGBh-Dbt4goaKYj5J9Tg3Yr4iVSpI6J';

// Todas las fechas/horas de la app usan siempre horario de Ecuador
// (America/Guayaquil, sin horario de verano), sin importar la zona horaria
// configurada en el proyecto de Apps Script o en la hoja de cálculo.
var TZ = 'America/Guayaquil';
function hoyEcuador_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function ahoraEcuadorStamp_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'); }
function ahoraEcuadorISO_() { return Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd'T'HH:mm:ss"); }

var TABS = {
  CONFIG: 'Config',
  USUARIOS: 'Usuarios',
  TAREAS: 'Tareas',
  SPRINTS: 'Sprints',
  ARCHIVOS: 'Archivos',
  ACCESOS: 'Accesos',
  HISTORIAL: 'Historial'
};

// ============================================================
// ENTRY POINTS
// ============================================================
function doGet(e) {
  try {
    var p = (e && e.parameter) || {};
    var action = p.action || '';
    if (!action) return jsonResponse({ ok: true, app: 'Gestión de Proyectos API', ts: new Date().toISOString() });
    if (action === 'listProjects') return jsonResponse(listProjects());
    if (action === 'getProjectData') return jsonResponse(getProjectData(p.projectId));
    if (action === 'getDocAccessLog') return jsonResponse(getDocAccessLog(p));
    return jsonResponse({ ok: false, error: 'Acción no reconocida: ' + action });
  } catch (err) {
    return jsonResponse({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  try {
    var body = {};
    if (e && e.postData && e.postData.contents) body = JSON.parse(e.postData.contents);
    var action = body.action || '';
    var handlers = {
      createProject: createProject,
      login: login,
      addParticipant: addParticipant,
      saveConfig: saveConfig,
      addTask: addTask,
      updateTask: updateTask,
      deleteTask: deleteTask,
      addSprint: addSprint,
      updateSprint: updateSprint,
      uploadFile: uploadFile,
      deleteFile: deleteFile,
      importTareas: importTareas,
      addTaskLink: addTaskLink,
      createOnlineDoc: createOnlineDoc,
      deleteTaskLink: deleteTaskLink,
      logDocAccess: logDocAccess,
      addTaskComment: addTaskComment,
      deleteTaskComment: deleteTaskComment
    };
    if (!handlers[action]) return jsonResponse({ ok: false, error: 'Acción no reconocida: ' + action });
    var result = handlers[action](body);
    if (body.projectId) invalidateProjectCache_(body.projectId);
    else { try { CacheService.getScriptCache().remove('listProjects'); } catch (e) { /* no-op */ } } // ej. createProject
    return jsonResponse(result);
  } catch (err) {
    return jsonResponse({ ok: false, error: String(err) });
  }
}

function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// HELPERS: Drive / Sheet
// ============================================================
function rootFolder_() { return DriveApp.getFolderById(ROOT_FOLDER_ID); }

function getProjectFolder_(projectId) {
  return DriveApp.getFolderById(projectId);
}

// ============================================================
// ÍNDICE DE PROYECTOS
// Cada proyecto puede vivir en una carpeta de Drive distinta (la que
// elija quien lo crea), así que se guarda un índice liviano en
// PropertiesService para poder listarlos a todos sin importar dónde estén.
// ============================================================
function getIndex_() {
  var raw = PropertiesService.getScriptProperties().getProperty('PROJECT_INDEX');
  return raw ? JSON.parse(raw) : [];
}
function saveIndex_(arr) {
  PropertiesService.getScriptProperties().setProperty('PROJECT_INDEX', JSON.stringify(arr));
}
function addToIndex_(id, nombre, controlSheetId) {
  var idx = getIndex_();
  var found = idx.filter(function (p) { return p.id === id; })[0];
  if (!found) {
    idx.push({ id: id, nombre: nombre, controlSheetId: controlSheetId || '' });
    saveIndex_(idx);
  } else if (controlSheetId && !found.controlSheetId) {
    found.controlSheetId = controlSheetId;
    saveIndex_(idx);
  }
}

// Evita tener que buscar el Sheet "Control" por nombre dentro de la carpeta
// (una llamada lenta a Drive) en cada solicitud: se guarda su ID una sola
// vez en el índice y de ahí en adelante se abre directo por ID.
function getControlSheetId_(projectId) {
  var idx = getIndex_();
  var entry = idx.filter(function (p) { return p.id === projectId; })[0];
  if (entry && entry.controlSheetId) return entry.controlSheetId;
  var ss = getControlSheet_(getProjectFolder_(projectId));
  addToIndex_(projectId, entry ? entry.nombre : '', ss.getId());
  return ss.getId();
}
function openControlSheet_(projectId) {
  return SpreadsheetApp.openById(getControlSheetId_(projectId));
}

// Cache corto (segundos) del resultado de getProjectData, para que cambiar
// de pestaña o volver a abrir el mismo proyecto sea instantáneo. Se invalida
// automáticamente en doPost tras cualquier acción que reciba projectId.
function invalidateProjectCache_(projectId) {
  try {
    var cache = CacheService.getScriptCache();
    cache.remove('projdata_' + projectId);
    cache.remove('listProjects');
  } catch (e) { /* no-op */ }
}

function getOrCreateSubfolder_(parent, name) {
  var it = parent.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return parent.createFolder(name);
}

// Agrega una columna con este encabezado si todavía no existe en la hoja
// (permite añadir campos nuevos a proyectos creados antes de este cambio).
function ensureColumn_(sh, headerName) {
  var lastCol = sh.getLastColumn();
  var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var idx = headers.indexOf(headerName);
  if (idx > -1) return idx + 1;
  var newCol = lastCol + 1;
  sh.getRange(1, newCol).setValue(headerName);
  return newCol;
}

function getControlSheet_(projectFolder) {
  var it = projectFolder.getFilesByName('Control');
  if (!it.hasNext()) throw new Error('No se encontró el Sheet de Control en este proyecto.');
  return SpreadsheetApp.open(it.next());
}

// Por pedido explícito: la contraseña se guarda TAL CUAL en la columna
// PasswordHash de la pestaña Usuarios (puede ser números, texto o ambos),
// para que un admin pueda escribirla o cambiarla directamente en el Sheet.
// Sin cifrado — aceptable porque es una herramienta interna con pocos
// proyectos y participantes de confianza.
function hashPassword_(pw) {
  return String(pw || '');
}

// Todas estas funciones leen encabezados + datos en UNA sola llamada a
// getValues() (en vez de una lectura para encabezados y otra para datos, o
// peor, una lectura por fila) — cada llamada a Sheets tiene su propio costo
// fijo en Apps Script, así que juntarlas es lo que hace que agregar,
// actualizar o eliminar se sienta inmediato en vez de demorado.

// Lee una pestaña como array de objetos usando la fila 1 como encabezados.
function readTable_(ss, tabName) {
  var sh = ss.getSheetByName(tabName);
  if (!sh || sh.getLastRow() < 2) return [];
  var all = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
  var headers = all[0];
  return all.slice(1).map(function (row) {
    var obj = {};
    headers.forEach(function (h, i) { obj[h] = row[i]; });
    return obj;
  }).filter(function (o) { return o[headers[0]] !== '' && o[headers[0]] !== null; });
}

function appendRow_(ss, tabName, obj) {
  var sh = ss.getSheetByName(tabName);
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var row = headers.map(function (h) { return obj[h] !== undefined ? obj[h] : ''; });
  sh.appendRow(row);
}

function updateRowById_(ss, tabName, idField, idValue, patch) {
  var sh = ss.getSheetByName(tabName);
  if (!sh || sh.getLastRow() < 2) return false;
  var lastCol = sh.getLastColumn();
  var all = sh.getRange(1, 1, sh.getLastRow(), lastCol).getValues();
  var headers = all[0];
  var idCol = headers.indexOf(idField);
  for (var i = 1; i < all.length; i++) {
    if (String(all[i][idCol]) === String(idValue)) {
      headers.forEach(function (h, c) { if (patch[h] !== undefined) all[i][c] = patch[h]; });
      sh.getRange(i + 1, 1, 1, lastCol).setValues([all[i]]);
      return true;
    }
  }
  return false;
}

function deleteRowsWhere_(ss, tabName, idField, idValues) {
  var sh = ss.getSheetByName(tabName);
  if (!sh || sh.getLastRow() < 2) return;
  var lastRow = sh.getLastRow();
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var idCol = headers.indexOf(idField);
  // Antes esto leía celda por celda dentro del bucle (una llamada a Sheets
  // por fila); ahora se lee la columna de IDs completa de una sola vez.
  var idColValues = sh.getRange(2, idCol + 1, lastRow - 1, 1).getValues();
  for (var r = lastRow; r >= 2; r--) {
    var val = String(idColValues[r - 2][0]);
    if (idValues.indexOf(val) > -1) sh.deleteRow(r);
  }
}

// Escribe un árbol de tareas {titulo, prioridad, estado, vencimiento, notas,
// cronogramaInicio, cronogramaFin, hijos:[...]} en la pestaña Tareas,
// agregándolo después de las filas que ya existan (no las borra).
function writeSeedTareas_(sh, nodes) {
  var rows = [];
  var hoy = hoyEcuador_();
  var writeNode = function (node, parentId) {
    var id = Utilities.getUuid();
    rows.push([
      id, parentId || '', node.titulo || 'Tarea', '', node.estado || 'No iniciado',
      node.vencimiento || '', node.prioridad || 'Media', node.notas || '',
      node.cronogramaInicio || '', node.cronogramaFin || '', '', hoy
    ]);
    (node.hijos || []).forEach(function (h) { writeNode(h, id); });
  };
  nodes.forEach(function (n) { writeNode(n, ''); });
  if (rows.length) sh.getRange(sh.getLastRow() + 1, 1, rows.length, 12).setValues(rows);
  return rows.length;
}

// Recolecta un id de tarea y todos sus descendientes (para borrado en cascada).
function collectDescendants_(tareas, id) {
  var out = [id];
  tareas.filter(function (t) { return String(t.ParentID) === String(id); })
    .forEach(function (t) { out = out.concat(collectDescendants_(tareas, t.ID)); });
  return out;
}

// ============================================================
// PROYECTOS
// ============================================================
function listProjects() {
  var cache = CacheService.getScriptCache();
  var cached;
  try { cached = cache.get('listProjects'); } catch (e) { cached = null; }
  if (cached) return JSON.parse(cached);

  var idx = getIndex_();

  // Migración: la primera vez, si el índice está vacío, se buscan los
  // proyectos creados con la versión anterior (todos bajo ROOT_FOLDER_ID)
  // para no perderlos.
  if (idx.length === 0) {
    try {
      var legacy = rootFolder_().getFolders();
      while (legacy.hasNext()) {
        var lf = legacy.next();
        addToIndex_(lf.getId(), lf.getName());
      }
      idx = getIndex_();
    } catch (e) { /* sin acceso a la carpeta original: se ignora */ }
  }

  var out = [];
  idx.forEach(function (entry) {
    try {
      var ss;
      if (entry.controlSheetId) {
        ss = SpreadsheetApp.openById(entry.controlSheetId);
      } else {
        // Proyecto sin controlSheetId guardado todavía (creado antes de este
        // cambio): se busca una vez por carpeta y se guarda para la próxima.
        var folder = DriveApp.getFolderById(entry.id);
        ss = getControlSheet_(folder);
        addToIndex_(entry.id, entry.nombre, ss.getId());
      }
      var config = readTable_(ss, TABS.CONFIG)[0] || {};
      var tareas = readTable_(ss, TABS.TAREAS);
      var total = tareas.length;
      var hechas = tareas.filter(function (t) { return t.Estado === 'Hecho'; }).length;
      var usuarios = readTable_(ss, TABS.USUARIOS).map(function (u) { return { nombre: u.Nombre, rol: u.Rol, email: u.Email || '' }; });
      out.push({
        id: entry.id, nombre: config.Nombre || entry.nombre,
        objetivo: config.Objetivo || '', resultados: config.Resultados || '',
        creado: config.Creado || '', total: total, hechas: hechas, usuarios: usuarios
      });
    } catch (err) { /* carpeta eliminada, inaccesible o sin Control válido: se ignora */ }
  });
  var result = { ok: true, proyectos: out };
  try { cache.put('listProjects', JSON.stringify(result), 20); } catch (e) { /* no-op */ }
  return result;
}

function createProject(body) {
  var nombre = (body.nombre || '').trim();
  if (!nombre) return { ok: false, error: 'El nombre del proyecto es obligatorio.' };
  var adminNombre = (body.adminNombre || '').trim();
  var adminPassword = body.adminPassword || '';
  var adminEmail = (body.adminEmail || '').trim(); // opcional: ya no se exige ni se usa para compartir
  if (!adminNombre || !adminPassword) return { ok: false, error: 'Debes definir un administrador y su contraseña.' };
  var rootFolderId = (body.rootFolderId || '').trim();
  if (!rootFolderId) return { ok: false, error: 'Debes indicar el enlace de la carpeta de Drive donde se guardará el proyecto.' };
  var participantes = body.participantes || []; // [{nombre, rol}]

  var parentFolder;
  try {
    parentFolder = DriveApp.getFolderById(rootFolderId);
  } catch (e) {
    return { ok: false, error: 'No se pudo acceder a esa carpeta de Drive. Verifica el enlace y que la carpeta esté compartida con la cuenta que ejecuta este backend.' };
  }

  var folder = parentFolder.createFolder(nombre);
  var ss = SpreadsheetApp.create('Control');
  DriveApp.getFileById(ss.getId()).moveTo(folder);

  ss.getSheets()[0].setName(TABS.CONFIG);
  ss.getSheetByName(TABS.CONFIG).getRange(1, 1, 1, 4).setValues([['Nombre', 'Objetivo', 'Resultados', 'Creado']]);
  ss.getSheetByName(TABS.CONFIG).getRange(2, 1, 1, 4).setValues([[nombre, body.objetivo || '', body.resultados || '', hoyEcuador_()]]);

  var shUsuarios = ss.insertSheet(TABS.USUARIOS);
  shUsuarios.getRange(1, 1, 1, 4).setValues([['Nombre', 'Rol', 'PasswordHash', 'Email']]);

  var shTareas = ss.insertSheet(TABS.TAREAS);
  shTareas.getRange(1, 1, 1, 14).setValues([['ID', 'ParentID', 'Titulo', 'Responsable', 'Estado', 'Vencimiento', 'Prioridad', 'Notas', 'CronogramaInicio', 'CronogramaFin', 'SprintID', 'Actualizado', 'EnlacesDocumento', 'Comentarios']]);

  // Plantilla opcional: si el creador eligió cargar una estructura de tareas
  // ya definida (ej. "PROYECTO DE CAPACITACIONES"), se escribe todo de una vez.
  if (body.seedTareas && body.seedTareas.length) writeSeedTareas_(shTareas, body.seedTareas);

  var shSprints = ss.insertSheet(TABS.SPRINTS);
  shSprints.getRange(1, 1, 1, 5).setValues([['ID', 'Nombre', 'Inicio', 'Fin', 'Meta']]);

  var shArchivos = ss.insertSheet(TABS.ARCHIVOS);
  shArchivos.getRange(1, 1, 1, 7).setValues([['ID', 'TareaID', 'Nombre', 'DriveFileId', 'Url', 'SubidoPor', 'Fecha']]);

  var shAccesos = ss.insertSheet(TABS.ACCESOS);
  shAccesos.getRange(1, 1, 1, 6).setValues([['Fecha', 'TareaID', 'LinkId', 'Documento', 'Usuario', 'Rol']]);

  var shHistorial = ss.insertSheet(TABS.HISTORIAL);
  shHistorial.getRange(1, 1, 1, 6).setValues([['Fecha', 'Usuario', 'Rol', 'Accion', 'Tarea', 'Detalle']]);

  // fila admin
  appendRow_(ss, TABS.USUARIOS, { Nombre: adminNombre, Rol: 'admin', PasswordHash: hashPassword_(adminPassword), Email: adminEmail });

  // participantes (evita duplicar al admin si aparece también en la lista)
  // Las carpetas de cada usuario ya no se crean aquí en la raíz: ahora se
  // organizan por módulo/tarea y se crean solas la primera vez que alguien
  // sube un archivo a esa tarea (ver uploadFile / getModuleFolder_).
  participantes.forEach(function (p) {
    var nombreP = (p.nombre || '').trim();
    if (!nombreP || nombreP.toLowerCase() === adminNombre.toLowerCase()) return;
    appendRow_(ss, TABS.USUARIOS, { Nombre: nombreP, Rol: p.rol === 'admin' ? 'admin' : 'participante', PasswordHash: p.rol === 'admin' ? hashPassword_(body.adminPassword) : '', Email: (p.email || '').trim() });
  });

  addToIndex_(folder.getId(), nombre, ss.getId());
  return { ok: true, projectId: folder.getId() };
}

function saveConfig(body) {
  var ss = openControlSheet_(body.projectId);
  var sh = ss.getSheetByName(TABS.CONFIG);
  sh.getRange(2, 2, 1, 2).setValues([[body.objetivo || '', body.resultados || '']]);
  if (body.nombre) {
    sh.getRange(2, 1).setValue(body.nombre);
    getProjectFolder_(body.projectId).setName(body.nombre);
  }
  return { ok: true };
}

// El Responsable ahora admite VARIAS personas por tarea, guardadas como
// array JSON en la misma columna "Responsable" (ej. ["Ana Perez","Luis
// Gomez"]). Si la tarea todavía tiene el valor anterior (un solo nombre en
// texto plano, de antes de este cambio), se interpreta igual sin migrar nada.
function leerResponsables_(valorCelda) {
  if (!valorCelda) return [];
  try {
    var arr = JSON.parse(valorCelda);
    if (Array.isArray(arr)) return arr.filter(Boolean);
  } catch (e) { /* no era JSON: era el texto plano de la versión anterior */ }
  return [String(valorCelda)];
}
function esUnoDeLosResponsables_(valorCelda, usuario) {
  var buscado = String(usuario || '').trim().toLowerCase();
  return leerResponsables_(valorCelda).some(function (r) { return String(r).trim().toLowerCase() === buscado; });
}

// Google Sheets a veces guarda un texto de fecha como un objeto Date real.
// Si no se formatea explícitamente, JSON.stringify lo convierte a algo como
// "2026-09-14T05:00:00.000Z" — estas funciones evitan eso.
function fmtDateOnly_(val) {
  if (!val) return '';
  if (val instanceof Date) return Utilities.formatDate(val, TZ, 'yyyy-MM-dd');
  return String(val);
}
function fmtDateTime_(val) {
  if (!val) return '';
  if (val instanceof Date) return Utilities.formatDate(val, TZ, 'yyyy-MM-dd HH:mm');
  return String(val);
}

function getProjectData(projectId) {
  var cache = CacheService.getScriptCache();
  var cacheKey = 'projdata_' + projectId;
  var cached;
  try { cached = cache.get(cacheKey); } catch (e) { cached = null; }
  if (cached) return JSON.parse(cached);

  var ss = openControlSheet_(projectId);
  var config = readTable_(ss, TABS.CONFIG)[0] || {};
  var usuarios = readTable_(ss, TABS.USUARIOS).map(function (u) { return { nombre: u.Nombre, rol: u.Rol, email: u.Email || '' }; }); // nunca exponer PasswordHash
  var tareasFlat = readTable_(ss, TABS.TAREAS);
  var archivos = readTable_(ss, TABS.ARCHIVOS);
  var sprints = readTable_(ss, TABS.SPRINTS).map(function (s) { return { id: s.ID, nombre: s.Nombre, inicio: s.Inicio, fin: s.Fin, meta: s.Meta }; });

  var archivosPorTarea = {};
  archivos.forEach(function (a) {
    (archivosPorTarea[a.TareaID] = archivosPorTarea[a.TareaID] || []).push({ id: a.ID, nombre: a.Nombre, driveFileId: a.DriveFileId, url: a.Url, subidoPor: a.SubidoPor, fecha: a.Fecha });
  });

  var byId = {};
  tareasFlat.forEach(function (t) {
    var responsables = leerResponsables_(t.Responsable);
    byId[t.ID] = {
      id: t.ID, parentId: t.ParentID || null, titulo: t.Titulo,
      // "responsables" (array) es lo nuevo; "responsable" (texto, unidos con
      // coma) se mantiene para lo que solo necesita mostrarlo como texto
      // (ej. reportes, mensajes de permiso), sin tener que tocar cada lugar.
      responsables: responsables, responsable: responsables.join(', '),
      estado: t.Estado || 'No iniciado', vencimiento: fmtDateOnly_(t.Vencimiento), prioridad: t.Prioridad || 'Media',
      notas: t.Notas || '', cronogramaInicio: fmtDateOnly_(t.CronogramaInicio), cronogramaFin: fmtDateOnly_(t.CronogramaFin),
      sprintId: t.SprintID || null, actualizado: fmtDateTime_(t.Actualizado), expandido: false,
      enlaces: leerEnlaces_(t),
      // Compatibilidad: si la tarea tenía una nota de texto simple (versión
      // anterior) y todavía no tiene comentarios, se muestra como el primero.
      comentarios: (function () { var c = leerComentarios_(t); return c.length ? c : (t.Notas ? [{ id: 'legacy', texto: String(t.Notas), usuario: '', fecha: '' }] : []); })(),
      archivos: archivosPorTarea[t.ID] || [], hijos: []
    };
  });
  var roots = [];
  tareasFlat.forEach(function (t) {
    var node = byId[t.ID];
    if (t.ParentID && byId[t.ParentID]) byId[t.ParentID].hijos.push(node);
    else roots.push(node);
  });

  // Las subtareas "Revisión..." siempre se muestran al final de cada módulo,
  // en orden entre ellas — sin tocar el Sheet, solo al armar la respuesta.
  // El resto de subtareas conserva su orden de siempre. Los módulos (tareas
  // de nivel superior) no se reordenan, solo sus subtareas.
  // IMPORTANTE: esto es una regla puntual SOLO para este proyecto
  // ("PROYECTO DE CAPACITACIONES"); los demás proyectos que se creen
  // mantienen el orden natural en que se van agregando sus tareas.
  if (String(config.Nombre || '').trim().toLowerCase() === 'proyecto de capacitaciones') {
    var esRevision_ = function (titulo) { return /^revisi[oó]n/i.test(String(titulo || '').trim()); };
    var moverRevisionesAlFinal_ = function (hijos) {
      hijos.forEach(function (n) { if (n.hijos.length) moverRevisionesAlFinal_(n.hijos); });
      var normales = hijos.filter(function (n) { return !esRevision_(n.titulo); });
      var revisiones = hijos.filter(function (n) { return esRevision_(n.titulo); });
      hijos.length = 0;
      Array.prototype.push.apply(hijos, normales.concat(revisiones));
    };
    roots.forEach(function (modulo) { moverRevisionesAlFinal_(modulo.hijos); });
  }

  var result = {
    ok: true,
    proyecto: { id: projectId, nombre: config.Nombre || '', objetivo: config.Objetivo || '', resultados: config.Resultados || '', creado: config.Creado || '' },
    usuarios: usuarios, tareas: roots, sprints: sprints
  };
  // Cache breve: si el proyecto es muy grande y no cabe (límite ~100KB de
  // CacheService), simplemente no se cachea, sin afectar la respuesta.
  try { cache.put(cacheKey, JSON.stringify(result), 30); } catch (e) { /* no-op */ }
  return result;
}

// ============================================================
// USUARIOS / LOGIN
// ============================================================
function requireAdmin_(ss, adminUsuario, adminPassword) {
  var usuarios = readTable_(ss, TABS.USUARIOS);
  var u = usuarios.filter(function (x) { return String(x.Nombre).toLowerCase() === String(adminUsuario || '').toLowerCase(); })[0];
  if (!u || u.Rol !== 'admin') return 'Ese usuario no es administrador de este proyecto.';
  if (String(u.PasswordHash) !== hashPassword_(adminPassword)) return 'Contraseña de administrador incorrecta.';
  return null;
}

function login(body) {
  var ss = openControlSheet_(body.projectId);
  var usuarios = readTable_(ss, TABS.USUARIOS);
  var u = usuarios.filter(function (x) { return String(x.Nombre).toLowerCase() === String(body.usuario || '').toLowerCase(); })[0];
  if (!u) return { ok: false, error: 'Usuario no encontrado en este proyecto.' };
  if (u.Rol === 'admin') {
    if (String(u.PasswordHash) !== hashPassword_(body.password)) return { ok: false, error: 'Contraseña incorrecta.' };
  }
  return { ok: true, usuario: u.Nombre, rol: u.Rol };
}

function addParticipant(body) {
  var ss = openControlSheet_(body.projectId);
  var err = requireAdmin_(ss, body.adminUsuario, body.adminPassword);
  if (err) return { ok: false, error: err };
  var nombre = (body.nombre || '').trim();
  var email = (body.email || '').trim(); // opcional: ya no se exige ni se usa para compartir
  if (!nombre) return { ok: false, error: 'Nombre requerido.' };
  var existentes = readTable_(ss, TABS.USUARIOS);
  if (existentes.some(function (u) { return String(u.Nombre).toLowerCase() === nombre.toLowerCase(); })) return { ok: false, error: 'Ese usuario ya existe en el proyecto.' };
  var esAdmin = body.rol === 'admin';
  appendRow_(ss, TABS.USUARIOS, { Nombre: nombre, Rol: esAdmin ? 'admin' : 'participante', PasswordHash: esAdmin ? hashPassword_(body.nuevoPassword || body.adminPassword) : '', Email: email });
  return { ok: true };
}

// ============================================================
// TAREAS
// ============================================================
function addTask(body) {
  var ss = openControlSheet_(body.projectId);
  var id = Utilities.getUuid();
  appendRow_(ss, TABS.TAREAS, {
    ID: id, ParentID: body.parentId || '', Titulo: body.titulo || 'Nueva tarea', Responsable: '',
    Estado: 'No iniciado', Vencimiento: '', Prioridad: 'Media', Notas: '', CronogramaInicio: '', CronogramaFin: '',
    SprintID: '', Actualizado: fmtDateTime_(new Date())
  });
  return { ok: true, id: id };
}

// Campos con permisos especiales: solo el Admin puede reasignar los
// responsables; estado, prioridad y las fechas solo las puede cambiar
// alguno de los responsables de la tarea (o el Admin, si no hay ninguno asignado).
var CAMPO_SOLO_ADMIN = { responsables: true };
var CAMPO_SOLO_RESPONSABLE_O_ADMIN = { estado: true, prioridad: true, vencimiento: true, cronogramaInicio: true, cronogramaFin: true };

function updateTask(body) {
  var ss = openControlSheet_(body.projectId);
  var patch = body.patch || {};
  var camposRestringidos = Object.keys(patch).filter(function (k) { return CAMPO_SOLO_ADMIN[k] || CAMPO_SOLO_RESPONSABLE_O_ADMIN[k]; });

  var tarea = null, usuarioActual = null;
  if (camposRestringidos.length) {
    tarea = encontrarTarea_(ss, body.taskId);
    if (!tarea) return { ok: false, error: 'Tarea no encontrada.' };
    var usuarios = readTable_(ss, TABS.USUARIOS);
    usuarioActual = usuarios.filter(function (x) { return String(x.Nombre).toLowerCase() === String(body.usuario || '').toLowerCase(); })[0];
    var esAdmin = usuarioActual && usuarioActual.Rol === 'admin';
    var esResponsable = esUnoDeLosResponsables_(tarea.Responsable, body.usuario);
    for (var i = 0; i < camposRestringidos.length; i++) {
      var campo = camposRestringidos[i];
      if (CAMPO_SOLO_ADMIN[campo] && !esAdmin) {
        return { ok: false, error: 'Solo el Admin puede reasignar el responsable.' };
      }
      if (CAMPO_SOLO_RESPONSABLE_O_ADMIN[campo] && !esAdmin && !esResponsable) {
        return { ok: false, error: 'Solo ' + (leerResponsables_(tarea.Responsable).join(' / ') || 'el Admin') + ' puede modificar ese campo de esta tarea.' };
      }
    }
  }

  var fieldMap = { titulo: 'Titulo', estado: 'Estado', vencimiento: 'Vencimiento', prioridad: 'Prioridad', notas: 'Notas', cronogramaInicio: 'CronogramaInicio', cronogramaFin: 'CronogramaFin', sprintId: 'SprintID' };
  var row = {};
  Object.keys(patch).forEach(function (k) {
    if (k === 'responsables') row.Responsable = JSON.stringify((patch.responsables || []).filter(Boolean));
    else if (fieldMap[k]) row[fieldMap[k]] = patch[k];
  });
  row.Actualizado = fmtDateTime_(new Date());
  var okUpd = updateRowById_(ss, TABS.TAREAS, 'ID', body.taskId, row);

  // Deja constancia de quién hizo el cambio en campos con permisos
  // especiales, para poder distinguir una acción del Admin (que puede
  // saltarse la restricción) de una del propio responsable.
  if (okUpd && camposRestringidos.length && usuarioActual) {
    camposRestringidos.forEach(function (campo) {
      var detalle = campo === 'responsables' ? (patch.responsables || []).join(', ') : String(patch[campo] || '');
      registrarHistorial_(ss, body.usuario, usuarioActual.Rol, 'Cambiar ' + campo, tarea.Titulo, detalle || '(en blanco)');
    });
  }

  return { ok: okUpd };
}

// ============================================================
// COMENTARIOS DE TAREA (registro/control tipo bitácora, no una
// nota única: se guardan como un array JSON en una sola columna,
// igual que EnlacesDocumento).
// ============================================================
function leerComentarios_(tareaRow) {
  try { return tareaRow.Comentarios ? JSON.parse(tareaRow.Comentarios) : []; }
  catch (e) { return []; }
}
// Lee la fila de una tarea y, en la MISMA pasada, escribe los cambios que
// devuelva `mutar` (una sola lectura + una sola escritura de la hoja
// Tareas). Antes, agregar o borrar un comentario leía la tarea por un lado
// y la volvía a leer completa para escribirla por otro — el doble de
// llamadas a Sheets de las necesarias, y eso era lo que se sentía lento.
// `mutar(filaActual)` devuelve el objeto de cambios a aplicar, o null/false
// para no escribir nada (ej. si no está autorizado).
function leerYEscribirTarea_(ss, taskId, mutar) {
  var sh = ss.getSheetByName(TABS.TAREAS);
  var lastCol = sh.getLastColumn();
  var all = sh.getRange(1, 1, sh.getLastRow(), lastCol).getValues();
  var headers = all[0];
  var idCol = headers.indexOf('ID');
  for (var i = 1; i < all.length; i++) {
    if (String(all[i][idCol]) === String(taskId)) {
      var obj = {};
      headers.forEach(function (h, c) { obj[h] = all[i][c]; });
      var patch = mutar(obj);
      if (patch) {
        headers.forEach(function (h, c) { if (patch[h] !== undefined) all[i][c] = patch[h]; });
        sh.getRange(i + 1, 1, 1, lastCol).setValues([all[i]]);
      }
      return obj;
    }
  }
  return null;
}

function addTaskComment(body) {
  var ss = openControlSheet_(body.projectId);
  var texto = (body.texto || '').trim();
  if (!texto) return { ok: false, error: 'El comentario no puede estar vacío.' };
  ensureColumn_(ss.getSheetByName(TABS.TAREAS), 'Comentarios');
  var nuevo = { id: Utilities.getUuid(), texto: texto, usuario: body.usuario || '', fecha: fmtDateTime_(new Date()) };
  var tarea = leerYEscribirTarea_(ss, body.taskId, function (t) {
    var comentarios = leerComentarios_(t);
    // Migra la nota de texto simple de la versión anterior (columna Notas) a
    // la bitácora, la primera vez que se agrega un comentario nuevo, para no perderla.
    if (!comentarios.length && t.Notas) comentarios.push({ id: Utilities.getUuid(), texto: String(t.Notas), usuario: '', fecha: '' });
    comentarios.push(nuevo);
    return { Comentarios: JSON.stringify(comentarios), Actualizado: nuevo.fecha };
  });
  if (!tarea) return { ok: false, error: 'Tarea no encontrada.' };
  return { ok: true, comentario: nuevo };
}

// Solo el admin o quien escribió el comentario puede borrarlo.
function deleteTaskComment(body) {
  var ss = openControlSheet_(body.projectId);
  var error = null;
  var tarea = leerYEscribirTarea_(ss, body.taskId, function (t) {
    var comentarios = leerComentarios_(t);
    var propio = comentarios.filter(function (c) { return c.id === body.commentId; })[0];
    if (!propio) { error = 'Comentario no encontrado.'; return null; }
    if (propio.usuario !== body.usuario) {
      var usuarios = readTable_(ss, TABS.USUARIOS);
      var u = usuarios.filter(function (x) { return String(x.Nombre).toLowerCase() === String(body.usuario || '').toLowerCase(); })[0];
      if (!u || u.Rol !== 'admin') { error = 'Solo el autor o el Admin pueden borrar este comentario.'; return null; }
    }
    var restantes = comentarios.filter(function (c) { return c.id !== body.commentId; });
    return { Comentarios: JSON.stringify(restantes) };
  });
  if (!tarea) return { ok: false, error: 'Tarea no encontrada.' };
  if (error) return { ok: false, error: error };
  return { ok: true };
}

// Importa una plantilla de tareas (árbol) a un proyecto YA EXISTENTE,
// sin borrar las tareas que ya tenga. Solo el admin del proyecto puede hacerlo.
function importTareas(body) {
  var ss = openControlSheet_(body.projectId);
  var err = requireAdmin_(ss, body.adminUsuario, body.adminPassword);
  if (err) return { ok: false, error: err };
  if (!body.seedTareas || !body.seedTareas.length) return { ok: false, error: 'No se recibió ninguna tarea para importar.' };
  var agregadas = writeSeedTareas_(ss.getSheetByName(TABS.TAREAS), body.seedTareas);
  return { ok: true, agregadas: agregadas };
}

// ============================================================
// ENLACES / DOCUMENTOS EN LÍNEA (una tarea principal puede tener varios)
// Se guardan como JSON en la columna EnlacesDocumento: [{id,tipo,nombre,url}]
// tipo: 'enlace' | 'doc' | 'sheet' | 'slide'
// Solo el admin del proyecto puede agregar, crear o eliminar. Los
// participantes solo pueden abrirlos (eso lo controla el frontend).
// ============================================================
// Igual que los archivos subidos, los documentos en línea creados desde la
// app se organizan dentro de la carpeta del módulo (tarea principal) al que
// pertenecen: <Proyecto>/<Módulo>/Documentos en línea/.
function getDocsFolder_(projectFolder, tareasFlat, taskId) {
  var modulo = encontrarModulo_(tareasFlat, taskId);
  var moduloFolder = getOrCreateSubfolder_(projectFolder, (modulo && modulo.Titulo) ? modulo.Titulo : 'General');
  return getOrCreateSubfolder_(moduloFolder, 'Documentos en línea');
}
function leerEnlaces_(tareaRow) {
  try { return tareaRow.EnlacesDocumento ? JSON.parse(tareaRow.EnlacesDocumento) : []; }
  catch (e) { return []; }
}
function guardarEnlaces_(ss, taskId, enlaces) {
  ensureColumn_(ss.getSheetByName(TABS.TAREAS), 'EnlacesDocumento');
  updateRowById_(ss, TABS.TAREAS, 'ID', taskId, { EnlacesDocumento: JSON.stringify(enlaces) });
}
function encontrarTarea_(ss, taskId) {
  var tareas = readTable_(ss, TABS.TAREAS);
  return tareas.filter(function (t) { return String(t.ID) === String(taskId); })[0];
}

// Agrega un enlace ya existente (pegado por el admin). Se guarda tal cual,
// sin mover ni tocar ningún archivo de Drive.
function addTaskLink(body) {
  var ss = openControlSheet_(body.projectId);
  var err = requireAdmin_(ss, body.adminUsuario, body.adminPassword);
  if (err) return { ok: false, error: err };
  var row = encontrarTarea_(ss, body.taskId);
  if (!row) return { ok: false, error: 'Tarea no encontrada.' };
  var url = (body.url || '').trim();
  if (!url) return { ok: false, error: 'La URL es obligatoria.' };

  var enlaces = leerEnlaces_(row);
  var nuevo = { id: Utilities.getUuid(), tipo: 'enlace', nombre: url, url: url };
  enlaces.push(nuevo);
  guardarEnlaces_(ss, body.taskId, enlaces);
  return { ok: true, enlace: nuevo };
}

// Crea un documento nuevo (Google Docs/Sheets/Slides, equivalentes en línea
// a Word/Excel/PowerPoint) con el nombre que pida el admin, guardado dentro
// de la carpeta del proyecto correspondiente, y lo enlaza a la tarea.
function createOnlineDoc(body) {
  var ss = openControlSheet_(body.projectId);
  var err = requireAdmin_(ss, body.adminUsuario, body.adminPassword);
  if (err) return { ok: false, error: err };
  var tareasFlat = readTable_(ss, TABS.TAREAS);
  var row = tareasFlat.filter(function (t) { return String(t.ID) === String(body.taskId); })[0];
  if (!row) return { ok: false, error: 'Tarea no encontrada.' };
  var titulo = (body.titulo || '').trim();
  if (!titulo) return { ok: false, error: 'El nombre del documento es obligatorio.' };

  var file, url;
  if (body.tipo === 'doc') { var d = DocumentApp.create(titulo); file = DriveApp.getFileById(d.getId()); url = d.getUrl(); }
  else if (body.tipo === 'sheet') { var s = SpreadsheetApp.create(titulo); file = DriveApp.getFileById(s.getId()); url = s.getUrl(); }
  else if (body.tipo === 'slide') { var p = SlidesApp.create(titulo); file = DriveApp.getFileById(p.getId()); url = p.getUrl(); }
  else return { ok: false, error: 'Tipo de documento no válido.' };

  // Se organiza igual que los archivos subidos: dentro de la carpeta del
  // módulo (tarea principal) al que pertenece esta tarea/subtarea.
  var docsFolder = getDocsFolder_(getProjectFolder_(body.projectId), tareasFlat, body.taskId);
  file.moveTo(docsFolder);

  // Se comparte la carpeta "Documentos en línea" completa (no archivo por
  // archivo) como "cualquiera con el enlace, editor" — igual que compartir
  // una carpeta en OneDrive: todo lo que haya adentro (este documento y los
  // que se creen después) queda accesible sin pedir cuenta de Google ni
  // contraseña. Se bloquea que los editores cambien el compartir: solo el
  // admin (desde esta app) agrega/quita documentos.
  docsFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);
  docsFolder.setShareableByEditors(false);

  var enlaces = leerEnlaces_(row);
  var nuevo = { id: Utilities.getUuid(), tipo: body.tipo, nombre: titulo, url: url };
  enlaces.push(nuevo);
  guardarEnlaces_(ss, body.taskId, enlaces);
  return { ok: true, enlace: nuevo };
}

// Quita un enlace de la lista de la tarea (no borra el archivo de Drive,
// solo la referencia — así no se elimina por accidente un documento compartido).
function deleteTaskLink(body) {
  var ss = openControlSheet_(body.projectId);
  var err = requireAdmin_(ss, body.adminUsuario, body.adminPassword);
  if (err) return { ok: false, error: err };
  var row = encontrarTarea_(ss, body.taskId);
  if (!row) return { ok: false, error: 'Tarea no encontrada.' };
  var enlaces = leerEnlaces_(row).filter(function (l) { return l.id !== body.linkId; });
  guardarEnlaces_(ss, body.taskId, enlaces);
  return { ok: true };
}

// ============================================================
// REGISTRO DE ACCESOS A DOCUMENTOS
// Como los documentos se comparten sin pedir cuenta de Google, Google Docs
// no puede identificar quién editó qué. Este es un registro propio de la
// app: cada vez que alguien abre un documento desde aquí, se guarda su
// nombre (el mismo con el que entró al proyecto), no un correo de Google.
// ============================================================
function ensureAccesosSheet_(ss) {
  var sh = ss.getSheetByName(TABS.ACCESOS);
  if (!sh) {
    sh = ss.insertSheet(TABS.ACCESOS);
    sh.getRange(1, 1, 1, 6).setValues([['Fecha', 'TareaID', 'LinkId', 'Documento', 'Usuario', 'Rol']]);
  }
  return sh;
}

// ============================================================
// HISTORIAL / AUDITORÍA DE CAMBIOS RESTRINGIDOS
// El Admin puede hacer cualquier cambio aunque no sea el responsable de la
// tarea; este registro deja constancia de quién hizo cada cambio sensible
// (eliminar tarea, cambiar estado/prioridad/fechas o reasignar responsable),
// para poder distinguir una acción del Admin de una del propio responsable.
// ============================================================
function ensureHistorialSheet_(ss) {
  var sh = ss.getSheetByName(TABS.HISTORIAL);
  if (!sh) {
    sh = ss.insertSheet(TABS.HISTORIAL);
    sh.getRange(1, 1, 1, 6).setValues([['Fecha', 'Usuario', 'Rol', 'Accion', 'Tarea', 'Detalle']]);
  }
  return sh;
}
function registrarHistorial_(ss, usuario, rol, accion, tareaTitulo, detalle) {
  var sh = ensureHistorialSheet_(ss);
  sh.appendRow([ahoraEcuadorStamp_(), usuario || '', rol || '', accion || '', tareaTitulo || '', detalle || '']);
}

function logDocAccess(body) {
  var ss = openControlSheet_(body.projectId);
  var sh = ensureAccesosSheet_(ss);
  sh.appendRow([ahoraEcuadorISO_(), body.taskId || '', body.linkId || '', body.documento || '', body.usuario || '', body.rol || '']);
  return { ok: true };
}

function getDocAccessLog(p) {
  var ss = openControlSheet_(p.projectId);
  ensureAccesosSheet_(ss);
  var rows = readTable_(ss, TABS.ACCESOS).filter(function (r) { return String(r.LinkId) === String(p.linkId); });
  rows.sort(function (a, b) { return new Date(b.Fecha) - new Date(a.Fecha); });
  return {
    ok: true,
    accesos: rows.slice(0, 10).map(function (r) { return { fecha: r.Fecha, usuario: r.Usuario, rol: r.Rol }; })
  };
}

// Solo el Admin o el responsable de ESA tarea/subtarea pueden eliminarla
// (si no tiene responsable asignado, solo el Admin).
function deleteTask(body) {
  var ss = openControlSheet_(body.projectId);
  var tareas = readTable_(ss, TABS.TAREAS);
  var tarea = tareas.filter(function (t) { return String(t.ID) === String(body.taskId); })[0];
  if (!tarea) return { ok: false, error: 'Tarea no encontrada.' };

  var usuarios = readTable_(ss, TABS.USUARIOS);
  var u = usuarios.filter(function (x) { return String(x.Nombre).toLowerCase() === String(body.usuario || '').toLowerCase(); })[0];
  var esAdmin = u && u.Rol === 'admin';
  var esResponsable = esUnoDeLosResponsables_(tarea.Responsable, body.usuario);
  if (!esAdmin && !esResponsable) {
    return { ok: false, error: 'Solo el Admin o ' + (leerResponsables_(tarea.Responsable).join(' / ') || 'el responsable asignado') + ' pueden eliminar esta tarea.' };
  }

  registrarHistorial_(ss, body.usuario, u ? u.Rol : '', 'Eliminar tarea', tarea.Titulo, '');
  var ids = collectDescendants_(tareas, body.taskId);
  deleteRowsWhere_(ss, TABS.TAREAS, 'ID', ids);
  deleteRowsWhere_(ss, TABS.ARCHIVOS, 'TareaID', ids);
  return { ok: true };
}

// ============================================================
// SPRINTS
// ============================================================
function addSprint(body) {
  var ss = openControlSheet_(body.projectId);
  var id = Utilities.getUuid();
  appendRow_(ss, TABS.SPRINTS, { ID: id, Nombre: body.nombre || 'Sprint', Inicio: body.inicio || '', Fin: body.fin || '', Meta: body.meta || '' });
  return { ok: true, id: id };
}

function updateSprint(body) {
  var ss = openControlSheet_(body.projectId);
  var patch = {};
  if (body.nombre !== undefined) patch.Nombre = body.nombre;
  if (body.inicio !== undefined) patch.Inicio = body.inicio;
  if (body.fin !== undefined) patch.Fin = body.fin;
  if (body.meta !== undefined) patch.Meta = body.meta;
  updateRowById_(ss, TABS.SPRINTS, 'ID', body.sprintId, patch);
  return { ok: true };
}

// ============================================================
// ARCHIVOS (con control de permisos por responsable)
// ============================================================
// Encuentra el módulo (tarea de nivel superior, sin ParentID) al que
// pertenece una tarea/subtarea, subiendo por la cadena de padres.
function encontrarModulo_(tareasFlat, taskId) {
  var byId = {};
  tareasFlat.forEach(function (t) { byId[t.ID] = t; });
  var actual = byId[taskId];
  if (!actual) return null;
  var visitados = {};
  while (actual.ParentID && byId[actual.ParentID] && !visitados[actual.ID]) {
    visitados[actual.ID] = true;
    actual = byId[actual.ParentID];
  }
  return actual;
}

// Estructura de carpetas: <Proyecto>/<Módulo>/<Usuario que sube>/archivo
// (antes todos los archivos de una persona iban juntos en una sola carpeta
// sin separar por módulo/tarea).
function getModuleUserFolder_(projectFolder, tareasFlat, taskId, usuario) {
  var modulo = encontrarModulo_(tareasFlat, taskId);
  var moduloFolder = getOrCreateSubfolder_(projectFolder, (modulo && modulo.Titulo) ? modulo.Titulo : 'General');
  return getOrCreateSubfolder_(moduloFolder, usuario || 'General');
}

function uploadFile(body) {
  var folder = getProjectFolder_(body.projectId);
  var ss = getControlSheet_(folder);
  var tareas = readTable_(ss, TABS.TAREAS);
  var tarea = tareas.filter(function (t) { return String(t.ID) === String(body.taskId); })[0];
  if (!tarea) return { ok: false, error: 'Tarea no encontrada.' };

  var usuarios = readTable_(ss, TABS.USUARIOS);
  var u = usuarios.filter(function (x) { return String(x.Nombre).toLowerCase() === String(body.usuario || '').toLowerCase(); })[0];
  if (!u) return { ok: false, error: 'Usuario no válido.' };

  var esResponsable = esUnoDeLosResponsables_(tarea.Responsable, body.usuario);
  if (u.Rol !== 'admin' && !esResponsable) {
    return { ok: false, error: 'No autorizado: solo ' + (leerResponsables_(tarea.Responsable).join(' / ') || 'el responsable asignado') + ' puede subir archivos a esta tarea.' };
  }

  // El archivo se organiza por módulo (tarea principal) y, dentro de este,
  // por la persona que lo sube: <Módulo>/<Usuario>/archivo.
  var subfolder = getModuleUserFolder_(folder, tareas, body.taskId, body.usuario);
  var bytes = Utilities.base64Decode(body.base64);
  var blob = Utilities.newBlob(bytes, body.mimeType || 'application/octet-stream', body.filename || 'archivo');
  var file = subfolder.createFile(blob);

  var id = Utilities.getUuid();
  var fecha = hoyEcuador_();
  appendRow_(ss, TABS.ARCHIVOS, { ID: id, TareaID: body.taskId, Nombre: file.getName(), DriveFileId: file.getId(), Url: file.getUrl(), SubidoPor: body.usuario, Fecha: fecha });

  return { ok: true, archivo: { id: id, nombre: file.getName(), driveFileId: file.getId(), url: file.getUrl(), subidoPor: body.usuario, fecha: fecha } };
}

function deleteFile(body) {
  var ss = openControlSheet_(body.projectId);
  var archivos = readTable_(ss, TABS.ARCHIVOS);
  var a = archivos.filter(function (x) { return String(x.ID) === String(body.archivoId); })[0];
  if (!a) return { ok: false, error: 'Archivo no encontrado.' };

  var usuarios = readTable_(ss, TABS.USUARIOS);
  var u = usuarios.filter(function (x) { return String(x.Nombre).toLowerCase() === String(body.usuario || '').toLowerCase(); })[0];
  var esDueno = a.SubidoPor && String(a.SubidoPor).toLowerCase() === String(body.usuario || '').toLowerCase();
  if (!u || (u.Rol !== 'admin' && !esDueno)) return { ok: false, error: 'No autorizado para eliminar este archivo.' };

  try { DriveApp.getFileById(a.DriveFileId).setTrashed(true); } catch (e) { /* si ya no existe, continuar */ }
  deleteRowsWhere_(ss, TABS.ARCHIVOS, 'ID', [body.archivoId]);
  return { ok: true };
}
