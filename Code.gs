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

var TABS = {
  CONFIG: 'Config',
  USUARIOS: 'Usuarios',
  TAREAS: 'Tareas',
  SPRINTS: 'Sprints',
  ARCHIVOS: 'Archivos',
  ACCESOS: 'Accesos'
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
      logDocAccess: logDocAccess
    };
    if (!handlers[action]) return jsonResponse({ ok: false, error: 'Acción no reconocida: ' + action });
    return jsonResponse(handlers[action](body));
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
function addToIndex_(id, nombre) {
  var idx = getIndex_();
  if (!idx.some(function (p) { return p.id === id; })) {
    idx.push({ id: id, nombre: nombre });
    saveIndex_(idx);
  }
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

// Lee una pestaña como array de objetos usando la fila 1 como encabezados.
function readTable_(ss, tabName) {
  var sh = ss.getSheetByName(tabName);
  if (!sh || sh.getLastRow() < 2) return [];
  var values = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  return values.map(function (row) {
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
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var idCol = headers.indexOf(idField);
  var data = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][idCol]) === String(idValue)) {
      headers.forEach(function (h, c) { if (patch[h] !== undefined) data[i][c] = patch[h]; });
      sh.getRange(i + 2, 1, 1, headers.length).setValues([data[i]]);
      return true;
    }
  }
  return false;
}

function deleteRowsWhere_(ss, tabName, idField, idValues) {
  var sh = ss.getSheetByName(tabName);
  if (!sh || sh.getLastRow() < 2) return;
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var idCol = headers.indexOf(idField);
  var lastRow = sh.getLastRow();
  for (var r = lastRow; r >= 2; r--) {
    var val = String(sh.getRange(r, idCol + 1).getValue());
    if (idValues.indexOf(val) > -1) sh.deleteRow(r);
  }
}

// Escribe un árbol de tareas {titulo, prioridad, estado, vencimiento, notas,
// cronogramaInicio, cronogramaFin, hijos:[...]} en la pestaña Tareas,
// agregándolo después de las filas que ya existan (no las borra).
function writeSeedTareas_(sh, nodes) {
  var rows = [];
  var hoy = new Date().toISOString().slice(0, 10);
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
      var folder = DriveApp.getFolderById(entry.id);
      var ss = getControlSheet_(folder);
      var config = readTable_(ss, TABS.CONFIG)[0] || {};
      var tareas = readTable_(ss, TABS.TAREAS);
      var total = tareas.length;
      var hechas = tareas.filter(function (t) { return t.Estado === 'Hecho'; }).length;
      var usuarios = readTable_(ss, TABS.USUARIOS).map(function (u) { return { nombre: u.Nombre, rol: u.Rol, email: u.Email || '' }; });
      out.push({
        id: folder.getId(), nombre: config.Nombre || folder.getName(),
        objetivo: config.Objetivo || '', resultados: config.Resultados || '',
        creado: config.Creado || '', total: total, hechas: hechas, usuarios: usuarios
      });
    } catch (err) { /* carpeta eliminada, inaccesible o sin Control válido: se ignora */ }
  });
  return { ok: true, proyectos: out };
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
  ss.getSheetByName(TABS.CONFIG).getRange(2, 1, 1, 4).setValues([[nombre, body.objetivo || '', body.resultados || '', new Date().toISOString().slice(0, 10)]]);

  var shUsuarios = ss.insertSheet(TABS.USUARIOS);
  shUsuarios.getRange(1, 1, 1, 4).setValues([['Nombre', 'Rol', 'PasswordHash', 'Email']]);

  var shTareas = ss.insertSheet(TABS.TAREAS);
  shTareas.getRange(1, 1, 1, 13).setValues([['ID', 'ParentID', 'Titulo', 'Responsable', 'Estado', 'Vencimiento', 'Prioridad', 'Notas', 'CronogramaInicio', 'CronogramaFin', 'SprintID', 'Actualizado', 'EnlacesDocumento']]);

  // Plantilla opcional: si el creador eligió cargar una estructura de tareas
  // ya definida (ej. "PROYECTO DE CAPACITACIONES"), se escribe todo de una vez.
  if (body.seedTareas && body.seedTareas.length) writeSeedTareas_(shTareas, body.seedTareas);

  var shSprints = ss.insertSheet(TABS.SPRINTS);
  shSprints.getRange(1, 1, 1, 5).setValues([['ID', 'Nombre', 'Inicio', 'Fin', 'Meta']]);

  var shArchivos = ss.insertSheet(TABS.ARCHIVOS);
  shArchivos.getRange(1, 1, 1, 7).setValues([['ID', 'TareaID', 'Nombre', 'DriveFileId', 'Url', 'SubidoPor', 'Fecha']]);

  var shAccesos = ss.insertSheet(TABS.ACCESOS);
  shAccesos.getRange(1, 1, 1, 6).setValues([['Fecha', 'TareaID', 'LinkId', 'Documento', 'Usuario', 'Rol']]);

  // fila admin
  appendRow_(ss, TABS.USUARIOS, { Nombre: adminNombre, Rol: 'admin', PasswordHash: hashPassword_(adminPassword), Email: adminEmail });
  getOrCreateSubfolder_(folder, adminNombre);

  // participantes (evita duplicar al admin si aparece también en la lista)
  participantes.forEach(function (p) {
    var nombreP = (p.nombre || '').trim();
    if (!nombreP || nombreP.toLowerCase() === adminNombre.toLowerCase()) return;
    appendRow_(ss, TABS.USUARIOS, { Nombre: nombreP, Rol: p.rol === 'admin' ? 'admin' : 'participante', PasswordHash: p.rol === 'admin' ? hashPassword_(body.adminPassword) : '', Email: (p.email || '').trim() });
    getOrCreateSubfolder_(folder, nombreP);
  });
  getOrCreateSubfolder_(folder, 'General');

  addToIndex_(folder.getId(), nombre);
  return { ok: true, projectId: folder.getId() };
}

function saveConfig(body) {
  var ss = getControlSheet_(getProjectFolder_(body.projectId));
  var sh = ss.getSheetByName(TABS.CONFIG);
  sh.getRange(2, 2, 1, 2).setValues([[body.objetivo || '', body.resultados || '']]);
  if (body.nombre) {
    sh.getRange(2, 1).setValue(body.nombre);
    getProjectFolder_(body.projectId).setName(body.nombre);
  }
  return { ok: true };
}

function getProjectData(projectId) {
  var folder = getProjectFolder_(projectId);
  var ss = getControlSheet_(folder);
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
    byId[t.ID] = {
      id: t.ID, parentId: t.ParentID || null, titulo: t.Titulo, responsable: t.Responsable || '',
      estado: t.Estado || 'No iniciado', vencimiento: t.Vencimiento || '', prioridad: t.Prioridad || 'Media',
      notas: t.Notas || '', cronogramaInicio: t.CronogramaInicio || '', cronogramaFin: t.CronogramaFin || '',
      sprintId: t.SprintID || null, actualizado: t.Actualizado || '', expandido: false,
      enlaces: leerEnlaces_(t),
      archivos: archivosPorTarea[t.ID] || [], hijos: []
    };
  });
  var roots = [];
  tareasFlat.forEach(function (t) {
    var node = byId[t.ID];
    if (t.ParentID && byId[t.ParentID]) byId[t.ParentID].hijos.push(node);
    else roots.push(node);
  });

  return {
    ok: true,
    proyecto: { id: projectId, nombre: config.Nombre || '', objetivo: config.Objetivo || '', resultados: config.Resultados || '', creado: config.Creado || '' },
    usuarios: usuarios, tareas: roots, sprints: sprints
  };
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
  var ss = getControlSheet_(getProjectFolder_(body.projectId));
  var usuarios = readTable_(ss, TABS.USUARIOS);
  var u = usuarios.filter(function (x) { return String(x.Nombre).toLowerCase() === String(body.usuario || '').toLowerCase(); })[0];
  if (!u) return { ok: false, error: 'Usuario no encontrado en este proyecto.' };
  if (u.Rol === 'admin') {
    if (String(u.PasswordHash) !== hashPassword_(body.password)) return { ok: false, error: 'Contraseña incorrecta.' };
  }
  return { ok: true, usuario: u.Nombre, rol: u.Rol };
}

function addParticipant(body) {
  var ss = getControlSheet_(getProjectFolder_(body.projectId));
  var err = requireAdmin_(ss, body.adminUsuario, body.adminPassword);
  if (err) return { ok: false, error: err };
  var nombre = (body.nombre || '').trim();
  var email = (body.email || '').trim(); // opcional: ya no se exige ni se usa para compartir
  if (!nombre) return { ok: false, error: 'Nombre requerido.' };
  var existentes = readTable_(ss, TABS.USUARIOS);
  if (existentes.some(function (u) { return String(u.Nombre).toLowerCase() === nombre.toLowerCase(); })) return { ok: false, error: 'Ese usuario ya existe en el proyecto.' };
  var esAdmin = body.rol === 'admin';
  appendRow_(ss, TABS.USUARIOS, { Nombre: nombre, Rol: esAdmin ? 'admin' : 'participante', PasswordHash: esAdmin ? hashPassword_(body.nuevoPassword || body.adminPassword) : '', Email: email });
  getOrCreateSubfolder_(getProjectFolder_(body.projectId), nombre);
  return { ok: true };
}

// ============================================================
// TAREAS
// ============================================================
function addTask(body) {
  var ss = getControlSheet_(getProjectFolder_(body.projectId));
  var id = Utilities.getUuid();
  appendRow_(ss, TABS.TAREAS, {
    ID: id, ParentID: body.parentId || '', Titulo: body.titulo || 'Nueva tarea', Responsable: '',
    Estado: 'No iniciado', Vencimiento: '', Prioridad: 'Media', Notas: '', CronogramaInicio: '', CronogramaFin: '',
    SprintID: '', Actualizado: new Date().toISOString().slice(0, 10)
  });
  return { ok: true, id: id };
}

function updateTask(body) {
  var ss = getControlSheet_(getProjectFolder_(body.projectId));
  var patch = body.patch || {};
  var fieldMap = { titulo: 'Titulo', responsable: 'Responsable', estado: 'Estado', vencimiento: 'Vencimiento', prioridad: 'Prioridad', notas: 'Notas', cronogramaInicio: 'CronogramaInicio', cronogramaFin: 'CronogramaFin', sprintId: 'SprintID' };
  var row = {};
  Object.keys(patch).forEach(function (k) { if (fieldMap[k]) row[fieldMap[k]] = patch[k]; });
  row.Actualizado = new Date().toISOString().slice(0, 10);
  var okUpd = updateRowById_(ss, TABS.TAREAS, 'ID', body.taskId, row);
  return { ok: okUpd };
}

// Importa una plantilla de tareas (árbol) a un proyecto YA EXISTENTE,
// sin borrar las tareas que ya tenga. Solo el admin del proyecto puede hacerlo.
function importTareas(body) {
  var ss = getControlSheet_(getProjectFolder_(body.projectId));
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
function getDocsFolder_(projectFolder) {
  return getOrCreateSubfolder_(projectFolder, 'Documentos en línea');
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
  var ss = getControlSheet_(getProjectFolder_(body.projectId));
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
  var ss = getControlSheet_(getProjectFolder_(body.projectId));
  var err = requireAdmin_(ss, body.adminUsuario, body.adminPassword);
  if (err) return { ok: false, error: err };
  var row = encontrarTarea_(ss, body.taskId);
  if (!row) return { ok: false, error: 'Tarea no encontrada.' };
  var titulo = (body.titulo || '').trim();
  if (!titulo) return { ok: false, error: 'El nombre del documento es obligatorio.' };

  var file, url;
  if (body.tipo === 'doc') { var d = DocumentApp.create(titulo); file = DriveApp.getFileById(d.getId()); url = d.getUrl(); }
  else if (body.tipo === 'sheet') { var s = SpreadsheetApp.create(titulo); file = DriveApp.getFileById(s.getId()); url = s.getUrl(); }
  else if (body.tipo === 'slide') { var p = SlidesApp.create(titulo); file = DriveApp.getFileById(p.getId()); url = p.getUrl(); }
  else return { ok: false, error: 'Tipo de documento no válido.' };

  var docsFolder = getDocsFolder_(getProjectFolder_(body.projectId));
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
  var ss = getControlSheet_(getProjectFolder_(body.projectId));
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

function logDocAccess(body) {
  var ss = getControlSheet_(getProjectFolder_(body.projectId));
  var sh = ensureAccesosSheet_(ss);
  sh.appendRow([new Date().toISOString(), body.taskId || '', body.linkId || '', body.documento || '', body.usuario || '', body.rol || '']);
  return { ok: true };
}

function getDocAccessLog(p) {
  var ss = getControlSheet_(getProjectFolder_(p.projectId));
  ensureAccesosSheet_(ss);
  var rows = readTable_(ss, TABS.ACCESOS).filter(function (r) { return String(r.LinkId) === String(p.linkId); });
  rows.sort(function (a, b) { return new Date(b.Fecha) - new Date(a.Fecha); });
  return {
    ok: true,
    accesos: rows.slice(0, 10).map(function (r) { return { fecha: r.Fecha, usuario: r.Usuario, rol: r.Rol }; })
  };
}

function deleteTask(body) {
  var ss = getControlSheet_(getProjectFolder_(body.projectId));
  var tareas = readTable_(ss, TABS.TAREAS);
  var ids = collectDescendants_(tareas, body.taskId);
  deleteRowsWhere_(ss, TABS.TAREAS, 'ID', ids);
  deleteRowsWhere_(ss, TABS.ARCHIVOS, 'TareaID', ids);
  return { ok: true };
}

// ============================================================
// SPRINTS
// ============================================================
function addSprint(body) {
  var ss = getControlSheet_(getProjectFolder_(body.projectId));
  var id = Utilities.getUuid();
  appendRow_(ss, TABS.SPRINTS, { ID: id, Nombre: body.nombre || 'Sprint', Inicio: body.inicio || '', Fin: body.fin || '', Meta: body.meta || '' });
  return { ok: true, id: id };
}

function updateSprint(body) {
  var ss = getControlSheet_(getProjectFolder_(body.projectId));
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
function uploadFile(body) {
  var folder = getProjectFolder_(body.projectId);
  var ss = getControlSheet_(folder);
  var tareas = readTable_(ss, TABS.TAREAS);
  var tarea = tareas.filter(function (t) { return String(t.ID) === String(body.taskId); })[0];
  if (!tarea) return { ok: false, error: 'Tarea no encontrada.' };

  var usuarios = readTable_(ss, TABS.USUARIOS);
  var u = usuarios.filter(function (x) { return String(x.Nombre).toLowerCase() === String(body.usuario || '').toLowerCase(); })[0];
  if (!u) return { ok: false, error: 'Usuario no válido.' };

  var esResponsable = tarea.Responsable && String(tarea.Responsable).trim().toLowerCase() === String(body.usuario).trim().toLowerCase();
  if (u.Rol !== 'admin' && !esResponsable) {
    return { ok: false, error: 'No autorizado: solo "' + (tarea.Responsable || 'el responsable asignado') + '" puede subir archivos a esta tarea.' };
  }

  var subfolderName = tarea.Responsable ? tarea.Responsable : 'General';
  var subfolder = getOrCreateSubfolder_(folder, subfolderName);
  var bytes = Utilities.base64Decode(body.base64);
  var blob = Utilities.newBlob(bytes, body.mimeType || 'application/octet-stream', body.filename || 'archivo');
  var file = subfolder.createFile(blob);

  var id = Utilities.getUuid();
  var fecha = new Date().toISOString().slice(0, 10);
  appendRow_(ss, TABS.ARCHIVOS, { ID: id, TareaID: body.taskId, Nombre: file.getName(), DriveFileId: file.getId(), Url: file.getUrl(), SubidoPor: body.usuario, Fecha: fecha });

  return { ok: true, archivo: { id: id, nombre: file.getName(), driveFileId: file.getId(), url: file.getUrl(), subidoPor: body.usuario, fecha: fecha } };
}

function deleteFile(body) {
  var ss = getControlSheet_(getProjectFolder_(body.projectId));
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
