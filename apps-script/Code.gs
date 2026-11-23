/**
 * CARNETS CTPL — Colegio Técnico Profesional de Liberia
 * Registro de personal (foto + datos) y administración de carnets (impresión A4).
 * Las páginas están en GitHub Pages; este proyecto es el servidor de datos.
 *
 * Hojas que usa:  "Personal"  (se crea con setup())
 * Carpeta Drive:  "CARNETS CTPL - FOTOS" (se crea con setup())
 */

const HOJA = 'Personal';
const COLS = ['Fecha registro', 'Cédula', 'Nombre completo', 'Puesto', 'Género', 'Foto (ID Drive)', 'Estado', 'Fecha impresión', 'QR'];
const PUESTOS = ['DOCENTE_TECNICO', 'DOCENTE_ACADEMICO', 'DOCENTE_IDIOMAS', 'ADMINISTRATIVO', 'AUXILIAR_ADMIN', 'COORD_TECNICO', 'COORD_EMPRESA', 'COORD_ACADEMICO', 'SUBDIRECTOR', 'DIRECTOR'];
const CLAVE_INICIAL = 'CTPL2027';

/* ===================== SERVIDOR ===================== */
/* Las páginas viven en GitHub Pages. Este proyecto solo atiende datos (ver Api.gs). */
function doGet() {
  return ContentService.createTextOutput(JSON.stringify({ ok: true, servicio: 'CTPL Carnets y Asistencia', hora: new Date().toISOString() }))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ===================== CONFIGURACIÓN (ejecutar 1 vez) ===================== */

function setup() {
  const props = PropertiesService.getScriptProperties();
  let ss;
  try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (err) { ss = null; }
  if (!ss) ss = SpreadsheetApp.create('CARNETS CTPL - Base de datos');
  props.setProperty('SS_ID', ss.getId());

  let sh = ss.getSheetByName(HOJA);
  if (!sh) sh = ss.insertSheet(HOJA);
  sh.getRange(1, 1, 1, COLS.length).setValues([COLS])
    .setFontWeight('bold').setBackground('#1f2f5c').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  sh.getRange('B:B').setNumberFormat('@');
  sh.setColumnWidth(3, 280);

  if (!props.getProperty('FOLDER_ID')) {
    const f = DriveApp.createFolder('CARNETS CTPL - FOTOS');
    props.setProperty('FOLDER_ID', f.getId());
  }
  if (!props.getProperty('ADMIN_PASS')) props.setProperty('ADMIN_PASS', CLAVE_INICIAL);
  if (!props.getProperty('REGISTRO_ABIERTO')) props.setProperty('REGISTRO_ABIERTO', 'SI');

  Logger.log('Listo. Hoja: ' + ss.getUrl());
  Logger.log('Carpeta de fotos: ' + DriveApp.getFolderById(props.getProperty('FOLDER_ID')).getUrl());
  Logger.log('Clave admin inicial: ' + props.getProperty('ADMIN_PASS') + '  (cámbiela desde el panel)');
}

/* ===================== UTILIDADES ===================== */

/* Abre la hoja de cálculo una sola vez por ejecución (cada apertura cuesta tiempo) */
let LIBRO_ = null;
const HOJAS_ = {};
function libro_() {
  if (!LIBRO_) {
    const id = PropertiesService.getScriptProperties().getProperty('SS_ID');
    if (!id) throw new Error('El sistema no está configurado. Ejecute setup() en el editor.');
    LIBRO_ = SpreadsheetApp.openById(id);
  }
  return LIBRO_;
}

function hoja_() {
  const id = PropertiesService.getScriptProperties().getProperty('SS_ID');
  if (!id) throw new Error('El sistema no está configurado. Ejecute setup() en el editor.');
  return libro_().getSheetByName(HOJA);
}

function carpeta_() {
  return DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty('FOLDER_ID'));
}

function limpiarCedula_(c) {
  return String(c || '').toUpperCase().replace(/[^0-9A-Z]/g, '');
}

function limpiarNombre_(n) {
  return String(n || '').replace(/\s+/g, ' ').trim().toUpperCase();
}

function filaPorCedula_(sh, ced) {
  const n = sh.getLastRow();
  if (n < 2) return -1;
  const vals = sh.getRange(2, 2, n - 1, 1).getDisplayValues();
  for (let i = 0; i < vals.length; i++) if (limpiarCedula_(vals[i][0]) === ced) return i + 2;
  return -1;
}

function guardarFoto_(dataUrl, ced, fotoAnterior) {
  const m = String(dataUrl || '').match(/^data:(image\/[a-z]+);base64,(.+)$/);
  if (!m) throw new Error('La fotografía no es válida.');
  const blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], ced + '.jpg');
  if (blob.getBytes().length > 3 * 1024 * 1024) throw new Error('La foto es demasiado pesada.');
  const file = carpeta_().createFile(blob);
  if (fotoAnterior) { try { DriveApp.getFileById(fotoAnterior).setTrashed(true); } catch (e) {} }
  return file.getId();
}

function fotoDataUrl_(id) {
  if (!id) return '';
  try {
    const b = DriveApp.getFileById(id).getBlob();
    return 'data:' + b.getContentType() + ';base64,' + Utilities.base64Encode(b.getBytes());
  } catch (e) { return ''; }
}

/* ===================== REGISTRO (público) ===================== */

function estadoRegistro() {
  return PropertiesService.getScriptProperties().getProperty('REGISTRO_ABIERTO') !== 'NO';
}

function registrar(d) {
  if (!estadoRegistro()) throw new Error('El registro se encuentra cerrado. Consulte con la administración.');
  const ced = limpiarCedula_(d.cedula);
  const nombre = limpiarNombre_(d.nombre);
  if (ced.length < 9 || ced.length > 12) throw new Error('La cédula debe tener entre 9 y 12 caracteres, sin guiones.');
  if (nombre.split(' ').length < 2) throw new Error('Escriba su nombre completo (nombre y apellidos).');
  if (PUESTOS.indexOf(d.puesto) < 0) throw new Error('Seleccione un puesto válido.');
  const genero = d.genero === 'F' ? 'F' : 'M';

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = hoja_();
    const fila = filaPorCedula_(sh, ced);
    const anterior = fila > 0 ? sh.getRange(fila, 1, 1, COLS.length).getValues()[0] : null;
    const fotoId = guardarFoto_(d.foto, ced, anterior ? anterior[5] : '');
    const reg = [new Date(), ced, nombre, d.puesto, genero, fotoId, 'PENDIENTE', '', anterior ? anterior[8] : ''];
    if (fila > 0) sh.getRange(fila, 1, 1, COLS.length).setValues([reg]);
    else sh.appendRow(reg);
    sh.getRange(fila > 0 ? fila : sh.getLastRow(), 2).setNumberFormat('@').setValue(ced);
    limpiarCachePersonal_();
    return { ok: true, actualizado: fila > 0, cedula: ced, nombre: nombre };
  } finally {
    lock.releaseLock();
  }
}

/* ===================== ADMINISTRACIÓN ===================== */

function verificar_(token) {
  if (!token || CacheService.getScriptCache().get('adm_' + token) !== '1') {
    throw new Error('SESION_EXPIRADA');
  }
}

function adminLogin(clave) {
  const real = PropertiesService.getScriptProperties().getProperty('ADMIN_PASS') || CLAVE_INICIAL;
  if (String(clave) !== real) throw new Error('Clave incorrecta.');
  const token = Utilities.getUuid();
  CacheService.getScriptCache().put('adm_' + token, '1', 21600); // 6 horas
  return token;
}

function adminListar(token) {
  verificar_(token);
  const sh = hoja_();
  const n = sh.getLastRow();
  if (n < 2) return { registros: [], abierto: estadoRegistro() };
  const vals = sh.getRange(2, 1, n - 1, COLS.length).getValues();
  const tz = Session.getScriptTimeZone();
  const fmt = v => (v instanceof Date) ? Utilities.formatDate(v, tz, 'dd/MM/yyyy HH:mm') : String(v || '');
  const registros = vals.filter(r => r[1]).map(r => ({
    fecha: fmt(r[0]), cedula: limpiarCedula_(r[1]), nombre: String(r[2]), puesto: String(r[3]),
    genero: String(r[4] || 'M'), fotoId: String(r[5]), estado: String(r[6] || 'PENDIENTE'),
    impreso: fmt(r[7]), qr: String(r[8] || '')
  }));
  return { registros: registros, abierto: estadoRegistro() };
}

function adminFotos(token, ids) {
  verificar_(token);
  const out = {};
  (ids || []).slice(0, 25).forEach(id => { out[id] = fotoDataUrl_(id); });
  return out;
}

function adminGuardar(token, d) {
  verificar_(token);
  const ced = limpiarCedula_(d.cedulaOriginal || d.cedula);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = hoja_();
    const fila = filaPorCedula_(sh, ced);
    if (fila < 0) throw new Error('Registro no encontrado.');
    const r = sh.getRange(fila, 1, 1, COLS.length).getValues()[0];
    const nuevaCed = limpiarCedula_(d.cedula);
    if (nuevaCed !== ced && filaPorCedula_(sh, nuevaCed) > 0) throw new Error('Ya existe otro registro con esa cédula.');
    r[1] = nuevaCed; r[2] = limpiarNombre_(d.nombre);
    if (PUESTOS.indexOf(d.puesto) >= 0) r[3] = d.puesto;
    r[4] = d.genero === 'F' ? 'F' : 'M';
    if (d.foto) r[5] = guardarFoto_(d.foto, nuevaCed, r[5]);
    sh.getRange(fila, 1, 1, COLS.length).setValues([r]);
    sh.getRange(fila, 2).setNumberFormat('@').setValue(nuevaCed);
    limpiarCachePersonal_();
    return { ok: true, fotoId: r[5] };
  } finally { lock.releaseLock(); }
}

function adminEliminar(token, cedula) {
  verificar_(token);
  const sh = hoja_();
  const fila = filaPorCedula_(sh, limpiarCedula_(cedula));
  if (fila < 0) throw new Error('Registro no encontrado.');
  const foto = sh.getRange(fila, 6).getValue();
  if (foto) { try { DriveApp.getFileById(foto).setTrashed(true); } catch (e) {} }
  sh.deleteRow(fila);
  limpiarCachePersonal_();
  return true;
}

function adminMarcar(token, cedulas, estado) {
  verificar_(token);
  const sh = hoja_();
  const n = sh.getLastRow();
  if (n < 2) return 0;
  const set = {};
  (cedulas || []).forEach(c => set[limpiarCedula_(c)] = true);
  const rng = sh.getRange(2, 2, n - 1, 7);
  const vals = rng.getValues();
  let k = 0;
  vals.forEach(r => {
    if (set[limpiarCedula_(r[0])]) {
      r[5] = estado === 'IMPRESO' ? 'IMPRESO' : 'PENDIENTE';
      r[6] = estado === 'IMPRESO' ? new Date() : '';
      k++;
    }
  });
  // solo escribir columnas Estado (G) y Fecha impresión (H)
  sh.getRange(2, 7, vals.length, 2).setValues(vals.map(r => [r[5], r[6]]));
  return k;
}

function adminRegistroAbierto(token, abierto) {
  verificar_(token);
  PropertiesService.getScriptProperties().setProperty('REGISTRO_ABIERTO', abierto ? 'SI' : 'NO');
  return abierto;
}

function adminCambiarClave(token, nueva) {
  verificar_(token);
  if (!nueva || String(nueva).length < 6) throw new Error('La clave debe tener al menos 6 caracteres.');
  PropertiesService.getScriptProperties().setProperty('ADMIN_PASS', String(nueva));
  return true;
}

/* ===================== PLANTILLA DEL CARNET ===================== */

function adminSubirPlantilla(token, dataUrl) {
  verificar_(token);
  const m = String(dataUrl || '').match(/^data:(image\/[a-z]+);base64,(.+)$/);
  if (!m) throw new Error('Archivo no válido. Use JPG o PNG.');
  const props = PropertiesService.getScriptProperties();
  const viejo = props.getProperty('PLANTILLA_ID');
  const f = carpeta_().createFile(Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], '_PLANTILLA_CARNET'));
  f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  props.setProperty('PLANTILLA_ID', f.getId());
  if (viejo) { try { DriveApp.getFileById(viejo).setTrashed(true); } catch (e) {} }
  return true;
}

/** Devuelve el enlace de la imagen de fondo (archivo de Drive compartido para lectura). */
function obtenerPlantilla() {
  const props = PropertiesService.getScriptProperties(), id = props.getProperty('PLANTILLA_ID');
  if (!id) return '';
  if (props.getProperty('PLANTILLA_PUB') !== id) {   // plantillas subidas antes de este cambio
    try { DriveApp.getFileById(id).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); props.setProperty('PLANTILLA_PUB', id); } catch (e) {}
  }
  return 'https://lh3.googleusercontent.com/d/' + id;
}
