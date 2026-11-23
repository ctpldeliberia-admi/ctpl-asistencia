/**
 * CTPL ASISTENCIA — Control de entrada y salida del personal con QR dinámico
 * Usa la misma hoja de cálculo del sistema de carnets (hoja "Personal").
 *
 * Seguridad:
 *  - La pantalla de marcas muestra un QR firmado que cambia cada N segundos (por defecto 15).
 *  - Cada funcionario marca desde SU teléfono. La cuenta queda ligada a un único dispositivo
 *    mediante una llave criptográfica no exportable guardada en el navegador.
 *  - Cambiar de dispositivo requiere aprobación del administrador.
 *  - El correo de cada funcionario debe ser aprobado por el administrador antes de crear su contraseña.
 *
 * Ejecutar UNA vez: setupAsistencia()
 */

const TZ = 'America/Costa_Rica';
const H_CUENTAS = 'Cuentas', H_MARCAS = 'Marcas', H_KIOSCOS = 'Pantallas', H_INFORMES = 'Informes', H_AUDIT = 'Bitácora';

const C_CUENTAS = ['Cédula', 'Nombre', 'Correo', 'Estado', 'Hash', 'Sal', 'Dispositivo ID', 'Llave dispositivo', 'Dispositivo',
  'Vinculado desde', 'Disp. pendiente ID', 'Disp. pendiente', 'Disp. pendiente estado', 'Código', 'Código vence', 'Horario',
  'Fecha solicitud', 'Último acceso', 'Último cambio automático'];
const CU = { CED: 0, NOM: 1, COR: 2, EST: 3, HASH: 4, SAL: 5, DID: 6, DKEY: 7, DINF: 8, DFEC: 9, PID: 10, PINF: 11, PEST: 12,
  COD: 13, CVEN: 14, HOR: 15, FSOL: 16, ULT: 17, AUTO: 18 };

const C_MARCAS = ['ID', 'Fecha y hora', 'Fecha', 'Hora', 'Cédula', 'Nombre', 'Tipo', 'Pantalla', 'Latitud', 'Longitud',
  'Precisión (m)', 'Método', 'Estado', 'Observación', 'Registrado por'];

const C_KIOSCOS = ['ID', 'Nombre', 'Llave', 'Estado', 'Creada', 'Último pulso'];
const C_INFORMES = ['Semana (lunes)', 'Cédula', 'Nombre', 'Puesto', 'Días programados', 'Días asistidos', 'Horas (hh:mm)',
  'Tardías', 'Salidas anticipadas', 'Ausencias', 'Marcas incompletas', 'Detalle', 'Generado'];

const CFG_DEF = {
  institucion: 'Colegio Técnico Profesional de Liberia',
  entrada: '07:00', salida: '16:20', tolerancia: 10,  // minutos de gracia para tardía
  slot: 15,                // segundos que dura cada QR
  separacion: 3,           // minutos mínimos entre dos marcas de la misma persona
  correos: '',             // correos de la administración (separados por coma) que reciben el informe semanal
  geo: false, lat: 0, lng: 0, radio: 250,   // geocerca opcional
  noLectivos: '',          // fechas AAAA-MM-DD separadas por coma
  autoDias: 30,            // el funcionario puede cambiar de teléfono por sí mismo (con código al correo) 1 vez cada N días; 0 = desactivado
  ventanaHasta: 0,         // vinculación libre de dispositivos hasta esta marca de tiempo (ms)
  urlApp: ''               // dirección de la app en GitHub Pages (para los botones de los correos)
};

/* ===================== CONFIGURACIÓN INICIAL ===================== */

function setupAsistencia() {
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('SS_ID')) throw new Error('Primero ejecute setup() del sistema de carnets.');
  const ss = ss_();
  const crear = (nombre, cols, texto) => {
    let sh = ss.getSheetByName(nombre);
    if (!sh) sh = ss.insertSheet(nombre);
    sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold').setBackground('#3a475c').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    (texto || []).forEach(c => sh.getRange(1, c, sh.getMaxRows(), 1).setNumberFormat('@'));
    return sh;
  };
  crear(H_CUENTAS, C_CUENTAS, [1, 7, 11, 14, 15, 16, 19]);
  crear(H_MARCAS, C_MARCAS, [1, 3, 4, 5]);
  crear(H_KIOSCOS, C_KIOSCOS, [1]);
  crear(H_INFORMES, C_INFORMES, [1, 2, 7]);
  crear(H_AUDIT, ['Fecha', 'Actor', 'Acción', 'Detalle']);

  if (!props.getProperty('ASIS_SECRET')) props.setProperty('ASIS_SECRET', nuevaLlave_());
  if (!props.getProperty('ASIS_CFG')) props.setProperty('ASIS_CFG', JSON.stringify(CFG_DEF));

  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'informeSemanalAuto')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'mantenerActivo').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('mantenerActivo').timeBased().everyMinutes(5).create();
  ScriptApp.newTrigger('informeSemanalAuto').timeBased().onWeekDay(ScriptApp.WeekDay.FRIDAY).atHour(22).inTimezone(TZ).create();

  MailApp.getRemainingDailyQuota(); // solicita el permiso de correo
  Logger.log('Asistencia lista. Informe automático: viernes 10:00 p. m.');
}

/* ===================== UTILIDADES ===================== */

function ss_() { return libro_(); }
function sh_(n) {
  if (!HOJAS_[n]) {
    HOJAS_[n] = ss_().getSheetByName(n);
    if (!HOJAS_[n]) throw new Error('Falta la hoja "' + n + '". Ejecute setupAsistencia().');
  }
  return HOJAS_[n];
}
let CFG_MEM_ = null;
function cfg_() {
  if (CFG_MEM_) return Object.assign({}, CFG_MEM_);
  let c = {};
  try { c = JSON.parse(PropertiesService.getScriptProperties().getProperty('ASIS_CFG') || '{}'); } catch (e) {}
  CFG_MEM_ = Object.assign({}, CFG_DEF, c);
  return Object.assign({}, CFG_MEM_);
}
function hoy_(d) { return Utilities.formatDate(d || new Date(), TZ, 'yyyy-MM-dd'); }
function hora_(d) { return Utilities.formatDate(d || new Date(), TZ, 'HH:mm:ss'); }
function min_(h) { const p = String(h).split(':'); return (+p[0]) * 60 + (+p[1] || 0) + (+(p[2] || 0)) / 60; }
function fechaD_(f) { return new Date(f + 'T12:00:00-06:00'); }
function diaSemana_(f) { const w = fechaD_(f).getUTCDay(); return w === 0 ? 7 : w; }
function sumarDias_(f, n) { const d = fechaD_(f); d.setUTCDate(d.getUTCDate() + n); return Utilities.formatDate(d, TZ, 'yyyy-MM-dd'); }
function lunesDe_(f) { return sumarDias_(f, 1 - diaSemana_(f)); }
function hhmm_(horas) { const m = Math.round(horas * 60); return Math.floor(m / 60) + ':' + ('0' + (m % 60)).slice(-2); }
function corto_(n) { const p = String(n).split(' '); return p.length > 2 ? p[0] + ' ' + p[p.length - 2] : n; }

function nuevaLlave_() {
  return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,
    Utilities.getUuid() + Utilities.getUuid() + Math.random() + Date.now()));
}
function firma_(llaveB64, msg) {
  return Utilities.base64Encode(Utilities.computeHmacSha256Signature(Utilities.newBlob(msg).getBytes(), Utilities.base64Decode(llaveB64)));
}
function hashClave_(clave, sal) {
  let h = sal + '|' + clave, b;
  for (let i = 0; i < 400; i++) { b = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, h); h = Utilities.base64Encode(b) + clave; }
  return Utilities.base64Encode(b);
}
function validarClave_(c) {
  c = String(c || '');
  if (c.length < 8 || !/[A-Za-z]/.test(c) || !/[0-9]/.test(c)) throw new Error('La contraseña debe tener al menos 8 caracteres, con letras y números.');
  return c;
}
function codigo6_() { return String(Math.floor(100000 + Math.random() * 900000)); }
function correoValido_(c) {
  c = String(c || '').trim().toLowerCase();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(c)) throw new Error('Correo electrónico no válido.');
  return c;
}
function ocultarCorreo_(c) { const p = String(c).split('@'); return p[0].slice(0, 2) + '•••@' + (p[1] || ''); }
function audit_(actor, accion, det) {
  try { sh_(H_AUDIT).appendRow([new Date(), actor, accion, det || '']); } catch (e) {}
}
function bloquear_(fn) {
  const l = LockService.getScriptLock(); l.waitLock(25000);
  try { return fn(); } finally { l.releaseLock(); }
}

/* ---------- Personal (hoja de carnets) ---------- */
let PERSONAL_MEM_ = null;
function mapaPersonal_() {
  if (PERSONAL_MEM_) return PERSONAL_MEM_;
  const cache = CacheService.getScriptCache(), c = cache.get('personal');
  if (c) { try { return PERSONAL_MEM_ = JSON.parse(c); } catch (e) {} }
  const sh = hoja_(), n = sh.getLastRow(), m = {};
  if (n < 2) return m;
  sh.getRange(2, 1, n - 1, 6).getDisplayValues().forEach(r => {
    const c = limpiarCedula_(r[1]);
    if (c) m[c] = { cedula: c, nombre: r[2], puesto: r[3], genero: r[4] || 'M', fotoId: r[5] };
  });
  try { const j = JSON.stringify(m); if (j.length < 95000) cache.put('personal', j, 600); } catch (e) {}
  return PERSONAL_MEM_ = m;
}
function limpiarCachePersonal_() { PERSONAL_MEM_ = null; CacheService.getScriptCache().remove('personal'); }
function personal_(ced) { return mapaPersonal_()[ced] || null; }

/* ---------- Cuentas ---------- */
function cuentas_() {
  const sh = sh_(H_CUENTAS), n = sh.getLastRow();
  return { sh: sh, v: n < 2 ? [] : sh.getRange(2, 1, n - 1, C_CUENTAS.length).getDisplayValues() };
}
function buscarCuenta_(ced) {
  const t = cuentas_();
  for (let i = 0; i < t.v.length; i++) if (limpiarCedula_(t.v[i][CU.CED]) === ced) return { sh: t.sh, fila: i + 2, r: t.v[i], todas: t.v };
  return { sh: t.sh, fila: -1, r: null, todas: t.v };
}
function guardarCuenta_(c) { c.sh.getRange(c.fila, 1, 1, C_CUENTAS.length).setValues([c.r]); }
function correoEnUso_(todas, correo, ced) {
  return todas.some(r => String(r[CU.COR]).toLowerCase() === correo && limpiarCedula_(r[CU.CED]) !== ced);
}
/* ---------- Horarios por tramos ----------
   Cada día puede ser:  "-" (no labora) · vacío (jornada general) · "07:00-12:10" · "07:00-12:10,17:30-21:45" (día y noche) */
function dosDig_(h) { const p = String(h).split(':'); return ('0' + (+p[0])).slice(-2) + ':' + ('0' + (+p[1])).slice(-2); }
function tramos_(v, cfg) {
  v = String(v == null ? '' : v).replace(/\s/g, '');
  if (v === '-') return [];
  if (!v) return [{ ent: cfg.entrada, sal: cfg.salida }];
  return v.split(/[,;y]/).filter(Boolean).map(x => {
    const m = x.match(/^(\d{1,2}:\d{2})-(\d{1,2}:\d{2})$/);
    if (!m) throw new Error('Formato de horario no válido: ' + x);
    return { ent: dosDig_(m[1]), sal: dosDig_(m[2]) };
  }).sort((a, b) => min_(a.ent) - min_(b.ent));
}
function horarioDia_(r, f, cfg) {
  const d = diaSemana_(f);
  if (d > 5) return null;
  let h = {};
  try { h = JSON.parse(r[CU.HOR] || '{}'); } catch (e) {}
  let tr;
  try { tr = tramos_(h[d], cfg); } catch (e) { tr = tramos_('', cfg); }
  if (!tr.length) return null;
  return { tramos: tr, ent: tr[0].ent, sal: tr[tr.length - 1].sal, txt: tr.map(t => t.ent + '-' + t.sal).join(' y ') };
}
/** Tramo al que pertenece una hora (el que la contiene, o el más cercano). */
function tramoDe_(tr, hora) {
  const x = min_(hora); let mejor = 0, dist = 1e9;
  tr.forEach((t, i) => { const a = min_(t.ent), b = min_(t.sal), dd = x < a ? a - x : x > b ? x - b : 0; if (dd < dist) { dist = dd; mejor = i; } });
  return mejor;
}
/** Evalúa un día con marcas ordenadas contra su horario: tardías, salidas anticipadas y jornadas ausentes. */
function evaluarDia_(h, md, cerrado, ahoraMin, tol) {
  const r = { tard: 0, salAnt: 0, aus: 0, obs: [] }, varios = h.tramos.length > 1;
  const grupos = h.tramos.map(() => []);
  md.forEach(m => grupos[tramoDe_(h.tramos, m.hora)].push(m));
  h.tramos.forEach((t, i) => {
    const g = grupos[i], et = varios ? ' (' + t.ent + '-' + t.sal + ')' : '';
    if (!g.length) {
      if (cerrado || ahoraMin > min_(t.sal)) { r.aus++; r.obs.push('Ausente' + et); }
      return;
    }
    const pe = g.find(m => m.tipo === 'ENTRADA'), ul = g[g.length - 1];
    if (pe && min_(pe.hora) > min_(t.ent) + tol) { r.tard++; r.obs.push('Tardía' + et); }
    if (ul.tipo === 'SALIDA' && min_(ul.hora) < min_(t.sal)) { r.salAnt++; r.obs.push('Salida anticipada' + et); }
  });
  return r;
}

/** Verifica que la solicitud venga del dispositivo vinculado (firma HMAC con llave no exportable). */
function verificarDisp_(a, accion, admitePendiente) {
  if (!a || !a.ced || !a.dev || !a.ts || !a.sig) throw new Error('DISPOSITIVO_NO_AUTORIZADO');
  if (Math.abs(Date.now() - Number(a.ts)) > 5 * 60 * 1000) throw new Error('Solicitud vencida. Revise que la fecha y hora del teléfono sean automáticas.');
  const c = buscarCuenta_(limpiarCedula_(a.ced));
  if (!c.r) throw new Error('DISPOSITIVO_NO_AUTORIZADO');
  if (c.r[CU.EST] === 'BLOQUEADA') throw new Error('Su cuenta está bloqueada. Consulte con la administración.');
  const estOk = c.r[CU.EST] === 'ACTIVA' || (admitePendiente && c.r[CU.EST] === 'PENDIENTE');
  if (!estOk || c.r[CU.DID] !== a.dev || !c.r[CU.DKEY]) throw new Error('DISPOSITIVO_NO_AUTORIZADO');
  if (firma_(c.r[CU.DKEY], [c.r[CU.CED], a.dev, a.ts, accion].join('|')) !== a.sig) throw new Error('DISPOSITIVO_NO_AUTORIZADO');
  return c;
}

/** ¿Puede cambiar de teléfono por sí mismo? (límite de días configurable) */
function puedeAuto_(r, cfg) {
  const dias = +cfg.autoDias || 0;
  if (!dias || !r[CU.COR]) return false;
  const ult = Number(r[CU.AUTO] || 0);
  return !ult || Date.now() - ult > dias * 86400e3;
}

function enviarCorreo_(para, asunto, html, adjuntos) {
  try {
    MailApp.sendEmail({ to: para, subject: asunto, name: 'CTPL Asistencia', attachments: adjuntos || [],
      htmlBody: '<div style="font-family:Arial,sans-serif;color:#1f2836;max-width:560px">' +
        '<div style="background:#3a475c;color:#fff;padding:16px 20px;border-radius:10px 10px 0 0"><b>CTPL Asistencia</b><br>' +
        '<span style="font-size:12px;opacity:.8">' + cfg_().institucion + '</span></div>' +
        '<div style="border:1px solid #d5dae2;border-top:0;padding:20px;border-radius:0 0 10px 10px">' + html + '</div></div>' });
    return true;
  } catch (e) { return false; }
}
function correoCodigo_(correo, nombre, codigo, motivo, vigencia) {
  const url = cfg_().urlApp;
  return enviarCorreo_(correo, 'Su código de ' + motivo + ' — CTPL Asistencia',
    '<p>Hola, ' + nombre + '.</p><p>Su código de ' + motivo + ' es:</p>' +
    '<p style="font-size:30px;letter-spacing:8px;font-weight:bold;color:#3a475c">' + codigo + '</p>' +
    '<p>Válido por ' + (vigencia || '72 horas') + '. Escríbalo en la aplicación desde <b>su propio teléfono</b>; la cuenta quedará ligada a ese dispositivo.</p>' +
    (url ? '<p><a href="' + url + '" style="background:#3a475c;color:#fff;padding:10px 16px;border-radius:8px;text-decoration:none">Abrir CTPL Asistencia</a></p>' : ''));
}

/* ---------- Marcas ---------- */
function filaMarca_(r) {
  return { id: r[0], fh: r[1], fecha: r[2], hora: r[3], ced: limpiarCedula_(r[4]), nombre: r[5], tipo: r[6], pantalla: r[7],
    lat: r[8], lng: r[9], pre: r[10], metodo: r[11], estado: r[12], obs: r[13], reg: r[14] };
}
/** Lee marcas entre dos fechas (AAAA-MM-DD). maxFilas limita la lectura a las filas más recientes. */
function leerMarcas_(desde, hasta, maxFilas) {
  const sh = sh_(H_MARCAS), n = sh.getLastRow();
  if (n < 2) return [];
  const k = Math.min(n - 1, maxFilas || 100000), ini = n - k + 1;
  // Lee solo la columna de fecha y luego únicamente los tramos de filas que corresponden al rango
  const fechas = sh.getRange(ini, 3, k, 1).getDisplayValues();
  const tramos = []; let t = null;
  for (let i = 0; i < k; i++) {
    const f = fechas[i][0], ok = f >= desde && f <= hasta;
    if (ok && !t) t = [i, i]; else if (ok) t[1] = i; else if (t) { tramos.push(t); t = null; }
  }
  if (t) tramos.push(t);
  let out = [];
  tramos.forEach(x => {
    out = out.concat(sh.getRange(ini + x[0], 1, x[1] - x[0] + 1, C_MARCAS.length).getDisplayValues()
      .filter(r => r[2] >= desde && r[2] <= hasta).map(filaMarca_));
  });
  return out;
}
function marcasDia_(ced, fecha) {
  const cache = CacheService.getScriptCache(), k = 'md_' + ced + '_' + fecha, c = cache.get(k);
  if (c) return JSON.parse(c);
  const l = leerMarcas_(fecha, fecha, 3000).filter(m => m.ced === ced && m.estado !== 'ANULADA')
    .map(m => ({ tipo: m.tipo, hora: m.hora })).sort((a, b) => a.hora < b.hora ? -1 : 1);
  cache.put(k, JSON.stringify(l), 21600);
  return l;
}
function limpiarCacheMarcas_(ced, fecha) { CacheService.getScriptCache().remove('md_' + ced + '_' + fecha); }
function distancia_(a, b, c, d) {
  const R = 6371000, r = x => x * Math.PI / 180, dl = r(c - a), dn = r(d - b);
  const h = Math.sin(dl / 2) ** 2 + Math.cos(r(a)) * Math.cos(r(c)) * Math.sin(dn / 2) ** 2;
  return 2 * R * Math.asin(Math.sqrt(h));
}

/* ---------- Pantallas (kioscos) ---------- */
function pantallas_() {
  const sh = sh_(H_KIOSCOS), n = sh.getLastRow();
  return { sh: sh, v: n < 2 ? [] : sh.getRange(2, 1, n - 1, C_KIOSCOS.length).getDisplayValues() };
}
function tokenQR_(kid, slot) {
  const s = PropertiesService.getScriptProperties().getProperty('ASIS_SECRET');
  return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(
    Utilities.newBlob('Q|' + kid + '|' + slot).getBytes(), Utilities.base64Decode(s))).slice(0, 16);
}

/* ===================== API PÚBLICA: APLICACIÓN DEL FUNCIONARIO ===================== */

function asisInfo() {
  const c = cfg_();
  return { ahora: Date.now(), institucion: c.institucion, geo: !!c.geo, slot: c.slot };
}

function solicitarAcceso(cedula, correo) {
  const ced = limpiarCedula_(cedula);
  correo = correoValido_(correo);
  const p = personal_(ced);
  if (!p) throw new Error('Esa cédula no está registrada en el sistema de carnets. Primero complete su registro de carnet.');
  return bloquear_(() => {
    const c = buscarCuenta_(ced);
    if (correoEnUso_(c.todas, correo, ced)) throw new Error('Ese correo ya está ligado a otra persona.');
    if (c.r) {
      const e = c.r[CU.EST];
      if (e === 'ACTIVA') throw new Error('Usted ya tiene una cuenta activa. Use "Iniciar sesión".');
      if (e === 'BLOQUEADA') throw new Error('Su cuenta está bloqueada. Consulte con la administración.');
      if (e === 'APROBADA') throw new Error('Su correo ya fue aprobado. Use "Activar cuenta" con el código que recibió.');
      c.r[CU.COR] = correo; c.r[CU.EST] = 'PENDIENTE'; c.r[CU.FSOL] = Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm');
      guardarCuenta_(c);
    } else {
      const r = new Array(C_CUENTAS.length).fill('');
      r[CU.CED] = ced; r[CU.NOM] = p.nombre; r[CU.COR] = correo; r[CU.EST] = 'PENDIENTE';
      r[CU.FSOL] = Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm');
      c.sh.appendRow(r);
    }
    audit_(ced, 'SOLICITUD', correo);
    return { ok: true, nombre: p.nombre };
  });
}

function activarCuenta(cedula, codigo, clave, dev, info) {
  const ced = limpiarCedula_(cedula);
  validarClave_(clave);
  return bloquear_(() => {
    const c = buscarCuenta_(ced);
    if (!c.r || c.r[CU.EST] !== 'APROBADA') throw new Error('No hay una activación pendiente para esa cédula.');
    if (String(codigo).trim() !== c.r[CU.COD]) throw new Error('Código incorrecto.');
    if (Date.now() > Number(c.r[CU.CVEN])) throw new Error('El código venció. Solicite uno nuevo a la administración.');
    const llave = nuevaLlave_(), sal = Utilities.getUuid();
    c.r[CU.SAL] = sal; c.r[CU.HASH] = hashClave_(clave, sal);
    c.r[CU.DID] = dev; c.r[CU.DKEY] = llave; c.r[CU.DINF] = String(info || '').slice(0, 120);
    c.r[CU.DFEC] = Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm');
    c.r[CU.EST] = 'ACTIVA'; c.r[CU.COD] = ''; c.r[CU.CVEN] = ''; c.r[CU.PID] = ''; c.r[CU.PINF] = ''; c.r[CU.PEST] = '';
    guardarCuenta_(c);
    audit_(ced, 'ACTIVACION', c.r[CU.DINF]);
    return { llave: llave, ced: ced, nombre: c.r[CU.NOM] };
  });
}

function iniciarSesion(cedula, clave, dev, info, ts, sig) {
  const ced = limpiarCedula_(cedula);
  return bloquear_(() => {
    const c = buscarCuenta_(ced);
    if (!c.r) throw new Error('No existe una cuenta con esa cédula. Use "Solicitar acceso".');
    const e = c.r[CU.EST];
    if (e === 'VERIFICANDO') throw new Error('VERIFICAR_CORREO');
    if (e === 'PENDIENTE') throw new Error('Su cuenta está creada y espera la aprobación de la administración.');
    if (e === 'APROBADA') throw new Error('Su correo fue aprobado. Use "Activar cuenta" con el código recibido.');
    if (e === 'RECHAZADA') throw new Error('Su solicitud fue rechazada. Consulte con la administración.');
    if (e === 'BLOQUEADA') throw new Error('Su cuenta está bloqueada. Consulte con la administración.');
    if (hashClave_(String(clave), c.r[CU.SAL]) !== c.r[CU.HASH]) { audit_(ced, 'CLAVE_ERRONEA', info); throw new Error('Contraseña incorrecta.'); }
    const ahora = Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm');

    // 1) Mismo dispositivo con su llave
    if (c.r[CU.DID] === dev && c.r[CU.DKEY] && sig && Math.abs(Date.now() - Number(ts)) < 300000 &&
        firma_(c.r[CU.DKEY], [ced, dev, ts, 'login'].join('|')) === sig) {
      c.r[CU.ULT] = ahora; guardarCuenta_(c);
      return { ok: true, ced: ced, nombre: c.r[CU.NOM] };
    }
    const vincular = (motivo) => {
      const llave = nuevaLlave_();
      c.r[CU.DID] = dev; c.r[CU.DKEY] = llave; c.r[CU.DINF] = String(info || '').slice(0, 120); c.r[CU.DFEC] = ahora;
      c.r[CU.PID] = ''; c.r[CU.PINF] = ''; c.r[CU.PEST] = ''; c.r[CU.ULT] = ahora;
      guardarCuenta_(c); audit_(ced, 'VINCULACION', motivo + ' · ' + c.r[CU.DINF]);
      return { llave: llave, ced: ced, nombre: c.r[CU.NOM] };
    };
    // 2) Cambio de dispositivo aprobado por el administrador
    if (c.r[CU.PEST] === 'APROBADO') {
      if (c.r[CU.PID] === dev) return vincular('cambio aprobado');
      throw new Error('Ya hay otro dispositivo aprobado esperando vincularse. Consulte con la administración.');
    }
    // 3) Dispositivo liberado por el administrador o ventana de vinculación abierta
    if (!c.r[CU.DID]) return vincular('dispositivo liberado');
    if (Date.now() < Number(cfg_().ventanaHasta || 0)) return vincular('ventana de vinculación');
    // 4) Dispositivo nuevo: si está dentro del límite, puede revincularlo él mismo con código al correo
    if (puedeAuto_(c.r, cfg_())) return { auto: true, correo: ocultarCorreo_(c.r[CU.COR]) };
    // 5) Fuera del límite: queda pendiente de aprobación del administrador
    c.r[CU.PID] = dev; c.r[CU.PINF] = String(info || '').slice(0, 120) + ' · ' + ahora; c.r[CU.PEST] = 'PENDIENTE';
    guardarCuenta_(c);
    audit_(ced, 'SOLICITUD_DISPOSITIVO', c.r[CU.PINF]);
    return { pendiente: true, limite: !!(+cfg_().autoDias) };
  });
}

/* ---------- REGISTRO UNIFICADO: carnet + cuenta + verificación de correo ---------- */

/** Un solo paso: guarda el carnet (datos y foto), crea la cuenta y envía un código al correo. */
function registroCompleto(d) {
  const ced = limpiarCedula_(d.cedula);
  const correo = correoValido_(d.correo);
  validarClave_(d.clave);
  if (!d.dev) throw new Error('No se pudo identificar este teléfono. Recargue la página.');
  const prev = buscarCuenta_(ced);
  if (prev.r && ['ACTIVA', 'BLOQUEADA', 'PENDIENTE'].indexOf(prev.r[CU.EST]) >= 0 && prev.r[CU.HASH])
    throw new Error('Ya existe una cuenta con esa cédula. Use "Ya tengo cuenta".');
  if (correoEnUso_(prev.todas, correo, ced)) throw new Error('Ese correo ya está ligado a otra persona.');
  const reg = registrar(d);   // valida y guarda el carnet
  return bloquear_(() => {
    const c = buscarCuenta_(ced);
    const r = c.r || new Array(C_CUENTAS.length).fill('');
    const sal = Utilities.getUuid();
    r[CU.CED] = ced; r[CU.NOM] = reg.nombre; r[CU.COR] = correo; r[CU.EST] = 'VERIFICANDO';
    r[CU.SAL] = sal; r[CU.HASH] = hashClave_(d.clave, sal);
    r[CU.PID] = d.dev; r[CU.PINF] = String(d.info || '').slice(0, 120); r[CU.PEST] = 'REGISTRO';
    r[CU.COD] = codigo6_(); r[CU.CVEN] = String(Date.now() + 30 * 60e3);
    r[CU.FSOL] = Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm');
    if (c.r) { c.r = r; guardarCuenta_(c); } else c.sh.appendRow(r);
    const env = correoCodigo_(correo, reg.nombre, r[CU.COD], 'verificación de correo', '30 minutos');
    audit_(ced, 'REGISTRO', correo);
    return { ced: ced, nombre: reg.nombre, correo: ocultarCorreo_(correo), enviado: env };
  });
}

/** Confirma el código del correo y liga el teléfono. La cuenta queda esperando aprobación del administrador. */
function verificarRegistro(cedula, codigo, dev, info) {
  const ced = limpiarCedula_(cedula);
  return bloquear_(() => {
    const c = buscarCuenta_(ced);
    if (!c.r || c.r[CU.EST] !== 'VERIFICANDO') throw new Error('No hay una verificación pendiente para esa cédula.');
    if (c.r[CU.PID] !== dev) throw new Error('Debe verificar desde el mismo teléfono donde creó la cuenta.');
    if (String(codigo).trim() !== c.r[CU.COD]) throw new Error('Código incorrecto.');
    if (Date.now() > Number(c.r[CU.CVEN])) throw new Error('El código venció. Toque "Enviar otro código".');
    const llave = nuevaLlave_();
    c.r[CU.DID] = dev; c.r[CU.DKEY] = llave; c.r[CU.DINF] = String(info || c.r[CU.PINF] || '').slice(0, 120);
    c.r[CU.DFEC] = Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm');
    c.r[CU.PID] = ''; c.r[CU.PINF] = ''; c.r[CU.PEST] = ''; c.r[CU.COD] = ''; c.r[CU.CVEN] = '';
    c.r[CU.EST] = 'PENDIENTE';
    guardarCuenta_(c);
    audit_(ced, 'CORREO_VERIFICADO', c.r[CU.DINF]);
    return { llave: llave, ced: ced, nombre: c.r[CU.NOM], estado: 'PENDIENTE' };
  });
}

/** Reenvía el código de verificación (puede usarse desde otro teléfono con la contraseña). */
function reenviarCodigoRegistro(cedula, clave, dev, info) {
  const ced = limpiarCedula_(cedula);
  return bloquear_(() => {
    const c = buscarCuenta_(ced);
    if (!c.r || c.r[CU.EST] !== 'VERIFICANDO') throw new Error('No hay una verificación pendiente para esa cédula.');
    if (hashClave_(String(clave), c.r[CU.SAL]) !== c.r[CU.HASH]) throw new Error('Contraseña incorrecta.');
    c.r[CU.COD] = codigo6_(); c.r[CU.CVEN] = String(Date.now() + 30 * 60e3);
    c.r[CU.PID] = dev; c.r[CU.PINF] = String(info || '').slice(0, 120);
    guardarCuenta_(c);
    correoCodigo_(c.r[CU.COR], c.r[CU.NOM], c.r[CU.COD], 'verificación de correo', '30 minutos');
    return { correo: ocultarCorreo_(c.r[CU.COR]) };
  });
}

/** Estado de la cuenta para la app (sirve mientras espera aprobación). */
function estadoCuenta(a) {
  const c = verificarDisp_(a, 'estado', true);
  const p = personal_(c.r[CU.CED]) || {};
  return { estado: c.r[CU.EST], nombre: c.r[CU.NOM], ced: c.r[CU.CED], puesto: p.puesto || '', genero: p.genero || 'M',
    foto: c.r[CU.EST] === 'PENDIENTE' ? fotoDataUrl_(p.fotoId) : '' };
}

/** Paso 1 del cambio de teléfono sin aprobación: envía un código al correo aprobado. */
function solicitarRevinculo(cedula, clave, dev, info) {
  const ced = limpiarCedula_(cedula);
  return bloquear_(() => {
    const c = buscarCuenta_(ced);
    if (!c.r || c.r[CU.EST] !== 'ACTIVA') throw new Error('No hay una cuenta activa con esa cédula.');
    if (hashClave_(String(clave), c.r[CU.SAL]) !== c.r[CU.HASH]) throw new Error('Contraseña incorrecta.');
    if (!puedeAuto_(c.r, cfg_())) throw new Error('Ya usó su cambio de teléfono de este periodo. Debe autorizarlo la administración.');
    c.r[CU.COD] = codigo6_(); c.r[CU.CVEN] = String(Date.now() + 15 * 60e3);
    c.r[CU.PID] = dev; c.r[CU.PINF] = String(info || '').slice(0, 120) + ' · ' + Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm');
    c.r[CU.PEST] = 'CODIGO';
    guardarCuenta_(c);
    enviarCorreo_(c.r[CU.COR], 'Código para cambiar de teléfono — CTPL Asistencia',
      '<p>Hola, ' + c.r[CU.NOM] + '.</p><p>Se solicitó ligar su cuenta a un teléfono nuevo (' + escapar_(c.r[CU.PINF]) + '). Su código es:</p>' +
      '<p style="font-size:30px;letter-spacing:8px;font-weight:bold;color:#3a475c">' + c.r[CU.COD] + '</p>' +
      '<p>Válido por 15 minutos. Al confirmarlo, el teléfono anterior dejará de funcionar.</p>' +
      '<p style="color:#a03b3b">Si usted no lo solicitó, no comparta este código y avise a la administración.</p>');
    audit_(ced, 'SOLICITA_REVINCULO', c.r[CU.PINF]);
    return { correo: ocultarCorreo_(c.r[CU.COR]) };
  });
}

/** Paso 2: confirma el código y liga el teléfono nuevo. El anterior queda revocado y se avisa a la administración. */
function confirmarRevinculo(cedula, codigo, dev, info) {
  const ced = limpiarCedula_(cedula), cfg = cfg_();
  const res = bloquear_(() => {
    const c = buscarCuenta_(ced);
    if (!c.r || c.r[CU.EST] !== 'ACTIVA' || c.r[CU.PEST] !== 'CODIGO' || c.r[CU.PID] !== dev) throw new Error('No hay un cambio de teléfono pendiente en este dispositivo.');
    if (String(codigo).trim() !== c.r[CU.COD] || Date.now() > Number(c.r[CU.CVEN])) throw new Error('Código incorrecto o vencido.');
    if (!puedeAuto_(c.r, cfg)) throw new Error('Ya usó su cambio de teléfono de este periodo.');
    const anterior = c.r[CU.DINF] || '—', ahora = Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm'), llave = nuevaLlave_();
    c.r[CU.DID] = dev; c.r[CU.DKEY] = llave; c.r[CU.DINF] = String(info || '').slice(0, 120); c.r[CU.DFEC] = ahora;
    c.r[CU.PID] = ''; c.r[CU.PINF] = ''; c.r[CU.PEST] = ''; c.r[CU.COD] = ''; c.r[CU.CVEN] = '';
    c.r[CU.AUTO] = String(Date.now()); c.r[CU.ULT] = ahora;
    guardarCuenta_(c);
    audit_(ced, 'AUTOVINCULO', anterior + ' → ' + c.r[CU.DINF]);
    return { llave: llave, ced: ced, nombre: c.r[CU.NOM], anterior: anterior, nuevo: c.r[CU.DINF], ahora: ahora };
  });
  if (cfg.correos) enviarCorreo_(cfg.correos, 'Cambio de teléfono: ' + res.nombre,
    '<p><b>' + res.nombre + '</b> (' + ced + ') ligó su cuenta a un teléfono nuevo usando el código enviado a su correo.</p>' +
    '<p>Anterior: ' + escapar_(res.anterior) + '<br>Nuevo: ' + escapar_(res.nuevo) + '<br>Fecha: ' + res.ahora + '</p>' +
    '<p style="font-size:12px;color:#5a6577">Puede revisarlo en Administración › Personal. Las marcas de hoy quedarán señaladas como "dispositivo nuevo".</p>');
  return { llave: res.llave, ced: res.ced, nombre: res.nombre };
}
function escapar_(t) { return String(t || '').replace(/[&<>]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[c])); }

function perfil(a, fotoConocida) {
  const c = verificarDisp_(a, 'perfil');
  const ced = c.r[CU.CED], p = personal_(ced) || {}, cfg = cfg_(), f = hoy_();
  const cache = CacheService.getScriptCache();
  if (!cache.get('ult_' + ced)) { c.r[CU.ULT] = Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm'); guardarCuenta_(c); cache.put('ult_' + ced, '1', 600); }
  const sem = calcularSemana_(lunesDe_(f), ced)[0] || null;
  const h = horarioDia_(c.r, f, cfg);
  let hor = {}; try { hor = JSON.parse(c.r[CU.HOR] || '{}'); } catch (e) {}
  return {
    ced: ced, nombre: c.r[CU.NOM], correo: c.r[CU.COR], puesto: p.puesto || '', genero: p.genero || 'M',
    fotoId: p.fotoId || '', foto: p.fotoId && p.fotoId === fotoConocida ? '' : fotoDataUrl_(p.fotoId), dispositivo: c.r[CU.DINF], vinculado: c.r[CU.DFEC],
    hoy: { fecha: f, marcas: marcasDia_(ced, f), horario: h ? h.txt : 'No labora hoy' },
    horario: [1, 2, 3, 4, 5].map(d => { let t; try { t = tramos_(hor[d], cfg); } catch (e) { t = tramos_('', cfg); } return t.length ? t.map(x => x.ent + '-' + x.sal).join(' y ') : 'No labora'; }),
    semana: sem, geo: !!cfg.geo, ahora: Date.now()
  };
}

function marcar(a, qr, geo) {
  const c = verificarDisp_(a, 'marcar|' + qr);
  const cfg = cfg_();
  const m = String(qr || '').match(/^CTPL\|([A-Z0-9]+)\|(\d+)\|([A-Za-z0-9_-]{16})$/);
  if (!m) throw new Error('Ese código no es un QR de marcas del CTPL.');
  const slot = Math.floor(Date.now() / 1000 / cfg.slot), s = Number(m[2]);
  if (s > slot + 1 || s < slot - 2) throw new Error('El código QR ya venció. Escanee el código que está en pantalla ahora.');
  const k = pantallas_().v.find(r => r[0] === m[1]);
  if (!k || k[3] !== 'ACTIVA') throw new Error('La pantalla de marcas no está autorizada.');
  if (tokenQR_(m[1], s) !== m[3]) throw new Error('Código QR no válido.');
  if (cfg.geo) {
    if (!geo || geo.lat == null) throw new Error('Active la ubicación del teléfono para marcar.');
    const d = distancia_(cfg.lat, cfg.lng, +geo.lat, +geo.lng) - Math.min(+geo.acc || 0, 150);
    if (d > cfg.radio) throw new Error('Debe estar dentro del colegio para marcar (a ' + Math.round(d) + ' m del punto autorizado).');
  }
  const ced = c.r[CU.CED];
  return bloquear_(() => {
    const ahora = new Date(), f = hoy_(ahora), h = hora_(ahora);
    const prev = marcasDia_(ced, f);
    const ult = prev[prev.length - 1];
    if (ult && (min_(h) - min_(ult.hora)) < cfg.separacion)
      throw new Error('Ya registró ' + ult.tipo + ' a las ' + ult.hora.slice(0, 5) + '. Espere ' + cfg.separacion + ' minutos para volver a marcar.');
    const tipo = ult && ult.tipo === 'ENTRADA' ? 'SALIDA' : 'ENTRADA';
    const id = 'M' + ahora.getTime().toString(36).toUpperCase() + Math.floor(Math.random() * 1296).toString(36).toUpperCase();
    sh_(H_MARCAS).appendRow([id, Utilities.formatDate(ahora, TZ, 'dd/MM/yyyy HH:mm:ss'), f, h, ced, c.r[CU.NOM], tipo, m[1],
      geo && geo.lat != null ? +(+geo.lat).toFixed(6) : '', geo && geo.lng != null ? +(+geo.lng).toFixed(6) : '',
      geo && geo.acc != null ? Math.round(geo.acc) : '', 'QR', 'VÁLIDA',
      c.r[CU.AUTO] && hoy_(new Date(Number(c.r[CU.AUTO]))) === f ? 'Dispositivo nuevo' : '', 'APP']);
    prev.push({ tipo: tipo, hora: h });
    CacheService.getScriptCache().put('md_' + ced + '_' + f, JSON.stringify(prev), 21600);
    let aviso = '';
    const hd = horarioDia_(c.r, f, cfg);
    if (!hd) aviso = 'Hoy no tiene jornada asignada; la marca quedó registrada.';
    else {
      const i = tramoDe_(hd.tramos, h), t = hd.tramos[i];
      const delTramo = prev.filter(m => tramoDe_(hd.tramos, m.hora) === i);
      if (tipo === 'ENTRADA' && delTramo.length === 1 && min_(h) > min_(t.ent) + cfg.tolerancia) aviso = 'Entrada registrada con tardía (horario ' + t.ent + ').';
      if (tipo === 'SALIDA' && min_(h) < min_(t.sal)) aviso = 'Salida antes del horario (' + t.sal + ').';
    }
    return { tipo: tipo, hora: h, fecha: f, nombre: c.r[CU.NOM], aviso: aviso, marcas: prev };
  });
}

function misMarcas(a, desde, hasta) {
  const c = verificarDisp_(a, 'marcas');
  const ced = c.r[CU.CED];
  return leerMarcas_(desde, hasta).filter(m => m.ced === ced)
    .map(m => ({ fecha: m.fecha, hora: m.hora, tipo: m.tipo, metodo: m.metodo, estado: m.estado, obs: m.obs }));
}

function misInformes(a) {
  const c = verificarDisp_(a, 'informes');
  const ced = c.r[CU.CED], sh = sh_(H_INFORMES), n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, C_INFORMES.length).getDisplayValues().filter(r => limpiarCedula_(r[1]) === ced)
    .map(informeFila_).sort((x, y) => x.semana < y.semana ? 1 : -1).slice(0, 26);
}

function cambiarClave(a, actual, nueva) {
  const c = verificarDisp_(a, 'clave');
  if (hashClave_(String(actual), c.r[CU.SAL]) !== c.r[CU.HASH]) throw new Error('La contraseña actual no es correcta.');
  validarClave_(nueva);
  c.r[CU.SAL] = Utilities.getUuid(); c.r[CU.HASH] = hashClave_(nueva, c.r[CU.SAL]);
  guardarCuenta_(c); audit_(c.r[CU.CED], 'CAMBIO_CLAVE', '');
  return true;
}

function recuperarClave(cedula) {
  const ced = limpiarCedula_(cedula);
  return bloquear_(() => {
    const c = buscarCuenta_(ced);
    if (!c.r || c.r[CU.EST] !== 'ACTIVA') throw new Error('No hay una cuenta activa con esa cédula.');
    c.r[CU.COD] = codigo6_(); c.r[CU.CVEN] = String(Date.now() + 72 * 3600e3);
    guardarCuenta_(c);
    correoCodigo_(c.r[CU.COR], c.r[CU.NOM], c.r[CU.COD], 'recuperación de contraseña');
    return { correo: ocultarCorreo_(c.r[CU.COR]) };
  });
}

function restablecerClave(cedula, codigo, nueva) {
  const ced = limpiarCedula_(cedula);
  validarClave_(nueva);
  return bloquear_(() => {
    const c = buscarCuenta_(ced);
    if (!c.r || c.r[CU.EST] !== 'ACTIVA' || !c.r[CU.COD]) throw new Error('No hay una recuperación pendiente.');
    if (String(codigo).trim() !== c.r[CU.COD] || Date.now() > Number(c.r[CU.CVEN])) throw new Error('Código incorrecto o vencido.');
    c.r[CU.SAL] = Utilities.getUuid(); c.r[CU.HASH] = hashClave_(nueva, c.r[CU.SAL]); c.r[CU.COD] = ''; c.r[CU.CVEN] = '';
    guardarCuenta_(c); audit_(ced, 'RECUPERACION_CLAVE', '');
    return true;
  });
}

/* ===================== API: PANTALLA DE MARCAS ===================== */

function kioscoRegistrar(clave, nombre) {
  const real = PropertiesService.getScriptProperties().getProperty('ADMIN_PASS') || CLAVE_INICIAL;
  if (String(clave) !== real) throw new Error('Clave de administrador incorrecta.');
  const kid = 'P' + Utilities.getUuid().replace(/-/g, '').slice(0, 6).toUpperCase();
  const llave = nuevaLlave_();
  sh_(H_KIOSCOS).appendRow([kid, String(nombre || 'Pantalla').toUpperCase().slice(0, 40), llave, 'ACTIVA',
    Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm'), '']);
  audit_('ADMIN', 'PANTALLA_NUEVA', kid + ' ' + nombre);
  return { kid: kid, llave: llave, nombre: String(nombre || 'Pantalla').toUpperCase() };
}

function kioscoPulso(kid, ts, sig, ultimoId) {
  const t = pantallas_(), i = t.v.findIndex(r => r[0] === kid);
  if (i < 0 || t.v[i][3] !== 'ACTIVA') throw new Error('PANTALLA_NO_AUTORIZADA');
  if (Math.abs(Date.now() - Number(ts)) > 300000 || firma_(t.v[i][2], kid + '|' + ts + '|pulso') !== sig) throw new Error('PANTALLA_NO_AUTORIZADA');
  const cache = CacheService.getScriptCache();
  if (!cache.get('pulso_' + kid)) { t.sh.getRange(i + 2, 6).setValue(Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm')); cache.put('pulso_' + kid, '1', 120); }
  const cfg = cfg_(), slot = Math.floor(Date.now() / 1000 / cfg.slot);
  const f = hoy_();
  const ult = leerMarcas_(f, f, 40).filter(m => m.estado !== 'ANULADA').slice(-8).reverse()
    .map(m => ({ id: m.id, nombre: corto_(m.nombre), tipo: m.tipo, hora: m.hora.slice(0, 5), ced: m.ced }));
  let foto = null;
  if (ult.length && ult[0].id !== ultimoId) {
    const p = personal_(ult[0].ced);
    foto = { id: ult[0].id, src: p ? fotoDataUrl_(p.fotoId) : '' };
  }
  ult.forEach(u => delete u.ced);
  return { qr: 'CTPL|' + kid + '|' + slot + '|' + tokenQR_(kid, slot), slot: slot, seg: cfg.slot, ahora: Date.now(),
    marcas: ult, foto: foto, institucion: cfg.institucion, nombre: t.v[i][1] };
}

/* ===================== INFORMES ===================== */

function calcularSemana_(lunes, soloCed) {
  const cfg = cfg_(), viernes = sumarDias_(lunes, 4), hoyS = hoy_(), ahoraMin = min_(hora_());
  const marcas = leerMarcas_(lunes, viernes).filter(m => m.estado !== 'ANULADA');
  const noLect = {}; String(cfg.noLectivos || '').split(/[,\s]+/).filter(Boolean).forEach(f => noLect[f] = 1);
  const per = mapaPersonal_();
  return cuentas_().v.filter(r => (r[CU.EST] === 'ACTIVA' || r[CU.EST] === 'BLOQUEADA') && (!soloCed || r[CU.CED] === soloCed)).map(r => {
    const ced = limpiarCedula_(r[CU.CED]), mis = marcas.filter(m => m.ced === ced);
    const t = { prog: 0, asis: 0, horas: 0, tard: 0, salAnt: 0, aus: 0, inc: 0 }, dias = [];
    const fs = String(r[CU.FSOL] || '').match(/(\d{2})\/(\d{2})\/(\d{4})/), alta = fs ? fs[3] + '-' + fs[2] + '-' + fs[1] : '';
    for (let i = 0; i < 5; i++) {
      const f = sumarDias_(lunes, i), h = horarioDia_(r, f, cfg);
      const md = mis.filter(m => m.fecha === f).sort((a, b) => a.hora < b.hora ? -1 : 1);
      const d = { fecha: f, horario: h ? h.txt : 'No labora', marcas: md.map(m => m.tipo.charAt(0) + ' ' + m.hora.slice(0, 5)), horas: 0, obs: [] };
      if (noLect[f]) { d.obs.push('No lectivo'); dias.push(d); continue; }
      if (f > hoyS) { d.obs.push('Pendiente'); dias.push(d); continue; }
      if (!h && !md.length) { d.obs.push('No labora'); dias.push(d); continue; }
      if (alta && f < alta && !md.length) { d.obs.push('Antes del registro'); dias.push(d); continue; }
      if (h) t.prog++;
      if (md.length) {
        t.asis++;
        let ent = null;
        md.forEach(m => { if (m.tipo === 'ENTRADA') ent = m; else if (ent) { d.horas += (min_(m.hora) - min_(ent.hora)) / 60; ent = null; } });
        if (ent && f < hoyS) { t.inc++; d.obs.push('Sin salida'); }
        if (md.some(m => m.obs.indexOf('Dispositivo nuevo') >= 0)) d.obs.push('Dispositivo nuevo');
        if (h) { const e = evaluarDia_(h, md, f < hoyS, ahoraMin, cfg.tolerancia); t.tard += e.tard; t.salAnt += e.salAnt; t.aus += e.aus; d.obs = d.obs.concat(e.obs); }
        else d.obs.push('Día sin jornada');
      } else if (h) { const e = evaluarDia_(h, md, f < hoyS, ahoraMin, cfg.tolerancia); t.aus += e.aus; d.obs = d.obs.concat(e.obs); }
      else d.obs.push('No labora');
      t.horas += d.horas; d.horas = Math.round(d.horas * 100) / 100;
      dias.push(d);
    }
    const p = per[ced] || {};
    return { semana: lunes, ced: ced, nombre: r[CU.NOM], puesto: p.puesto || '', prog: t.prog, asis: t.asis,
      horas: hhmm_(t.horas), tard: t.tard, salAnt: t.salAnt, aus: t.aus, inc: t.inc, dias: dias };
  }).sort((a, b) => a.nombre.localeCompare(b.nombre));
}
function informeFila_(r) {
  let d = []; try { d = JSON.parse(r[11] || '[]'); } catch (e) {}
  return { semana: r[0], ced: limpiarCedula_(r[1]), nombre: r[2], puesto: r[3], prog: +r[4], asis: +r[5], horas: r[6],
    tard: +r[7], salAnt: +r[8], aus: +r[9], inc: +r[10], dias: d, generado: r[12] };
}

function generarInforme_(lunes, enviar) {
  const res = calcularSemana_(lunes);
  bloquear_(() => {
    const sh = sh_(H_INFORMES), n = sh.getLastRow();
    const otras = n < 2 ? [] : sh.getRange(2, 1, n - 1, C_INFORMES.length).getDisplayValues().filter(r => r[0] !== lunes);
    const gen = Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm');
    const nuevas = res.map(x => [x.semana, x.ced, x.nombre, x.puesto, x.prog, x.asis, x.horas, x.tard, x.salAnt, x.aus, x.inc, JSON.stringify(x.dias), gen]);
    const todas = otras.concat(nuevas);
    if (n > 1) sh.getRange(2, 1, n - 1, C_INFORMES.length).clearContent();
    if (todas.length) sh.getRange(2, 1, todas.length, C_INFORMES.length).setValues(todas);
  });
  let enviado = false;
  const cfg = cfg_();
  if (enviar && cfg.correos) {
    const tot = k => res.reduce((s, x) => s + x[k], 0);
    const filas = res.map(x => '<tr><td style="padding:4px 8px;border-bottom:1px solid #e3e7ee">' + x.nombre + '</td>' +
      [x.asis + '/' + x.prog, x.horas, x.tard, x.salAnt, x.aus, x.inc].map(v => '<td style="padding:4px 8px;text-align:center;border-bottom:1px solid #e3e7ee">' + v + '</td>').join('') + '</tr>').join('');
    const nom = 'Asistencia_CTPL_semana_' + lunes;
    enviado = enviarCorreo_(cfg.correos, 'Informe semanal de asistencia — semana del ' + fechaLarga_(lunes),
      '<p>Informe consolidado del personal, semana del <b>' + fechaLarga_(lunes) + '</b> al <b>' + fechaLarga_(sumarDias_(lunes, 4)) + '</b>.</p>' +
      '<p><b>' + res.length + '</b> funcionarios · <b>' + tot('tard') + '</b> tardías · <b>' + tot('aus') + '</b> ausencias · <b>' + tot('inc') + '</b> marcas incompletas.</p>' +
      '<table style="border-collapse:collapse;font-size:12px;width:100%"><tr style="background:#3a475c;color:#fff">' +
      ['Funcionario', 'Días', 'Horas', 'Tard.', 'Sal. ant.', 'Aus.', 'Incomp.'].map(t => '<th style="padding:6px 8px">' + t + '</th>').join('') + '</tr>' + filas + '</table>' +
      '<p style="font-size:12px;color:#5a6577">Se adjunta el informe en PDF, y el resumen y el detalle diario en CSV (separador punto y coma, abre en Excel).</p>',
      [pdfInforme_(res, lunes, false),
       Utilities.newBlob('﻿' + csvResumen_(res), 'text/csv', nom + '_resumen.csv'),
       Utilities.newBlob('﻿' + csvDetalle_(res), 'text/csv', nom + '_detalle.csv')]);
  }
  audit_('SISTEMA', 'INFORME', lunes + (enviado ? ' enviado' : ''));
  return { cantidad: res.length, enviado: enviado };
}
function fechaLarga_(f) { return Utilities.formatDate(fechaD_(f), TZ, 'dd/MM/yyyy'); }

/** Disparador cada 5 min: mantiene el servidor "despierto" de 5:00 a. m. a 10:00 p. m. (incluye jornada nocturna) */
function mantenerActivo() {
  const h = +Utilities.formatDate(new Date(), TZ, 'H'), d = +Utilities.formatDate(new Date(), TZ, 'u');
  if (d > 5 || h < 5 || h >= 22) return;
  mapaPersonal_(); cfg_();
}

/** Disparador automático: viernes 10:00 p. m. (después de la jornada nocturna) */
function informeSemanalAuto() { generarInforme_(lunesDe_(hoy_()), true); }

/* ---------- CSV ---------- */
function csv_(filas) {
  return filas.map(r => r.map(v => { v = String(v == null ? '' : v); return /[;"\n]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v; }).join(';')).join('\r\n');
}
function csvResumen_(res) {
  return csv_([['Semana', 'Cédula', 'Nombre', 'Puesto', 'Días programados', 'Días asistidos', 'Horas (hh:mm)', 'Tardías', 'Salidas anticipadas', 'Ausencias', 'Marcas incompletas']]
    .concat(res.map(x => [x.semana, x.ced, x.nombre, x.puesto, x.prog, x.asis, x.horas, x.tard, x.salAnt, x.aus, x.inc])));
}
function csvDetalle_(res) {
  const f = [['Fecha', 'Cédula', 'Nombre', 'Horario', 'Marcas', 'Horas (hh:mm)', 'Observaciones']];
  res.forEach(x => x.dias.forEach(d => f.push([d.fecha, x.ced, x.nombre, d.horario, d.marcas.join(' | '), hhmm_(d.horas), d.obs.join(', ')])));
  return csv_(f);
}
function csvMarcas_(l) {
  return csv_([['Fecha', 'Hora', 'Cédula', 'Nombre', 'Tipo', 'Pantalla', 'Método', 'Estado', 'Observación', 'Latitud', 'Longitud', 'Registrado por']]
    .concat(l.map(m => [m.fecha, m.hora, m.ced, m.nombre, m.tipo, m.pantalla, m.metodo, m.estado, m.obs, m.lat, m.lng, m.reg])));
}
function empaquetar_(nombre, texto, formato) {
  const b = Utilities.newBlob('﻿' + texto, 'text/csv', nombre + '.csv');
  if (formato === 'zip') {
    const z = Utilities.zip([b], nombre + '.zip');
    return { nombre: nombre + '.zip', mime: 'application/zip', b64: Utilities.base64Encode(z.getBytes()) };
  }
  return { nombre: nombre + '.csv', mime: 'text/csv', b64: Utilities.base64Encode(b.getBytes()) };
}

/** Descarga del propio funcionario */
function miExportacion(a, desde, hasta) {
  const c = verificarDisp_(a, 'exportar');
  const l = leerMarcas_(desde, hasta).filter(m => m.ced === c.r[CU.CED]);
  return empaquetar_('Mis_marcas_' + desde + '_' + hasta, csvMarcas_(l), 'csv');
}

/** Informe semanal del propio funcionario en PDF */
function miInformePdf(a, lunes) {
  const c = verificarDisp_(a, 'informepdf');
  const ced = c.r[CU.CED]; lunes = lunesDe_(lunes);
  const sh = sh_(H_INFORMES), n = sh.getLastRow();
  let filas = n < 2 ? [] : sh.getRange(2, 1, n - 1, C_INFORMES.length).getDisplayValues()
    .filter(r => r[0] === lunes && limpiarCedula_(r[1]) === ced).map(informeFila_);
  if (!filas.length) filas = calcularSemana_(lunes, ced);
  if (!filas.length) throw new Error('No hay información para esa semana.');
  const b = pdfInforme_(filas, lunes, true);
  return { nombre: b.getName(), mime: 'application/pdf', b64: Utilities.base64Encode(b.getBytes()) };
}

/* ---------- Informe en PDF (formato institucional) ---------- */
function pdfInforme_(filas, lunes, individual) {
  const html = htmlInforme_(filas, lunes, individual);
  const nombre = individual ? 'Informe_asistencia_' + filas[0].ced + '_' + lunes + '.pdf' : 'Informe_asistencia_CTPL_' + lunes + '.pdf';
  return Utilities.newBlob(html, 'text/html', 'informe.html').getAs('application/pdf').setName(nombre);
}
function htmlInforme_(filas, lunes, individual) {
  const cfg = cfg_(), e = escapar_, dias = ['Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes'];
  const PU = { DOCENTE_TECNICO: 'Docente técnico', DOCENTE_ACADEMICO: 'Docente académico', DOCENTE_IDIOMAS: 'Docente de Idiomas', ADMINISTRATIVO: 'Personal administrativo',
    AUXILIAR_ADMIN: 'Auxiliar administrativo', COORD_TECNICO: 'Coordinación técnica', COORD_EMPRESA: 'Coordinación con la empresa', COORD_ACADEMICO: 'Coordinación académica',
    SUBDIRECTOR: 'Subdirección', DIRECTOR: 'Dirección' };
  const tot = k => filas.reduce((s, x) => s + (+x[k] || 0), 0);
  const num = v => (+v ? '<span style="background:#f8eedb;color:#7d5517;padding:1px 6px;border-radius:8px;font-weight:bold">' + v + '</span>' : '<span style="color:#9aa3b5">0</span>');
  const gen = Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm');
  const th = 'background:#3a475c;color:#ffffff;font-size:8pt;text-transform:uppercase;letter-spacing:.5px;padding:6px 7px;text-align:left';
  const td = 'padding:5px 7px;border-bottom:1px solid #e2e6ec;font-size:8.6pt;vertical-align:top';
  let h = '<html><head><meta charset="utf-8"><style>@page{size:A4;margin:14mm 13mm}body{font-family:Arial,Helvetica,sans-serif;color:#1f2836;margin:0}' +
    'table{border-collapse:collapse;width:100%}.blq{page-break-inside:avoid}</style></head><body>';
  h += '<table><tr><td style="background:#2f3b4f;color:#fff;padding:14px 18px">' +
    '<div style="font-size:7.5pt;letter-spacing:2px;color:#c9d3e1;font-weight:bold">' + e(cfg.institucion).toUpperCase() + ' · MINISTERIO DE EDUCACIÓN PÚBLICA</div>' +
    '<div style="font-size:17pt;font-weight:bold;margin-top:4px">' + (individual ? 'Informe semanal de asistencia' : 'Informe semanal de asistencia del personal') + '</div>' +
    '<div style="font-size:9.5pt;color:#dfe4eb;margin-top:3px">Semana del ' + fechaLarga_(lunes) + ' al ' + fechaLarga_(sumarDias_(lunes, 4)) + '</div></td>' +
    '<td style="background:#46566f;color:#fff;padding:14px 16px;width:150px;text-align:right;font-size:8pt">Generado<br><b style="font-size:10pt">' + gen + '</b><br>CTPL Asistencia</td></tr></table>' +
    '<div style="height:4px;background:#c9d3e1"></div>';
  if (individual) {
    const x = filas[0];
    h += '<table style="margin-top:12px"><tr><td style="padding:10px 12px;background:#eef1f5;border:1px solid #cfd5de">' +
      '<div style="font-size:7.5pt;color:#55606f;text-transform:uppercase;letter-spacing:1px">Funcionario</div><div style="font-size:13pt;font-weight:bold">' + e(x.nombre) + '</div>' +
      '<div style="font-size:9pt;color:#55606f">Cédula ' + x.ced + ' · ' + e(PU[x.puesto] || x.puesto) + '</div></td></tr></table>';
  }
  const k = individual ? [[filas[0].asis + '/' + filas[0].prog, 'Días asistidos'], [filas[0].horas, 'Horas'], [tot('tard'), 'Tardías'], [tot('salAnt'), 'Salidas anticipadas'], [tot('aus'), 'Ausencias'], [tot('inc'), 'Sin marca de salida']]
    : [[filas.length, 'Funcionarios'], [tot('tard'), 'Tardías'], [tot('salAnt'), 'Salidas anticipadas'], [tot('aus'), 'Ausencias'], [tot('inc'), 'Sin marca de salida']];
  h += '<table style="margin-top:12px"><tr>' + k.map(v => '<td style="padding:4px"><div style="background:#f1f3f6;border:1px solid #cfd5de;padding:9px 6px;text-align:center">' +
    '<div style="font-size:16pt;font-weight:bold;color:#2f3b4f">' + v[0] + '</div><div style="font-size:7pt;color:#55606f;text-transform:uppercase;letter-spacing:.5px">' + v[1] + '</div></div></td>').join('') + '</tr></table>';
  if (!individual) {
    h += '<div style="font-size:10pt;font-weight:bold;color:#3a475c;margin:14px 0 6px;text-transform:uppercase;letter-spacing:1px">Resumen</div><table>' +
      '<tr><th style="' + th + ';width:22px">N.º</th><th style="' + th + '">Funcionario</th><th style="' + th + '">Puesto</th><th style="' + th + ';text-align:center">Días</th><th style="' + th + ';text-align:center">Horas</th>' +
      '<th style="' + th + ';text-align:center">Tard.</th><th style="' + th + ';text-align:center">Sal. ant.</th><th style="' + th + ';text-align:center">Aus.</th><th style="' + th + ';text-align:center">S/ salida</th></tr>' +
      filas.map((x, i) => '<tr style="background:' + (i % 2 ? '#f7f8fa' : '#ffffff') + '"><td style="' + td + ';color:#7d8796">' + (i + 1) + '</td><td style="' + td + '"><b>' + e(x.nombre) + '</b><br><span style="color:#7d8796;font-size:7.5pt">' + x.ced + '</span></td>' +
        '<td style="' + td + ';font-size:7.8pt">' + e(PU[x.puesto] || x.puesto) + '</td><td style="' + td + ';text-align:center">' + x.asis + '/' + x.prog + '</td><td style="' + td + ';text-align:center"><b>' + x.horas + '</b></td>' +
        ['tard', 'salAnt', 'aus', 'inc'].map(c => '<td style="' + td + ';text-align:center">' + num(x[c]) + '</td>').join('') + '</tr>').join('') + '</table>';
  }
  h += '<div style="font-size:10pt;font-weight:bold;color:#3a475c;margin:16px 0 6px;text-transform:uppercase;letter-spacing:1px;' + (individual ? '' : 'page-break-before:always') + '">Detalle diario</div>';
  filas.forEach(x => {
    h += '<div class="blq" style="margin-bottom:12px">' + (individual ? '' : '<table><tr><td style="background:#46566f;color:#fff;padding:6px 9px;font-size:9pt"><b>' + e(x.nombre) + '</b> · ' + x.ced + ' · ' + e(PU[x.puesto] || x.puesto) +
      '</td><td style="background:#46566f;color:#dfe4eb;padding:6px 9px;font-size:8pt;text-align:right">Días ' + x.asis + '/' + x.prog + ' · Horas ' + x.horas + '</td></tr></table>') +
      '<table><tr><th style="' + th + ';background:#5b7190">Día</th><th style="' + th + ';background:#5b7190">Horario</th><th style="' + th + ';background:#5b7190">Marcas</th><th style="' + th + ';background:#5b7190;text-align:center">Horas</th><th style="' + th + ';background:#5b7190">Observaciones</th></tr>' +
      (x.dias || []).map((d, i) => '<tr style="background:' + (i % 2 ? '#f7f8fa' : '#ffffff') + '"><td style="' + td + '"><b>' + dias[i] + '</b><br><span style="color:#7d8796;font-size:7.5pt">' + fechaLarga_(d.fecha) + '</span></td>' +
        '<td style="' + td + ';font-size:8pt">' + e(d.horario) + '</td><td style="' + td + ';font-size:8pt">' + (d.marcas.length ? d.marcas.map(m => (m.charAt(0) === 'E' ? 'Entrada ' : 'Salida ') + m.slice(2)).join('<br>') : '<span style="color:#9aa3b5">—</span>') + '</td>' +
        '<td style="' + td + ';text-align:center">' + (d.horas ? hhmm_(d.horas) : '—') + '</td><td style="' + td + ';font-size:8pt">' +
        d.obs.map(o => '<span style="background:' + (/Tardía|Ausente|Sin salida|anticipada/.test(o) ? '#f8eedb;color:#7d5517' : '#e6ecf5;color:#3a475c') + ';padding:1px 6px;border-radius:8px;margin-right:3px">' + e(o) + '</span>').join(' ') + '</td></tr>').join('') +
      '</table></div>';
  });
  h += '<table class="blq" style="margin-top:26px"><tr>' + (individual ? ['Firma del funcionario', 'Visto bueno · Dirección'] : ['Elaborado por', 'Revisado por · Dirección']).map(t =>
    '<td style="width:50%;padding:30px 18px 0;text-align:center;font-size:8.5pt;color:#55606f"><div style="border-top:1px solid #7d8796;padding-top:5px">' + t + '</div></td>').join('') + '</tr></table>' +
    '<div style="margin-top:18px;font-size:7pt;color:#9aa3b5;text-align:center">Documento generado automáticamente por CTPL Asistencia. Las marcas manuales y anulaciones constan en la bitácora del sistema.</div>' +
    '</body></html>';
  return h;
}

/* ===================== API: ADMINISTRACIÓN ===================== */

function asisAdminDatos(token) {
  verificar_(token);
  const cfg = cfg_(), per = mapaPersonal_(), f = hoy_();
  const cu = {};
  cuentas_().v.forEach(r => { cu[limpiarCedula_(r[CU.CED])] = r; });
  const hoyM = leerMarcas_(f, f, 3000).filter(m => m.estado !== 'ANULADA');
  const lista = Object.keys(per).map(ced => {
    const p = per[ced], r = cu[ced], mis = hoyM.filter(m => m.ced === ced).sort((a, b) => a.hora < b.hora ? -1 : 1);
    const ent = mis.find(m => m.tipo === 'ENTRADA'), ult = mis[mis.length - 1];
    const sal = mis.slice().reverse().find(m => m.tipo === 'SALIDA');
    let hor = {}; try { hor = JSON.parse(r ? r[CU.HOR] || '{}' : '{}'); } catch (e) {}
    return {
      ced: ced, nombre: p.nombre, puesto: p.puesto, fotoId: p.fotoId,
      estado: r ? r[CU.EST] : 'SIN_CUENTA', correo: r ? r[CU.COR] : '', disp: r ? r[CU.DINF] : '', dispDesde: r ? r[CU.DFEC] : '',
      pend: r && (r[CU.PEST] === 'PENDIENTE' || r[CU.PEST] === 'APROBADO') ? { estado: r[CU.PEST], info: r[CU.PINF] } : null,
      verificado: !!(r && r[CU.DKEY] && r[CU.HASH]),
      codigo: r && r[CU.EST] === 'APROBADA' ? r[CU.COD] : '', fsol: r ? r[CU.FSOL] : '', ult: r ? r[CU.ULT] : '',
      auto: r && r[CU.AUTO] ? Utilities.formatDate(new Date(Number(r[CU.AUTO])), TZ, 'dd/MM/yyyy') : '',
      horario: hor,
      hoy: { entrada: ent ? ent.hora.slice(0, 5) : '', salida: sal ? sal.hora.slice(0, 5) : '', dentro: !!(ult && ult.tipo === 'ENTRADA'), n: mis.length }
    };
  }).sort((a, b) => a.nombre.localeCompare(b.nombre));
  // cuentas cuya cédula ya no está en Personal
  Object.keys(cu).filter(c => !per[c]).forEach(c => lista.push({ ced: c, nombre: cu[c][CU.NOM], puesto: '', estado: cu[c][CU.EST],
    correo: cu[c][CU.COR], disp: cu[c][CU.DINF], pend: null, codigo: '', horario: {}, hoy: { entrada: '', salida: '', dentro: false, n: 0 } }));
  const pt = pantallas_().v.map(r => ({ id: r[0], nombre: r[1], estado: r[3], creada: r[4], pulso: r[5] }));
  return { personal: lista, pantallas: pt, cfg: cfg, hoy: f, lunes: lunesDe_(f) };
}

function asisVincularCorreo(token, cedula, correo) {
  verificar_(token);
  const ced = limpiarCedula_(cedula);
  correo = correoValido_(correo);
  const p = personal_(ced);
  return bloquear_(() => {
    const c = buscarCuenta_(ced);
    if (correoEnUso_(c.todas, correo, ced)) throw new Error('Ese correo ya está ligado a otra persona.');
    if (c.r && (c.r[CU.EST] === 'ACTIVA' || c.r[CU.EST] === 'BLOQUEADA' || (c.r[CU.EST] === 'PENDIENTE' && c.r[CU.DKEY]))) {
      c.r[CU.COR] = correo; guardarCuenta_(c); audit_('ADMIN', 'CAMBIO_CORREO', ced + ' ' + correo);
      return { ok: true, codigo: '', enviado: false };
    }
    const r = c.r || new Array(C_CUENTAS.length).fill('');
    r[CU.CED] = ced; r[CU.NOM] = (p && p.nombre) || r[CU.NOM]; r[CU.COR] = correo; r[CU.EST] = 'APROBADA';
    r[CU.COD] = codigo6_(); r[CU.CVEN] = String(Date.now() + 72 * 3600e3);
    if (!r[CU.FSOL]) r[CU.FSOL] = Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm');
    if (c.r) { c.r = r; guardarCuenta_(c); } else c.sh.appendRow(r);
    const env = correoCodigo_(correo, r[CU.NOM], r[CU.COD], 'activación');
    audit_('ADMIN', 'APROBACION_CORREO', ced + ' ' + correo);
    return { ok: true, codigo: r[CU.COD], enviado: env };
  });
}

function asisResolverSolicitud(token, cedula, aprobar) {
  verificar_(token);
  const ced = limpiarCedula_(cedula);
  const c = buscarCuenta_(ced);
  if (!c.r) throw new Error('Solicitud no encontrada.');
  if (aprobar && c.r[CU.EST] === 'PENDIENTE' && c.r[CU.HASH] && c.r[CU.DKEY]) {   // registro unificado: ya verificó su correo
    c.r[CU.EST] = 'ACTIVA'; guardarCuenta_(c); audit_('ADMIN', 'APROBACION', ced);
    return { ok: true, activada: true };
  }
  if (aprobar) return asisVincularCorreo(token, ced, c.r[CU.COR]);
  c.r[CU.EST] = 'RECHAZADA'; guardarCuenta_(c); audit_('ADMIN', 'RECHAZO', ced);
  return { ok: true };
}

function asisResolverDispositivo(token, cedula, aprobar) {
  verificar_(token);
  return bloquear_(() => {
    const c = buscarCuenta_(limpiarCedula_(cedula));
    if (!c.r || !c.r[CU.PID]) throw new Error('No hay solicitud de dispositivo.');
    if (aprobar) { c.r[CU.PEST] = 'APROBADO'; c.r[CU.DID] = ''; c.r[CU.DKEY] = ''; }   // el dispositivo anterior queda revocado
    else { c.r[CU.PID] = ''; c.r[CU.PINF] = ''; c.r[CU.PEST] = ''; }
    guardarCuenta_(c); audit_('ADMIN', aprobar ? 'APRUEBA_DISPOSITIVO' : 'RECHAZA_DISPOSITIVO', c.r[CU.CED]);
    return true;
  });
}

function asisLiberarDispositivo(token, cedula) {
  verificar_(token);
  return bloquear_(() => {
    const c = buscarCuenta_(limpiarCedula_(cedula));
    if (!c.r) throw new Error('Cuenta no encontrada.');
    c.r[CU.DID] = ''; c.r[CU.DKEY] = ''; c.r[CU.DINF] = ''; c.r[CU.PID] = ''; c.r[CU.PINF] = ''; c.r[CU.PEST] = '';
    guardarCuenta_(c); audit_('ADMIN', 'LIBERA_DISPOSITIVO', c.r[CU.CED]);
    return true;
  });
}

function asisEstadoCuenta(token, cedula, estado) {
  verificar_(token);
  if (['ACTIVA', 'BLOQUEADA'].indexOf(estado) < 0) throw new Error('Estado no válido.');
  const c = buscarCuenta_(limpiarCedula_(cedula));
  if (!c.r || !c.r[CU.HASH]) throw new Error('La cuenta aún no ha sido activada.');
  c.r[CU.EST] = estado; guardarCuenta_(c); audit_('ADMIN', estado, c.r[CU.CED]);
  return true;
}

function asisHorario(token, cedula, hor) {
  verificar_(token);
  const c = buscarCuenta_(limpiarCedula_(cedula));
  if (!c.r) throw new Error('La persona aún no tiene cuenta.');
  const h = {}, dias = ['', 'lunes', 'martes', 'miércoles', 'jueves', 'viernes'], cfg = cfg_();
  [1, 2, 3, 4, 5].forEach(d => {
    const v = String((hor || {})[d] || '').replace(/\s/g, '');
    if (!v) return;
    if (v === '-') { h[d] = '-'; return; }
    let tr;
    try { tr = tramos_(v, cfg); } catch (e) { throw new Error('Horario del ' + dias[d] + ' no válido. Use 07:00-16:20, o dos jornadas: 07:00-12:10,17:30-21:45'); }
    if (tr.length > 3) throw new Error('Máximo 3 jornadas por día (' + dias[d] + ').');
    tr.forEach((t, i) => {
      if (min_(t.sal) <= min_(t.ent)) throw new Error('En ' + dias[d] + ' la salida debe ser después de la entrada (' + t.ent + '-' + t.sal + ').');
      if (i && min_(t.ent) < min_(tr[i - 1].sal)) throw new Error('En ' + dias[d] + ' las jornadas se traslapan.');
    });
    h[d] = tr.map(t => t.ent + '-' + t.sal).join(',');
  });
  c.r[CU.HOR] = JSON.stringify(h); guardarCuenta_(c);
  return true;
}

function asisMarcas(token, desde, hasta) {
  verificar_(token);
  return leerMarcas_(desde, hasta).reverse().slice(0, 4000);
}

function asisMarcaManual(token, d) {
  verificar_(token);
  const ced = limpiarCedula_(d.ced);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(d.fecha) || !/^\d{2}:\d{2}$/.test(d.hora)) throw new Error('Fecha u hora no válidas.');
  if (['ENTRADA', 'SALIDA'].indexOf(d.tipo) < 0) throw new Error('Tipo no válido.');
  if (!String(d.motivo || '').trim()) throw new Error('Indique el motivo de la marca manual.');
  const c = buscarCuenta_(ced), p = personal_(ced);
  const nombre = (c.r && c.r[CU.NOM]) || (p && p.nombre);
  if (!nombre) throw new Error('Cédula no encontrada.');
  sh_(H_MARCAS).appendRow(['A' + Date.now().toString(36).toUpperCase(), d.fecha.split('-').reverse().join('/') + ' ' + d.hora + ':00',
    d.fecha, d.hora + ':00', ced, nombre, d.tipo, 'ADMIN', '', '', '', 'MANUAL', 'VÁLIDA', String(d.motivo).trim(), 'ADMIN']);
  limpiarCacheMarcas_(ced, d.fecha); audit_('ADMIN', 'MARCA_MANUAL', ced + ' ' + d.fecha + ' ' + d.hora + ' ' + d.tipo);
  return true;
}

function asisAnularMarca(token, id, motivo) {
  verificar_(token);
  if (!String(motivo || '').trim()) throw new Error('Indique el motivo.');
  const sh = sh_(H_MARCAS), n = sh.getLastRow();
  const ids = sh.getRange(2, 1, n - 1, 1).getDisplayValues();
  for (let i = ids.length - 1; i >= 0; i--) if (ids[i][0] === id) {
    const r = sh.getRange(i + 2, 1, 1, C_MARCAS.length).getDisplayValues()[0];
    sh.getRange(i + 2, 13, 1, 2).setValues([['ANULADA', (r[13] ? r[13] + ' · ' : '') + 'Anulada: ' + motivo]]);
    limpiarCacheMarcas_(limpiarCedula_(r[4]), r[2]); audit_('ADMIN', 'ANULA_MARCA', id + ' ' + motivo);
    return true;
  }
  throw new Error('Marca no encontrada.');
}

function asisInforme(token, lunes, generar, enviar) {
  verificar_(token);
  lunes = lunesDe_(lunes);
  let info = null;
  if (generar) info = generarInforme_(lunes, !!enviar);
  const sh = sh_(H_INFORMES), n = sh.getLastRow();
  let filas = n < 2 ? [] : sh.getRange(2, 1, n - 1, C_INFORMES.length).getDisplayValues().filter(r => r[0] === lunes).map(informeFila_);
  const guardado = filas.length > 0;
  if (!guardado) filas = calcularSemana_(lunes);
  return { lunes: lunes, filas: filas, guardado: guardado, generado: guardado ? filas[0].generado : '', info: info };
}

function asisExportar(token, tipo, desde, hasta, formato) {
  verificar_(token);
  if (tipo === 'pdf') {
    const lun = lunesDe_(desde), b = pdfInforme_(asisInforme(token, lun).filas, lun, false);
    return { nombre: b.getName(), mime: 'application/pdf', b64: Utilities.base64Encode(b.getBytes()) };
  }
  if (tipo === 'marcas') return empaquetar_('Marcas_CTPL_' + desde + '_' + hasta, csvMarcas_(leerMarcas_(desde, hasta)), formato);
  const lunes = lunesDe_(desde), r = asisInforme(token, lunes).filas;
  if (tipo === 'detalle') return empaquetar_('Asistencia_CTPL_' + lunes + '_detalle', csvDetalle_(r), formato);
  return empaquetar_('Asistencia_CTPL_' + lunes + '_resumen', csvResumen_(r), formato);
}

function asisGuardarConfig(token, d) {
  verificar_(token);
  const c = cfg_(), hh = /^\d{1,2}:\d{2}$/;
  if (d.entrada != null) { if (!hh.test(d.entrada)) throw new Error('Hora de entrada no válida.'); c.entrada = d.entrada; }
  if (d.salida != null) { if (!hh.test(d.salida)) throw new Error('Hora de salida no válida.'); c.salida = d.salida; }
  if (d.tolerancia != null) c.tolerancia = Math.max(0, Math.min(60, +d.tolerancia || 0));
  if (d.slot != null) c.slot = Math.max(10, Math.min(60, +d.slot || 15));
  if (d.separacion != null) c.separacion = Math.max(1, Math.min(120, +d.separacion || 3));
  if (d.correos != null) c.correos = String(d.correos).split(/[,;\s]+/).filter(Boolean).map(correoValido_).join(',');
  if (d.geo != null) c.geo = !!d.geo;
  if (d.lat != null) c.lat = +d.lat || 0;
  if (d.lng != null) c.lng = +d.lng || 0;
  if (d.radio != null) c.radio = Math.max(50, Math.min(2000, +d.radio || 250));
  if (d.noLectivos != null) c.noLectivos = String(d.noLectivos).split(/[,\s]+/).filter(f => /^\d{4}-\d{2}-\d{2}$/.test(f)).join(',');
  if (d.urlApp != null) c.urlApp = String(d.urlApp).trim().slice(0, 200);
  if (d.autoDias != null) c.autoDias = Math.max(0, Math.min(365, Math.round(+d.autoDias || 0)));
  if (d.institucion) c.institucion = String(d.institucion).slice(0, 80);
  if (c.geo && (!c.lat || !c.lng)) throw new Error('Para activar la geocerca indique la ubicación del colegio.');
  PropertiesService.getScriptProperties().setProperty('ASIS_CFG', JSON.stringify(c)); CFG_MEM_ = null;
  audit_('ADMIN', 'CONFIG', '');
  return c;
}

function asisVentana(token, horas) {
  verificar_(token);
  const c = cfg_();
  c.ventanaHasta = horas > 0 ? Date.now() + Math.min(72, +horas) * 3600e3 : 0;
  PropertiesService.getScriptProperties().setProperty('ASIS_CFG', JSON.stringify(c)); CFG_MEM_ = null;
  audit_('ADMIN', 'VENTANA_VINCULACION', String(horas));
  return c.ventanaHasta;
}

function asisPantalla(token, kid, estado) {
  verificar_(token);
  const t = pantallas_(), i = t.v.findIndex(r => r[0] === kid);
  if (i < 0) throw new Error('Pantalla no encontrada.');
  t.sh.getRange(i + 2, 4).setValue(estado === 'ACTIVA' ? 'ACTIVA' : 'REVOCADA');
  audit_('ADMIN', 'PANTALLA_' + estado, kid);
  return true;
}
