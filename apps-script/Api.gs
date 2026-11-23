/**
 * API JSON para el sitio de GitHub Pages.
 * El navegador envía POST {fn, args} con Content-Type text/plain (sin preflight CORS).
 * Solo se pueden llamar las funciones de esta lista; las de administración validan su token.
 */
const API_FUNCIONES = [
  // carnet y cuenta
  'asisInfo', 'obtenerPlantilla', 'estadoRegistro', 'registroCompleto', 'verificarRegistro', 'reenviarCodigoRegistro',
  'activarCuenta', 'iniciarSesion', 'estadoCuenta', 'perfil', 'marcar', 'misMarcas', 'misInformes', 'miExportacion', 'miInformePdf',
  'cambiarClave', 'recuperarClave', 'restablecerClave', 'solicitarRevinculo', 'confirmarRevinculo',
  // pantalla de marcas
  'kioscoRegistrar', 'kioscoPulso',
  // administración
  'adminLogin', 'adminListar', 'adminFotos', 'adminGuardar', 'adminEliminar', 'adminMarcar', 'adminRegistroAbierto',
  'adminCambiarClave', 'adminSubirPlantilla', 'asisAdminDatos', 'asisVincularCorreo', 'asisResolverSolicitud',
  'asisResolverDispositivo', 'asisLiberarDispositivo', 'asisEstadoCuenta', 'asisHorario', 'asisMarcas', 'asisMarcaManual',
  'asisAnularMarca', 'asisInforme', 'asisExportar', 'asisGuardarConfig', 'asisVentana', 'asisPantalla'
];

function doPost(e) {
  let out;
  try {
    const q = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (API_FUNCIONES.indexOf(q.fn) < 0) throw new Error('Operación no permitida.');
    const r = globalThis[q.fn].apply(null, Array.isArray(q.args) ? q.args : []);
    out = { ok: true, r: r === undefined ? null : r };
  } catch (err) {
    out = { ok: false, error: (err && err.message) || String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}
