/**
 * ACACIA MEZCAL — Google Apps Script
 * Deploy como Web App:
 *   - Ejecutar como: Yo (tu cuenta)
 *   - Acceso: Cualquier persona, incluso anónima
 *
 * Después de deployar, copia la URL de la Web App y pégala en index.html
 * donde dice: 'https://script.google.com/macros/s/REEMPLAZAR_APPS_SCRIPT_URL/exec'
 */

const SHEET_NAME = 'landing page';
const EMAIL_DESTINO = 'mezcal.acacia@gmail.com';

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    guardarLead(data);
    enviarNotificacion(data);
    return jsonOk();
  } catch (err) {
    return jsonOk(); // Siempre devolver 200 para evitar CORS errors en el front
  }
}

// Permite hacer GET para verificar que el script funciona
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', msg: 'ACACIA Apps Script activo' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function guardarLead(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);

  // Crear hoja si no existe
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['Fecha', 'Nombre', 'Email', 'WhatsApp', 'Interés', 'Mensaje']);
    // Formato encabezados
    const header = sheet.getRange(1, 1, 1, 6);
    header.setFontWeight('bold');
    header.setBackground('#c8922a');
    header.setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    sheet.setColumnWidths(1, 6, 160);
  }

  sheet.appendRow([
    data.fecha    || new Date().toLocaleString('es-MX'),
    data.nombre   || '',
    data.email    || '',
    data.telefono || '',
    data.interes  || '',
    data.mensaje  || ''
  ]);
}

function enviarNotificacion(data) {
  const asunto = `🌵 Nuevo lead ACACIA — ${data.interes || 'Sin especificar'}`;

  const cuerpo = `
NUEVO CONTACTO — ACACIA MEZCAL
================================

Fecha:     ${data.fecha || new Date().toLocaleString('es-MX')}
Nombre:    ${data.nombre || '—'}
Email:     ${data.email || '—'}
WhatsApp:  ${data.telefono || '—'}
Interés:   ${data.interes || '—'}

Mensaje:
${data.mensaje || '(sin mensaje)'}

================================
Responder a: ${data.email || '—'}
`.trim();

  MailApp.sendEmail({
    to: EMAIL_DESTINO,
    subject: asunto,
    body: cuerpo,
    replyTo: data.email || EMAIL_DESTINO
  });
}

function jsonOk() {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}
