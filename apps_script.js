/**
 * ACACIA MEZCAL — Google Apps Script
 * ====================================
 * Google Sheet: https://docs.google.com/spreadsheets/d/1S3pLVs8JkuuCAJX0-BTM1OEyH6X5lMe5EX6swUARWaI
 *
 * CÓMO DEPLOYAR (paso a paso):
 * 1. Abre el Google Sheet → menú "Extensiones" → "Apps Script"
 * 2. Borra todo el código que aparece por defecto
 * 3. Pega este archivo completo
 * 4. Guarda con Ctrl+S → pon nombre "ACACIA Landing"
 * 5. Clic en "Implementar" → "Nueva implementación"
 * 6. Tipo: "Aplicación web"
 * 7. Descripción: "ACACIA v1"
 * 8. Ejecutar como: "Yo (tu correo)"
 * 9. Quién tiene acceso: "Cualquier persona"
 * 10. Clic "Implementar" → autoriza permisos
 * 11. Copia la URL que aparece (termina en /exec)
 * 12. En index.html busca "REEMPLAZAR_APPS_SCRIPT_URL" y pégala
 */

const SHEET_ID   = '1S3pLVs8JkuuCAJX0-BTM1OEyH6X5lMe5EX6swUARWaI';
const SHEET_NAME = 'landing page';
const EMAIL_DEST = 'mezcal.acacia@gmail.com';

// ── Colores ACACIA ──────────────────────────────────────────────
const C_HEADER_BG   = '#1a0f00';
const C_HEADER_TEXT = '#c8922a';
const C_ROW_ODD     = '#ffffff';
const C_ROW_EVEN    = '#fdf5e8';
const C_BORDER      = '#d4b896';
const C_DATE        = '#7a6a52';
const C_ACCENT      = '#c8922a';

// ── Ancho de columnas (px) ──────────────────────────────────────
const COL_WIDTHS = [160, 160, 220, 140, 180, 300];

// ================================================================
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    guardarLead(data);
    enviarNotificacion(data);
  } catch (err) {
    Logger.log('Error: ' + err.message);
  }
  // Siempre 200 para evitar errores CORS en el cliente
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// Permite GET para verificar que el script está activo
function doGet() {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', app: 'ACACIA Landing Page' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ================================================================
function guardarLead(data) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  let sheet   = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    crearEncabezados(sheet);
  }

  const fecha = data.fecha || Utilities.formatDate(new Date(), 'America/Mexico_City', 'dd/MM/yyyy HH:mm');
  const fila  = [
    fecha,
    data.nombre   || '',
    data.email    || '',
    data.telefono || '',
    data.interes  || '',
    data.mensaje  || ''
  ];

  sheet.appendRow(fila);
  const rowNum = sheet.getLastRow();
  formatearFila(sheet, rowNum);

  // Ajustar altura de fila si hay mensaje largo
  if (data.mensaje && data.mensaje.length > 80) {
    sheet.setRowHeight(rowNum, 72);
  }
}

// ── Encabezados con formato ACACIA ──────────────────────────────
function crearEncabezados(sheet) {
  const headers = ['Fecha', 'Nombre', 'Email', 'WhatsApp', 'Interés', 'Mensaje'];
  sheet.appendRow(headers);

  const rng = sheet.getRange(1, 1, 1, headers.length);
  rng.setBackground(C_HEADER_BG)
     .setFontColor(C_HEADER_TEXT)
     .setFontWeight('bold')
     .setFontFamily('Arial')
     .setFontSize(10)
     .setHorizontalAlignment('center')
     .setVerticalAlignment('middle')
     .setBorder(true, true, true, true, true, true, C_BORDER, SpreadsheetApp.BorderStyle.SOLID);

  sheet.setRowHeight(1, 36);
  sheet.setFrozenRows(1);

  // Anchos de columna
  COL_WIDTHS.forEach((w, i) => sheet.setColumnWidth(i + 1, w));

  // Columna Mensaje: wrap text
  sheet.getRange(2, 6, 1000, 1).setWrap(true);

  // Columna Fecha: formato centrado
  sheet.getRange(2, 1, 1000, 1).setHorizontalAlignment('center').setFontColor(C_DATE);

  // Email: azul con subrayado
  sheet.getRange(2, 3, 1000, 1).setFontColor('#1155cc');
}

// ── Formato por fila (filas pares/impares alternadas) ──────────
function formatearFila(sheet, rowNum) {
  const rng = sheet.getRange(rowNum, 1, 1, 6);
  const bg  = rowNum % 2 === 0 ? C_ROW_EVEN : C_ROW_ODD;

  rng.setBackground(bg)
     .setFontFamily('Arial')
     .setFontSize(10)
     .setVerticalAlignment('middle')
     .setBorder(false, false, true, false, false, false,
                C_BORDER, SpreadsheetApp.BorderStyle.SOLID_THIN);

  // Interés: gold + bold
  sheet.getRange(rowNum, 5)
       .setFontColor(C_ACCENT)
       .setFontWeight('bold');

  // Fecha: centrada y gris
  sheet.getRange(rowNum, 1)
       .setHorizontalAlignment('center')
       .setFontColor(C_DATE);

  // Email: link style
  sheet.getRange(rowNum, 3)
       .setFontColor('#1155cc');

  // Mensaje: texto normal ajustado
  sheet.getRange(rowNum, 6)
       .setFontColor('#333333')
       .setWrap(true);
}

// ── Email de notificación ───────────────────────────────────────
function enviarNotificacion(data) {
  const asunto = `🌵 ACACIA — Nuevo lead: ${data.interes || 'Sin especificar'} · ${data.nombre || '?'}`;

  const html = `
<div style="font-family:Arial,sans-serif;max-width:560px;background:#fdf5e8;padding:24px;border-radius:8px;border:1px solid #d4b896">
  <h2 style="color:#c8922a;margin:0 0 16px;font-size:18px">🌵 Nuevo contacto — ACACIA Mezcal</h2>
  <table style="width:100%;border-collapse:collapse;font-size:14px">
    <tr><td style="padding:8px 12px;background:#1a0f00;color:#c8922a;font-weight:bold;width:110px">Nombre</td>
        <td style="padding:8px 12px;background:#fff;border-bottom:1px solid #e8d5b0">${data.nombre || '—'}</td></tr>
    <tr><td style="padding:8px 12px;background:#1a0f00;color:#c8922a;font-weight:bold">Email</td>
        <td style="padding:8px 12px;background:#fdf5e8;border-bottom:1px solid #e8d5b0">
          <a href="mailto:${data.email}" style="color:#1155cc">${data.email || '—'}</a></td></tr>
    <tr><td style="padding:8px 12px;background:#1a0f00;color:#c8922a;font-weight:bold">WhatsApp</td>
        <td style="padding:8px 12px;background:#fff;border-bottom:1px solid #e8d5b0">${data.telefono || '—'}</td></tr>
    <tr><td style="padding:8px 12px;background:#1a0f00;color:#c8922a;font-weight:bold">Interés</td>
        <td style="padding:8px 12px;background:#fdf5e8;border-bottom:1px solid #e8d5b0;font-weight:bold;color:#c8922a">${data.interes || '—'}</td></tr>
    <tr><td style="padding:8px 12px;background:#1a0f00;color:#c8922a;font-weight:bold;vertical-align:top">Mensaje</td>
        <td style="padding:8px 12px;background:#fff;line-height:1.6">${(data.mensaje || '(sin mensaje)').replace(/\n/g,'<br>')}</td></tr>
    <tr><td style="padding:8px 12px;background:#1a0f00;color:#c8922a;font-weight:bold">Fecha</td>
        <td style="padding:8px 12px;background:#fdf5e8;color:#7a6a52">${data.fecha || '—'}</td></tr>
  </table>
  <p style="margin:16px 0 0;font-size:12px;color:#7a6a52;text-align:center">
    Responder a: <a href="mailto:${data.email}" style="color:#c8922a">${data.email || '—'}</a>
  </p>
</div>`;

  MailApp.sendEmail({
    to:       EMAIL_DEST,
    subject:  asunto,
    htmlBody: html,
    replyTo:  data.email || EMAIL_DEST
  });
}
