/**
 * ==============================================================
 *  📌 Nombre del Proyecto: Exportador de Eventos de Google Calendar a Google Sheets
 *  📌 Descripción:
 *     Este script extrae eventos de los últimos 120 días del calendario predeterminado
 *     del usuario y los registra en una hoja de cálculo de Google Sheets. El archivo 
 *     se almacena en la raíz de Google Drive con un nombre basado en el correo del usuario 
 *     y la fecha de extracción. También añade un resumen de la extracción al final del documento.
 *
 *  📌 Funcionalidades:
 *     ✅ Extrae eventos desde 120 días atrás hasta la fecha actual.
 *     ✅ Crea automáticamente un Google Sheets en la raíz de Google Drive.
 *     ✅ Registra detalles de los eventos como título, fecha, ubicación y descripción.
 *     ✅ Agrega una leyenda con el nombre del calendario y el rango de fechas extraídas.
 *     ✅ Notifica el éxito del proceso y proporciona un enlace al archivo generado.
 *
 * ==============================================================
 *
 *  📝 CONTROL DE CAMBIOS Y VERSIONES
 *  --------------------------------------------------------------
 *  📅 Fecha        | 🔢 Versión | 📌 Descripción del Cambio
 *  --------------------------------------------------------------
 *  2024-01-29     | 1.0       | Versión inicial del script con extracción básica de eventos.
 *  2024-01-30     | 1.1       | Se agregaron encabezados en la hoja de cálculo y mejoras en el formato.
 *  2024-01-31     | 1.2       | Se incluyó control de cambios y mejoras en la estructura del código.
 *  2026-10-13     | 2.0       | Refactor a Web App "Gtoolbox": doGet, funciones especializadas y módulo Reporte de Actividades.
 *  --------------------------------------------------------------
 */

/**
 * Punto de entrada de la Web App Gtoolbox.
 * @param {Object} e Parámetros del evento GET.
 * @return {HtmlOutput}
 */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Gtoolbox')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Extrae eventos del calendario predeterminado entre dos fechas (inclusive).
 * @param {string} startDateStr Fecha inicio 'yyyy-MM-dd'.
 * @param {string} endDateStr Fecha fin 'yyyy-MM-dd'.
 * @return {{calendarName: string, start: Date, end: Date, events: Object[]}}
 */
function getCalendarEvents_(startDateStr, endDateStr) {
  const start = parseDate_(startDateStr);
  const end = parseDate_(endDateStr);
  if (!start || !end) throw new Error('Fechas inválidas.');
  if (start > end) throw new Error('La fecha de inicio no puede ser posterior a la fecha de fin.');
  end.setHours(23, 59, 59, 999);

  const calendar = CalendarApp.getDefaultCalendar();
  const tz = Session.getScriptTimeZone();
  const events = calendar.getEvents(start, end).map(function (event) {
    return {
      titulo: event.getTitle(),
      inicio: Utilities.formatDate(event.getStartTime(), tz, "yyyy-MM-dd HH:mm:ss"),
      fin: Utilities.formatDate(event.getEndTime(), tz, "yyyy-MM-dd HH:mm:ss"),
      ubicacion: event.getLocation() || 'Sin ubicación',
      descripcion: event.getDescription() || 'Sin descripción',
      id: event.getId()
    };
  });
  return { calendarName: calendar.getName(), start: start, end: end, events: events };
}

/**
 * Crea un Google Sheet en la raíz de Drive con los eventos.
 * @param {Object[]} events Eventos serializables.
 * @param {string} calendarName Nombre del calendario.
 * @param {string} startDateStr Fecha inicio.
 * @param {string} endDateStr Fecha fin.
 * @return {{url: string, fileName: string}}
 */
function createSheet_(events, calendarName, startDateStr, endDateStr) {
  const tz = Session.getScriptTimeZone();
  const userEmail = Session.getActiveUser().getEmail();
  const timestamp = Utilities.formatDate(new Date(), tz, 'yyyyMMdd_HHmmss');
  const fileName = userEmail + '_' + timestamp;
  const ss = SpreadsheetApp.create(fileName);
  const sheet = ss.getActiveSheet();
  sheet.setName(Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd'));

  const rows = [['Título', 'Inicio', 'Fin', 'Ubicación', 'Descripción', 'ID del Evento']];
  events.forEach(function (ev) {
    rows.push([ev.titulo, ev.inicio, ev.fin, ev.ubicacion, ev.descripcion, ev.id]);
  });
  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  sheet.getRange(1, 1, 1, rows[0].length).setFontWeight('bold');

  const legendRow = sheet.getLastRow() + 3;
  sheet.getRange(legendRow, 1).setValue('Datos extraídos del calendario: ' + calendarName);
  sheet.getRange(legendRow + 1, 1).setValue('Rango de fechas: ' + startDateStr + ' a ' + endDateStr);
  return { url: ss.getUrl(), fileName: fileName };
}

/**
 * Envía el correo de notificación al usuario.
 * @param {string} fileName Nombre del archivo.
 * @param {string} fileUrl URL del Sheet.
 */
function sendReportEmail_(fileName, fileUrl) {
  MailApp.sendEmail({
    to: Session.getActiveUser().getEmail(),
    subject: 'Extracción de Eventos Completada',
    body: 'La extracción de eventos se ha completado exitosamente.\n\n📂 Archivo: ' + fileName +
      '\n📍 Ubicación: Raíz de Google Drive\n🔗 Accede aquí: ' + fileUrl
  });
}

/**
 * Función invocada desde el frontend (google.script.run).
 * Extrae eventos, crea el Sheet, envía el correo y devuelve los eventos.
 * @param {string} startDateStr 'yyyy-MM-dd'.
 * @param {string} endDateStr 'yyyy-MM-dd'.
 * @return {{events: Object[], sheetUrl: string, fileName: string, calendarName: string}}
 */
function generateActivityReport(startDateStr, endDateStr) {
  const data = getCalendarEvents_(startDateStr, endDateStr);
  const sheet = createSheet_(data.events, data.calendarName, startDateStr, endDateStr);
  sendReportEmail_(sheet.fileName, sheet.url);
  return {
    events: data.events,
    sheetUrl: sheet.url,
    fileName: sheet.fileName,
    calendarName: data.calendarName
  };
}

/** Convierte 'yyyy-MM-dd' en Date local; devuelve null si es inválida. */
function parseDate_(str) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(str || '');
  if (!m) return null;
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), 0, 0, 0, 0);
}
