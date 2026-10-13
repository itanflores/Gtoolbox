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
 *  --------------------------------------------------------------
 */

function exportEventsToSheet() {
  try {
    // Obtener la dirección de correo electrónico del usuario
    const userEmail = Session.getActiveUser().getEmail();

    // Obtener el calendario del usuario
    const calendar = CalendarApp.getDefaultCalendar();
    const calendarName = calendar.getName();

    // Definir el rango de tiempo: desde 120 días atrás hasta hoy
    const now = new Date();
    const pastDate = new Date();
    pastDate.setDate(now.getDate() - 2 );

    // Obtener eventos en el rango de fechas
    const events = calendar.getEvents(pastDate, now);

    // Crear un nuevo archivo de Google Sheets en la raíz de Google Drive
    const timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyyMMdd_HHmmss");
    const fileName = `${userEmail}_${timestamp}`;
    const sheet = SpreadsheetApp.create(fileName);
    const sheetId = sheet.getId();

    // Obtener la hoja activa y cambiar su nombre a la fecha de extracción
    const activeSheet = sheet.getActiveSheet();
    activeSheet.setName(Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd"));

    // Agregar encabezados a la hoja
    const headers = ["Título", "Inicio", "Fin", "Ubicación", "Descripción", "ID del Evento"];
    activeSheet.appendRow(headers);

    // Agregar eventos a la hoja de cálculo
    events.forEach(event => {
      activeSheet.appendRow([
        event.getTitle(),
        event.getStartTime(),
        event.getEndTime(),
        event.getLocation() || "Sin ubicación",
        event.getDescription() || "Sin descripción",
        event.getId()
      ]);
    });

    // Agregar una leyenda al final de la tabla
    const lastRow = activeSheet.getLastRow() + 3;
    activeSheet.getRange(lastRow, 1).setValue(`Datos extraídos del calendario: ${calendarName}`);
    activeSheet.getRange(lastRow + 1, 1).setValue(`Rango de fechas: ${Utilities.formatDate(pastDate, Session.getScriptTimeZone(), "yyyy-MM-dd")} a ${Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd")}`);

    // Obtener la URL del archivo creado
    const fileUrl = sheet.getUrl();

    // Notificar éxito en la ejecución
    Logger.log(`Extracción completada exitosamente.`);
    Logger.log(`Nombre del archivo: ${fileName}`);
    Logger.log(`Ubicación: Raíz de Google Drive`);
    Logger.log(`Enlace de acceso: ${fileUrl}`);

    // Notificación por correo en lugar de alert()
    MailApp.sendEmail({
      to: userEmail,
      subject: "Extracción de Eventos Completada",
      body: `La extracción de eventos se ha completado exitosamente.\n\n📂 Archivo: ${fileName}\n📍 Ubicación: Raíz de Google Drive\n🔗 Accede aquí: ${fileUrl}`
    });

  } catch (error) {
    // Capturar y mostrar errores en caso de fallo
    Logger.log(`Error: ${error.message}`);

    // Enviar error por correo
    MailApp.sendEmail({
      to: Session.getActiveUser().getEmail(),
      subject: "Error en la Extracción de Eventos",
      body: `Ocurrió un error al extraer eventos del calendario.\n\n❌ Error: ${error.message}`
    });
  }
}
