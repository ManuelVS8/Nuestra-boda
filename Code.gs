/* =================================================================
   BACKEND DE LA INVITACIÓN DE BODA — Google Apps Script
   -----------------------------------------------------------------
   Cómo instalarlo:
   1. Abre (o crea) una Google Sheet donde se guardarán las respuestas.
   2. Extensiones > Apps Script.
   3. Borra el contenido de Code.gs y pega este archivo entero.
   4. Cambia TOKEN_ADMIN por una clave tuya (letras y números, sin espacios).
   5. Implementar > Nueva implementación > tipo "Aplicación web".
      - Ejecutar como: Yo
      - Quién tiene acceso: Cualquier usuario
   6. Copia la URL que te da (termina en /exec). Esa es tu googleScriptUrl:
      - Pégala en CONFIG.googleScriptUrl de invitacion-boda.html
      - Pégala en CONFIG.googleScriptUrl de panel-confirmaciones.html
      - Pon el mismo TOKEN_ADMIN en CONFIG.token de panel-confirmaciones.html
   ================================================================= */

const TOKEN_ADMIN = "BodaSM-MZlmff4QHL";     // clave para poder LEER las confirmaciones
const NOMBRE_HOJA = "Confirmaciones";

/* Recibe las confirmaciones que manda el formulario de la invitación */
function doPost(e) {
  const hoja = obtenerHoja();
  const datos = JSON.parse(e.postData.contents);

  hoja.appendRow([
    new Date(),
    datos.nombre || "",
    datos.asistencia || "",
    datos.acompanantes || "",
    datos.intolerancias || "",
    datos.intolerancias_detalle || "",
    datos.cancion || "",
    datos.mensaje || ""
  ]);

  return ContentService
    .createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

/* Devuelve todas las confirmaciones en JSON, solo si el token es correcto */
function doGet(e) {
  const tokenRecibido = e.parameter.token;

  if (tokenRecibido !== TOKEN_ADMIN) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: "No autorizado" }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const hoja = obtenerHoja();
  const filas = hoja.getDataRange().getValues();
  filas.shift(); // quitamos la fila de cabeceras

  const confirmaciones = filas
    .filter(f => f[1]) // ignora filas vacías
    .map(f => ({
      fecha: f[0] instanceof Date ? f[0].toISOString() : String(f[0]),
      nombre: f[1],
      asistencia: f[2],
      acompanantes: f[3],
      intolerancias: f[4],
      intolerancias_detalle: f[5],
      cancion: f[6],
      mensaje: f[7]
    }));

  return ContentService
    .createTextOutput(JSON.stringify(confirmaciones))
    .setMimeType(ContentService.MimeType.JSON);
}

/* Crea la hoja "Confirmaciones" con cabeceras si todavía no existe */
function obtenerHoja() {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = libro.getSheetByName(NOMBRE_HOJA);
  if (!hoja) {
    hoja = libro.insertSheet(NOMBRE_HOJA);
    hoja.appendRow([
      "Fecha", "Nombre", "Asistencia", "Acompañantes",
      "Intolerancias", "Detalle intolerancias", "Canción", "Mensaje"
    ]);
    hoja.setFrozenRows(1);
  }
  return hoja;
}
