/* =================================================================
   BACKEND DE LA INVITACIÓN DE BODA — Google Apps Script
   -----------------------------------------------------------------
   Cómo instalarlo:
   1. Abre (o crea) una Google Sheet donde se guardarán las respuestas.
   2. Extensiones > Apps Script.
   3. Borra el contenido de Code.gs y pega este archivo entero.
   4. Cambia TOKEN_ADMIN por tu clave (la misma que en panel-confirmaciones.html).
   5. Implementar > Nueva implementación > tipo "Aplicación web".
      - Ejecutar como: Yo
      - Quién tiene acceso: Cualquier usuario
      (Si ya tenías una implementación: Gestionar implementaciones > Editar > Nueva versión)
   6. Copia la URL (termina en /exec) y pégala como googleScriptUrl en:
      - invitacion-boda.html
      - panel-confirmaciones.html

   IMPORTANTE — cambio de esquema:
   Esta versión guarda los invitados de cada familia juntos en una sola
   fila (columna "Invitados" en formato JSON), en vez de un simple número
   de acompañantes. Si ya tenías filas de prueba con el formato antiguo,
   puedes borrarlas — no son compatibles con el panel nuevo.
   ================================================================= */

const TOKEN_ADMIN = "boda2026";                 // clave para leer y para guardar el plano de mesas
const NOMBRE_HOJA = "Confirmaciones";
const NOMBRE_HOJA_MESAS = "Mesas";

/* ---------------------------------------------------------------
   ENTRADA — recibe confirmaciones del formulario y guardados de mesas
   --------------------------------------------------------------- */
function doPost(e) {
  const datos = JSON.parse(e.postData.contents);

  if (datos.tipo === "mesas") {
    if (datos.token !== TOKEN_ADMIN) {
      return respuestaJSON({ ok: false, error: "No autorizado" });
    }
    guardarMesas(datos);
    return respuestaJSON({ ok: true });
  }

  // Confirmación de asistencia normal
  const hoja = obtenerHoja();
  hoja.appendRow([
    new Date(),
    datos.id || generarId(),
    datos.asistencia || "",
    JSON.stringify(datos.invitados || []),
    datos.intolerancias || "",
    datos.intolerancias_detalle || "",
    datos.cancion || "",
    datos.mensaje || ""
  ]);

  return respuestaJSON({ ok: true });
}

/* ---------------------------------------------------------------
   SALIDA — lee confirmaciones o el plano de mesas, según ?tipo=
   --------------------------------------------------------------- */
function doGet(e) {
  if (e.parameter.token !== TOKEN_ADMIN) {
    return respuestaJSON({ error: "No autorizado" });
  }

  if (e.parameter.tipo === "mesas") {
    return respuestaJSON(leerMesas());
  }

  const hoja = obtenerHoja();
  const filas = hoja.getDataRange().getValues();
  filas.shift(); // cabecera

  const confirmaciones = filas
    .filter(f => f[1])
    .map(f => {
      let invitados = [];
      try { invitados = JSON.parse(f[3]); } catch (err) { invitados = []; }
      return {
        fecha: f[0] instanceof Date ? f[0].toISOString() : String(f[0]),
        id: String(f[1]),
        asistencia: f[2],
        invitados: invitados,
        intolerancias: f[4],
        intolerancias_detalle: f[5],
        cancion: f[6],
        mensaje: f[7]
      };
    });

  return respuestaJSON(confirmaciones);
}

/* ---------------------------------------------------------------
   MESAS
   --------------------------------------------------------------- */
function guardarMesas(datos) {
  const hoja = obtenerHojaMesas();
  const plano = {
    mesas: datos.mesas || [],
    asignaciones: datos.asignaciones || {}
  };
  hoja.getRange(2, 1).setValue(JSON.stringify(plano));
}

function leerMesas() {
  const hoja = obtenerHojaMesas();
  const valor = hoja.getRange(2, 1).getValue();
  if (!valor) return { mesas: [], asignaciones: {} };
  try {
    return JSON.parse(valor);
  } catch (err) {
    return { mesas: [], asignaciones: {} };
  }
}

/* ---------------------------------------------------------------
   Utilidades
   --------------------------------------------------------------- */
function obtenerHoja() {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = libro.getSheetByName(NOMBRE_HOJA);
  if (!hoja) {
    hoja = libro.insertSheet(NOMBRE_HOJA);
    hoja.appendRow([
      "Fecha", "ID", "Asistencia", "Invitados (JSON)",
      "Intolerancias", "Detalle intolerancias", "Canción", "Mensaje"
    ]);
    hoja.setFrozenRows(1);
  }
  return hoja;
}

function obtenerHojaMesas() {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = libro.getSheetByName(NOMBRE_HOJA_MESAS);
  if (!hoja) {
    hoja = libro.insertSheet(NOMBRE_HOJA_MESAS);
    hoja.appendRow(["Plano de mesas (no editar a mano, se gestiona desde el panel)"]);
    hoja.appendRow([JSON.stringify({ mesas: [], asignaciones: {} })]);
  }
  return hoja;
}

function generarId() {
  return "inv_" + new Date().getTime() + "_" + Math.floor(Math.random() * 10000);
}

function respuestaJSON(objeto) {
  return ContentService
    .createTextOutput(JSON.stringify(objeto))
    .setMimeType(ContentService.MimeType.JSON);
}

/* =================================================================
   SOLO PARA PROBAR A MANO DESDE EL EDITOR
   Selecciona "probarEnvio" en el desplegable de arriba y pulsa Ejecutar.
   Simula un envío real y añade una fila de prueba a la hoja.
   ================================================================= */
function probarEnvio() {
  const eventoFalso = {
    postData: {
      contents: JSON.stringify({
        id: generarId(),
        asistencia: "si",
        invitados: ["Prueba de Manuel", "Acompañante de prueba"],
        intolerancias: "Gluten, Lactosa",
        intolerancias_detalle: "Esto es solo una prueba",
        cancion: "Perfect - Ed Sheeran",
        mensaje: "Fila de prueba, puedes borrarla"
      })
    }
  };
  const resultado = doPost(eventoFalso);
  Logger.log(resultado.getContent());
}
