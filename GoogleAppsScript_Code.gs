/**
 * ===================================================================
 * RECEPTOR DE CALIFICACIONES - MOODLE XML QUIZ & GOOGLE SHEETS
 * ===================================================================
 * Este script recibe automáticamente las respuestas y calificaciones
 * enviadas por los estudiantes y las registra ordenadamente en la hoja.
 * 
 * ORDEN EXACTO DE LAS COLUMNAS REGISTRADAS:
 * 1. Nombre del estudiante
 * 2. Matrícula / Código
 * 3. Fecha y hora de envío
 * 4. Asignatura
 * 5. Profesor
 * 6. Título / Cuestionario
 * 7. Calificación obtenida
 * 8. Calificación máxima
 * 9. Porcentaje (%)
 * ===================================================================
 */

// Encabezados oficiales para la hoja de cálculo
var ENCABEZADOS = [
  "Nombre del estudiante",
  "Matrícula / Código",
  "Fecha y hora de envío",
  "Asignatura",
  "Profesor",
  "Título / Cuestionario",
  "Calificación obtenida",
  "Calificación máxima",
  "Porcentaje (%)"
];

/**
 * Función principal que atiende las solicitudes POST enviadas por la app web.
 */
function doPost(e) {
  var lock = LockService.getScriptLock();
  // Esperar hasta 10 segundos para evitar colisiones si varios estudiantes envían a la vez
  try {
    lock.waitLock(10000);
  } catch (err) {
    return crearRespuestaJSON({
      status: "error",
      message: "El servidor estaba ocupado. Por favor, reintenta."
    });
  }

  try {
    var hoja = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    prepararHojaSiEstaVacia(hoja);

    // Obtener los datos JSON enviados en el cuerpo de la petición
    var datos = {};
    if (e && e.postData && e.postData.contents) {
      try {
        datos = JSON.parse(e.postData.contents);
      } catch (parseErr) {
        datos = e.parameter || {};
      }
    } else if (e && e.parameter) {
      datos = e.parameter;
    }

    var nombreEstudiante = datos.studentName || "Estudiante anónimo";
    var matricula = datos.matricula || datos.studentId || "S/M";
    var fechaEnvio = datos.submissionDate || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
    var asignatura = datos.subject || "";
    var profesor = datos.teacher || "";
    var tituloCuestionario = datos.quizTitle || "Cuestionario Moodle";
    var notaObtenida = Number(datos.scoreEarned != null ? datos.scoreEarned : 0);
    var notaMaxima = Number(datos.scoreMax != null ? datos.scoreMax : 0);
    var porcentaje = Number(datos.percentage != null ? datos.percentage : (notaMaxima > 0 ? (notaObtenida / notaMaxima) * 100 : 0));

    // Formatear porcentaje a 2 decimales con símbolo %
    var porcentajeTexto = porcentaje.toFixed(2) + "%";

    // Insertar la fila con el orden exacto solicitado (9 columnas)
    hoja.appendRow([
      nombreEstudiante,
      matricula,
      fechaEnvio,
      asignatura,
      profesor,
      tituloCuestionario,
      notaObtenida,
      notaMaxima,
      porcentajeTexto
    ]);

    // Resaltar visualmente según el porcentaje alcanzado (columna 9)
    var ultimaFila = hoja.getLastRow();
    var celdaPorcentaje = hoja.getRange(ultimaFila, 9);
    if (porcentaje >= 60) {
      celdaPorcentaje.setBackground("#dcfce7"); // Verde suave
      celdaPorcentaje.setFontColor("#15803d");
    } else {
      celdaPorcentaje.setBackground("#fee2e2"); // Rojo suave
      celdaPorcentaje.setFontColor("#b91c1c");
    }

    return crearRespuestaJSON({
      status: "success",
      message: "Calificación registrada correctamente en Google Sheets.",
      fila: ultimaFila,
      estudiante: nombreEstudiante
    });

  } catch (error) {
    return crearRespuestaJSON({
      status: "error",
      message: "Error en el script: " + error.toString()
    });
  } finally {
    lock.releaseLock();
  }
}

/**
 * Función GET para verificar la conexión (Ping Test) o envíos de respaldo.
 */
function doGet(e) {
  // Si el docente prueba la conexión desde la app o navegador:
  if (!e || !e.parameter || !e.parameter.action) {
    return crearRespuestaJSON({
      status: "success",
      message: "Conexión exitosa con Google Apps Script. El webhook está listo para recibir calificaciones.",
      timestamp: new Date().toISOString()
    });
  }

  // Soporte de respaldo por GET si la red de la institución bloquea peticiones POST
  return doPost(e);
}

/**
 * Crea o formatea la fila de encabezados si la hoja está recién creada.
 */
function prepararHojaSiEstaVacia(hoja) {
  if (hoja.getLastRow() === 0) {
    hoja.appendRow(ENCABEZADOS);
    var rangoEncabezado = hoja.getRange(1, 1, 1, ENCABEZADOS.length);
    rangoEncabezado.setFontWeight("bold");
    rangoEncabezado.setBackground("#1e293b"); // Slate oscuro
    rangoEncabezado.setFontColor("#ffffff");
    rangoEncabezado.setHorizontalAlignment("center");
    hoja.setFrozenRows(1);
    for (var col = 1; col <= ENCABEZADOS.length; col++) {
      hoja.autoResizeColumn(col);
    }
  }
}

/**
 * Retorna una respuesta JSON compatible con CORS para navegadores web.
 */
function crearRespuestaJSON(objetoRespuesta) {
  return ContentService
    .createTextOutput(JSON.stringify(objetoRespuesta))
    .setMimeType(ContentService.MimeType.JSON);
}
