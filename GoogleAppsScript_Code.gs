/**
 * ===================================================================
 * RECEPTOR DE CALIFICACIONES Y SINCRONIZACIÓN EN LA NUBE (MULTI-DISPOSITIVO)
 * ===================================================================
 * Este script realiza 3 funciones clave:
 * 1. Recibe automáticamente las respuestas enviadas por estudiantes.
 * 2. Permite sincronizar las calificaciones en la nube para consultarlas
 *    desde cualquier computadora, teléfono o tableta.
 * 3. Permite guardar y descargar el cuestionario Moodle activo en la hoja
 *    para abrirlo y editarlo en múltiples dispositivos sin reenviar archivos.
 * 
 * ORDEN EXACTO DE LAS COLUMNAS REGISTRADAS EN LA HOJA:
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
    var libro = SpreadsheetApp.getActiveSpreadsheet();

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

    // ACCIÓN 1: GUARDAR CUESTIONARIO EN LA NUBE (MULTI-DISPOSITIVO)
    if (datos.action === "saveQuiz") {
      var hojaQuiz = libro.getSheetByName("_Cuestionario");
      if (!hojaQuiz) {
        hojaQuiz = libro.insertSheet("_Cuestionario");
      }
      hojaQuiz.clear();
      hojaQuiz.appendRow(["Propiedad", "Valor"]);
      hojaQuiz.appendRow(["subject", datos.subject || ""]);
      hojaQuiz.appendRow(["teacher", datos.teacher || ""]);
      hojaQuiz.appendRow(["quizTitle", datos.quizTitle || ""]);
      hojaQuiz.appendRow(["updatedAt", Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss")]);
      
      // Guardar el XML (si es muy extenso, guardarlo en filas contiguas)
      var xmlContent = datos.xmlString || "";
      var chunkSize = 45000;
      if (xmlContent.length <= chunkSize) {
        hojaQuiz.appendRow(["xmlPart_0", xmlContent]);
      } else {
        var numParts = Math.ceil(xmlContent.length / chunkSize);
        hojaQuiz.appendRow(["xmlPartsCount", numParts]);
        for (var p = 0; p < numParts; p++) {
          var chunk = xmlContent.substring(p * chunkSize, (p + 1) * chunkSize);
          hojaQuiz.appendRow(["xmlPart_" + p, chunk]);
        }
      }

      var headerR = hojaQuiz.getRange(1, 1, 1, 2);
      headerR.setFontWeight("bold");
      headerR.setBackground("#334155");
      headerR.setFontColor("#ffffff");

      return crearRespuestaJSON({
        status: "success",
        message: "Cuestionario guardado con éxito en tu Google Sheets. Ahora puedes cargarlo desde cualquier dispositivo."
      });
    }

    // ACCIÓN 2: REGISTRAR CALIFICACIÓN DE UN ESTUDIANTE
    var hoja = libro.getSheets()[0];
    prepararHojaSiEstaVacia(hoja);

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
 * Función GET para sincronización de datos y verificación de conexión.
 */
function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || "";
  var libro = SpreadsheetApp.getActiveSpreadsheet();

  // 1. SINCRONIZAR CALIFICACIONES (Para que cualquier dispositivo vea todas las notas)
  if (action === "getRecords") {
    try {
      var hoja = libro.getSheets()[0];
      var ultimaFila = hoja.getLastRow();
      if (ultimaFila <= 1) {
        return crearRespuestaJSON({
          status: "success",
          count: 0,
          records: []
        });
      }

      var numColumnas = Math.min(hoja.getLastColumn(), 9);
      var filas = hoja.getRange(2, 1, ultimaFila - 1, numColumnas).getValues();
      var registros = [];

      for (var i = 0; i < filas.length; i++) {
        var fila = filas[i];
        if (!fila[0] && !fila[1]) continue;

        var pctNum = 0;
        if (typeof fila[8] === "number") {
          pctNum = fila[8] <= 1 ? fila[8] * 100 : fila[8];
        } else if (fila[8]) {
          pctNum = parseFloat(String(fila[8]).replace("%", "").replace(",", ".")) || 0;
        }

        var fechaStr = "";
        if (fila[2] instanceof Date) {
          fechaStr = Utilities.formatDate(fila[2], Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
        } else {
          fechaStr = String(fila[2] || "");
        }

        registros.push({
          studentName: String(fila[0] || "Anónimo"),
          matricula: String(fila[1] || "S/M"),
          submissionDate: fechaStr,
          subject: String(fila[3] || ""),
          teacher: String(fila[4] || ""),
          quizTitle: String(fila[5] || ""),
          scoreEarned: Number(fila[6] || 0),
          scoreMax: Number(fila[7] || 0),
          percentage: Math.round(pctNum * 100) / 100
        });
      }

      return crearRespuestaJSON({
        status: "success",
        count: registros.length,
        records: registros
      });
    } catch (err) {
      return crearRespuestaJSON({
        status: "error",
        message: "Error al leer calificaciones: " + err.toString()
      });
    }
  }

  // 2. DESCARGAR CUESTIONARIO ACTIVO (Para abrir el examen en otro dispositivo)
  if (action === "getQuiz") {
    try {
      var hojaQuiz = libro.getSheetByName("_Cuestionario");
      if (!hojaQuiz || hojaQuiz.getLastRow() < 2) {
        return crearRespuestaJSON({
          status: "error",
          message: "No hay ningún cuestionario guardado en esta hoja. Primero presiona 'Guardar en la nube' en el dispositivo original."
        });
      }

      var valores = hojaQuiz.getDataRange().getValues();
      var propMap = {};
      var xmlParts = [];
      var partsCount = 0;

      for (var r = 1; r < valores.length; r++) {
        var key = String(valores[r][0]);
        var val = valores[r][1];
        if (key === "xmlPartsCount") {
          partsCount = Number(val);
        } else if (key.indexOf("xmlPart_") === 0) {
          var partIndex = parseInt(key.replace("xmlPart_", ""), 10);
          xmlParts[partIndex] = String(val);
        } else {
          propMap[key] = String(val || "");
        }
      }

      var fullXml = xmlParts.join("");

      return crearRespuestaJSON({
        status: "success",
        quiz: {
          subject: propMap.subject || "",
          teacher: propMap.teacher || "",
          quizTitle: propMap.quizTitle || "",
          updatedAt: propMap.updatedAt || "",
          xmlString: fullXml
        }
      });
    } catch (err) {
      return crearRespuestaJSON({
        status: "error",
        message: "Error al recuperar cuestionario: " + err.toString()
      });
    }
  }

  // 3. PING TEST / VERIFICACIÓN DE CONEXIÓN
  if (!action || action === "ping") {
    return crearRespuestaJSON({
      status: "success",
      message: "Conexión exitosa con Google Apps Script. Webhook, sincronización de calificaciones y nube activos.",
      timestamp: new Date().toISOString()
    });
  }

  // 4. Si es un GET con datos de calificación (soporte de respaldo)
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
