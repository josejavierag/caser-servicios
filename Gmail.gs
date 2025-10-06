function etiquetarYArchivarCorreos() {
  Logger.log("--- INICIO DEL PROCESO: Etiquetar y Archivar Correos ---");

  // 1. Búsqueda de hilos de correos
  var query = 'is:inbox label:unread newer_than:1d from:(caser.es OR integraval.es)';
  Logger.log("1. Ejecutando búsqueda en Gmail con la consulta: '" + query + "'");
  var threads = GmailApp.search(query);

  if (!threads || threads.length === 0) {
    Logger.log("2. No se encontraron hilos nuevos. Proceso finalizado.");
    return;
  }
  Logger.log("2. Se encontraron " + threads.length + " hilos para procesar.");

  // 3. Carga de etiquetas (con formato anidado)
  Logger.log("3. Obteniendo las etiquetas de Gmail...");
  var etiquetas = {
    "caser-citas": GmailApp.getUserLabelByName("caser/citas"),
    "caser-seguimientos": GmailApp.getUserLabelByName("caser/seguimientos"),
    "caser-servicios": GmailApp.getUserLabelByName("caser/servicios"),
    "caser-incidencias": GmailApp.getUserLabelByName("caser/incidencias"),
    "caser-devoluciones": GmailApp.getUserLabelByName("caser/devoluciones"),
    "integraval": GmailApp.getUserLabelByName("integraval"),
    "integraval-servicios": GmailApp.getUserLabelByName("integraval/servicios")
  };

  // Log de diagnóstico para verificar que las etiquetas existen
  for (var key in etiquetas) {
    Logger.log("   -> Verificando etiqueta '" + key + "': " + (etiquetas[key] ? "ENCONTRADA ✔️" : "NO ENCONTRADA ❌"));
  }

  function aplicarEtiqueta(thread, label, mensajeLog) {
    if (label) {
      thread.addLabel(label);
      Logger.log("      ✔️ ETIQUETADO: " + mensajeLog);
    } else {
      Logger.log("      ⚠️ ETIQUETA NO ENCONTRADA para: " + mensajeLog);
    }
  }

  // 4. Procesamiento de cada hilo
  Logger.log("4. Iniciando el bucle para procesar cada hilo...");
  threads.forEach(function(thread, index) {
    var asunto = thread.getFirstMessageSubject();
    Logger.log("\n--- Procesando Hilo " + (index + 1) + ": '" + asunto + "' ---");

    var mensaje = thread.getMessages()[0]; // Analizamos solo el primer mensaje
    var from = mensaje.getFrom().toLowerCase();
    var body = mensaje.getPlainBody().toLowerCase();
    var attachments = mensaje.getAttachments();
    var etiquetado = false;

    // --- LÓGICA PARA CORREOS DE CASER ---
    if (from.includes("caser.es")) {
      Logger.log("   -> Remitente es de caser.es. Aplicando reglas en orden de prioridad...");
      var attName = attachments.length > 0 ? attachments[0].getName().toLowerCase() : "";

      // Regla 1: SERVICIOS (por adjunto)
      if (attachments.length > 0 && attName.startsWith("produccion_aviso de prestaciones")) {
        aplicarEtiqueta(thread, etiquetas["caser-servicios"], "CASER/SERVICIOS");
        etiquetado = true;
      }
      // Regla 2: CITAS (por texto en el cuerpo)
      else if (body.includes("solicitud de cita")) {
        aplicarEtiqueta(thread, etiquetas["caser-citas"], "CASER/CITAS");
        etiquetado = true;
      }
      // Regla 3: INCIDENCIAS (por adjunto)
      else if (attachments.length > 0 && attName.startsWith("produccion_aviso incidencia")) {
        aplicarEtiqueta(thread, etiquetas["caser-incidencias"], "CASER/INCIDENCIAS");
        etiquetado = true;
      }
      // Regla 4: DEVOLUCIONES (por texto en el cuerpo)
      else if (attachments.length > 0 && attName.startsWith("produccion_devolucion")) {
        aplicarEtiqueta(thread, etiquetas["caser-devoluciones"], "CASER/DEVOLUCIONES");
        etiquetado = true;
      }
      // Regla 5: SEGUIMIENTOS (si no se cumple ninguna anterior)
      else {
        aplicarEtiqueta(thread, etiquetas["caser-seguimientos"], "CASER/SEGUIMIENTOS (por defecto)");
        etiquetado = true;
      }
    }

    // --- LÓGICA PARA CORREOS DE INTEGRAVAL ---
    if (from.includes("integraval.es")) {
      Logger.log("   -> Remitente es de integraval.es. Analizando...");
      var etiquetadoIntegravalServicios = false;
      if (attachments.length > 0) {
        attachments.forEach(function(att) {
          if (att.getName().toLowerCase().includes("orden de trabajo")) {
            etiquetadoIntegravalServicios = true;
          }
        });
      }
      
      // Regla 1: INTEGRAVAL/SERVICIOS (por adjunto)
      if (etiquetadoIntegravalServicios) {
        aplicarEtiqueta(thread, etiquetas["integraval-servicios"], "INTEGRAVAL/SERVICIOS");
        etiquetado = true;
      }
      // Regla 2: INTEGRAVAL (si no se cumple la anterior)
      else {
        aplicarEtiqueta(thread, etiquetas["integraval"], "INTEGRAVAL (general)");
        etiquetado = true;
      }
    }

    // 5. Archivar el hilo SÓLO SI FUE ETIQUETADO
    if (etiquetado) {
      thread.moveToArchive();
      Logger.log("   -> ✅ Hilo archivado con éxito (permanece no leído).");
    } else {
      Logger.log("   -> ℹ️ No se cumplió ninguna regla de etiquetado. El hilo se mantiene en Recibidos.");
    }
  });

  Logger.log("\n--- PROCESO FINALIZADO ---");
}