/**
 * Carga los seguimientos de los servicios desde correos no leídos con la etiqueta 'caser-seguimientos'
 * y del remitente 'reparadorpap@caser.es'. Extrae la información y la guarda en la hoja "SEGUIMIENTOS".
 */
function cargarSeguimientos() {
  const hojaSeguimientos = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("SEGUIMIENTOS");

  if (!hojaSeguimientos) {
    Logger.log(`Error: La hoja "SEGUIMIENTOS" no se encontró.`);
    return;
  }
  
  const COLUMNAS_SEGUIMIENTOS = {
    ID: 1,
    CUERPO: 2,
    FECHA: 3,
    SERVICIO: 4
  };

  const GMAIL_QUERY = 'label:caser-seguimientos is:unread from:reparadorpap@caser.es';
  const hilos = GmailApp.search(GMAIL_QUERY, 0, 100);

  if (hilos.length === 0) {
    Logger.log("No se encontraron nuevos correos de seguimiento para procesar.");
    return;
  }

  Logger.log(`Se encontraron ${hilos.length} hilos de seguimiento para procesar.`);

  hilos.forEach(hilo => {
    const mensaje = hilo.getMessages()[0];
    const asunto = mensaje.getSubject();
    const adjuntos = mensaje.getAttachments();
    let cuerpoMensaje = mensaje.getPlainBody();
    let docIdTemporal = null;
    let adjuntoEncontrado = false;

    Logger.log(`--- Procesando correo de seguimiento: "${asunto}" ---`);

    try {
      const matchServicio = asunto.match(/(\d{4}\/\d{4,7}\b)/);
      if (!matchServicio) {
        Logger.log(`Advertencia: No se pudo extraer el número de servicio del asunto: "${asunto}". Se marcará como leído.`);
        hilo.markRead();
        return;
      }
      const numeroServicio = matchServicio[1];
      Logger.log(`Número de servicio extraído: ${numeroServicio}`);

      if (adjuntos && adjuntos.length > 0) {
        for (const adjunto of adjuntos) {
          if (adjunto.getContentType() === 'application/pdf' || adjunto.getName().toLowerCase().endsWith(".pdf")) {
            Logger.log(`Procesando adjunto PDF: ${adjunto.getName()}`);

            const archivoTemporalBlob = adjunto.copyBlob();
            const tempFile = DriveApp.createFile(archivoTemporalBlob);
            const convertedFile = Drive.Files.copy({ mimeType: 'application/vnd.google-apps.document' }, tempFile.getId());
            docIdTemporal = convertedFile.id;
            tempFile.setTrashed(true);

            const doc = DocumentApp.openById(docIdTemporal);
            const textoCompletoDoc = doc.getBody().getText();
            
            const markerInicio = "(ESPAÑA)";
            const markerFin = "Reciban un cordial saludo.";
            
            const posInicio = textoCompletoDoc.indexOf(markerInicio);
            const posFin = textoCompletoDoc.indexOf(markerFin);
            
            if (posInicio !== -1 && posFin !== -1 && posFin > posInicio) {
              const inicioCuerpo = posInicio + markerInicio.length;
              cuerpoMensaje = textoCompletoDoc.substring(inicioCuerpo, posFin).trim();
            } else {
              cuerpoMensaje = "No se pudo extraer el texto filtrado del PDF.";
            }

            Logger.log(`Texto extraído del PDF:\n"${cuerpoMensaje}"`);
            adjuntoEncontrado = true;
            break; 
          }
        }
      }

      if (!adjuntoEncontrado) {
        Logger.log("No se encontró ningún adjunto PDF. Se usará el cuerpo del correo.");
      }

      const ahora = new Date();
      const zonaHorariaScript = Session.getScriptTimeZone();
      const fechaFormateada = Utilities.formatDate(ahora, zonaHorariaScript, "dd/MM/yy");
      const idAleatorio = Utilities.getUuid();

      // Creamos la nueva fila con los datos exactos que queremos guardar
      const nuevaFila = [
        idAleatorio,
        cuerpoMensaje,
        fechaFormateada,
        numeroServicio
      ];
      
      const filaDestino = hojaSeguimientos.getLastRow() + 1;
      hojaSeguimientos.getRange(filaDestino, 1, 1, nuevaFila.length).setValues([nuevaFila]);
      Logger.log(`Seguimiento para el servicio ${numeroServicio} añadido en la fila ${filaDestino}.`);

      hilo.markRead();
      Logger.log(`Correo de seguimiento para ${numeroServicio} marcado como leído.`);

    } catch (error) {
      Logger.log(`ERROR al procesar seguimiento (Asunto: "${asunto}"): ${error.message}\nStack: ${error.stack}`);
    } finally {
      if (docIdTemporal) {
        try {
          DriveApp.getFileById(docIdTemporal).setTrashed(true);
          Logger.log(`Archivo temporal ${docIdTemporal} eliminado.`);
        } catch (e) {
          Logger.log(`Advertencia: No se pudo eliminar el archivo temporal ${docIdTemporal}: ${e.message}`);
        }
      }
    }
    Logger.log(`--- Fin de procesamiento del correo ---`);
  });
}