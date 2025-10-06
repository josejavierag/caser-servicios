/**
 * Carga citas de correos electrónicos no leídos con la etiqueta 'caser-seguimientos'.
 * Extrae el número de servicio, fecha y hora de la cita del asunto y cuerpo del correo.
 * Actualiza la hoja "SERVICIOS" con esta información y marca el correo como leído.
 * Si se actualiza alguna cita, llama a la función filtroAgenda().
 * Esta función está diseñada para ejecutarse automáticamente mediante un activador temporal,
 * por lo que no muestra mensajes en pantalla al usuario, sino que registra la información.
 */
function cargarCitas() {
  const hojaServicios = HOJA_SERVICIOS;

  if (!hojaServicios) {
    Logger.log(`ERROR: No se encontró la hoja "${NOMBRES_HOJAS.SERVICIOS}". Asegúrate de que existe.`);
    return;
  }

  const ultimaFila = hojaServicios.getLastRow();
  if (ultimaFila < 2) {
    Logger.log(`ADVERTENCIA: No hay datos para procesar en la hoja "${NOMBRES_HOJAS.SERVICIOS}" (menos de 2 filas).`);
    return;
  }

  // Busca hasta 50 hilos no leídos con la etiqueta específica
  const hilosEmail = GmailApp.search('label:caser-citas is:unread').slice(0, 50);
  let seActualizoAlgunaCita = false;

  // Obtener solo la columna de Servicio (Referencia) para buscar de forma eficiente
  const referenciasServiciosEnHoja = hojaServicios.getRange(2, COLUMNAS_SERVICIOS.SERVICIO, ultimaFila - 1, 1).getValues().map(row => String(row[0]).trim());

  hilosEmail.forEach(hilo => {
    try {
      const mensaje = hilo.getMessages().pop(); // Obtener el último mensaje del hilo
      if (!mensaje) {
        Logger.log(`ADVERTENCIA: Hilo sin mensajes. Asunto: "${hilo.getFirstMessageSubject()}"`);
        return;
      }
      const asunto = mensaje.getSubject();
      const cuerpo = mensaje.getPlainBody();

      // Expresión regular para extraer el número de servicio (ej: exp 2024/12345)
      const matchExpediente = asunto.match(/exp\s*.*?(\d{4})\/(\d{4,7})/i);
      if (!matchExpediente) {
        Logger.log(`ADVERTENCIA: No se pudo extraer el número de servicio del asunto: "${asunto}"`);
        return;
      }

      const numeroServicio = `${matchExpediente[1]}/${matchExpediente[2]}`;

      // Expresión regular para extraer fecha y hora (ej: fecha 01-01-2024 de 10:30)
      const matchFechaHora = cuerpo.match(/fecha\s*(\d{2})-(\d{2})-(\d{4})\s*de\s*(\d{2}:\d{2})/i);
      if (!matchFechaHora) {
        Logger.log(`ADVERTENCIA: No se pudo extraer la fecha y hora de la cita del cuerpo del correo para el servicio ${numeroServicio}. Cuerpo: "${cuerpo.substring(0, Math.min(cuerpo.length, 100))}..."`);
        return;
      }

      const dia = parseInt(matchFechaHora[1], 10);
      const mes = parseInt(matchFechaHora[2], 10) - 1; // Meses son 0-indexados en JavaScript
      const año = parseInt(matchFechaHora[3], 10);
      const hora = matchFechaHora[4];

      const fechaCitaObj = new Date(año, mes, dia);
      if (isNaN(fechaCitaObj.getTime())) {
        Logger.log(`ERROR: Fecha inválida extraída para el servicio ${numeroServicio}: ${matchFechaHora[3]}-${matchFechaHora[2]}-${matchFechaHora[1]}`);
        return;
      }

      // Buscar el servicio en la hoja usando la referencia
      const indiceFila = referenciasServiciosEnHoja.indexOf(numeroServicio);

      if (indiceFila !== -1) {
        const filaRealEnHoja = indiceFila + 2; // +1 por 0-indexación, +1 por fila de cabecera
        hojaServicios.getRange(filaRealEnHoja, COLUMNAS_SERVICIOS.FECHA).setValue(fechaCitaObj);
        hojaServicios.getRange(filaRealEnHoja, COLUMNAS_SERVICIOS.HORA).setValue(hora);
        hojaServicios.getRange(filaRealEnHoja, COLUMNAS_SERVICIOS.ESTADO).setValue("CITADO");

        Logger.log(`Servicio ${numeroServicio} actualizado en la fila ${filaRealEnHoja} con fecha ${Utilities.formatDate(fechaCitaObj, HOJA_ACTIVA.getSpreadsheetTimeZone(), "dd/MM/yyyy")} y hora ${hora}.`);
        seActualizoAlgunaCita = true;
        hilo.markRead(); // Marcar el hilo como leído
        Logger.log(`Hilo de correo para ${numeroServicio} marcado como leído.`);
      } else {
        Logger.log(`ADVERTENCIA: No se encontró el servicio ${numeroServicio} en la hoja "${NOMBRES_HOJAS.SERVICIOS}".`);
      }

    } catch (error) {
      Logger.log(`ERROR procesando hilo de correo (Asunto: "${hilo.getFirstMessageSubject()}"): ${error.message}`);
    }
  });

  // Nota: La llamada a filtroAgenda() se mantiene si se necesita al final del proceso.
}

