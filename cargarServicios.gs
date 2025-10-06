function cargarServicios() {
  Logger.log("--- INICIO DEL PROCESO: cargarServicios ---");
  const hojaServicios = HOJA_SERVICIOS;
  if (!hojaServicios) {
    Logger.log(`Error: La hoja "${NOMBRES_HOJAS.SERVICIOS}" no se encontró.`);
    return;
  }
  const GMAIL_QUERY = 'label:caser-servicios is:unread has:attachment';
  const hilos = GmailApp.search(GMAIL_QUERY, 0, 1);
  if (hilos.length === 0) {
    Logger.log("No se encontraron nuevos correos para procesar.");
    return;
  }
  const hilo = hilos[0];
  const mensaje = hilo.getMessages()[0];
  const adjuntos = mensaje.getAttachments();
  const asunto = mensaje.getSubject();
  let docIdTemporal = null;

  if (!adjuntos || adjuntos.length === 0) {
    hilo.markRead();
    Logger.log(`Advertencia: Hilo sin adjuntos (Asunto: "${asunto}"). Marcado como leído.`);
    return;
  }

  const matchServicio = asunto.match(/\b\d{4}\/\d{4,7}\b/);
  if (!matchServicio) {
    hilo.markRead();
    Logger.log(`No se pudo extraer el nº de servicio del asunto: "${asunto}". Marcado como leído.`);
    return;
  }
  const numeroServicio = matchServicio[0];
  Logger.log(`--- Iniciando procesamiento para servicio: ${numeroServicio} ---`);

  try {
    const adjunto = adjuntos[0];
    
    // 1. GESTIÓN DEL ARCHIVO EN DRIVE Y EXTRACCIÓN DE URL (CORREGIDO)
    Logger.log("1. Guardando archivo PDF y extrayendo URL...");
    const carpetaRaizServiciosId = '1GMkb-r0rpjJKerOuIr3JE_95GZ2DSjo1';
    const carpetaRaiz = DriveApp.getFolderById(carpetaRaizServiciosId);
    
    const archivoGuardado = carpetaRaiz.createFile(adjunto.copyBlob()).setName(`${numeroServicio}_aviso.pdf`);
    const urlArchivo = archivoGuardado.getUrl();
    const idCarpetaFotos = '';

    // 2. EXTRACCIÓN DE TEXTO DEL DOCUMENTO
    Logger.log("2. Extrayendo texto del documento...");
    const tempFile = DriveApp.createFile(adjunto.copyBlob());
    const convertedFile = Drive.Files.copy({ mimeType: 'application/vnd.google-apps.document' }, tempFile.getId());
    docIdTemporal = convertedFile.id;
    tempFile.setTrashed(true);
    const doc = DocumentApp.openById(docIdTemporal);
    const textoCompletoDoc = doc.getBody().getText();

    // 3. RECOPILACIÓN COMPLETA DE DATOS
    Logger.log("3. Recopilando todos los datos del texto...");
    const plataforma = obtenerCodigoDestinatario(extraerValor(textoCompletoDoc, /DESTINATARIO: (.+)/));
    const tipoTratamiento = extraerValor(textoCompletoDoc, /TRATAMIENTO: (.+)/i);
    const poliza = tipoTratamiento ? (tipoTratamiento.toLowerCase().includes("oro") ? "ORO" : (tipoTratamiento.toLowerCase().includes("plata") ? "PLATA" : (tipoTratamiento.toLowerCase().includes("platino") ? "PLATINO" : "NORMAL"))) : "NORMAL";
    const observaciones = extraerValor(textoCompletoDoc, /URGENTE: (.+)/) === "SÍ" ? "URGENCIA" : "";
    const { fechaCita, horaCita } = extraerFechaHoraCita(textoCompletoDoc);
    const direccion = extraerDireccion(textoCompletoDoc);
    const localidad = extraerLocalidad(textoCompletoDoc);
    const descripcion = extraerValor(textoCompletoDoc, /DESCRIPCIÓN: (.*)/);
    const dni = extraerValor(textoCompletoDoc, /NIF\/CIF: (.+)/);
    const nombre = extraerValor(textoCompletoDoc, /NOMBRE: (.*)/);
    const telefonos = [...textoCompletoDoc.matchAll(/\b\d{9}\b/g)].map(match => match[0]);
    const asegurado = telefonos[0] || '';
    const inquilino = telefonos[1] || '';

    // 4. CONSTRUCCIÓN DEL ARRAY DE VALORES PARA LA HOJA
    Logger.log("4. Construyendo la fila de datos...");
    const valoresFila = Array(hojaServicios.getLastColumn()).fill('');
    valoresFila[COLUMNAS_SERVICIOS.PLATAFORMA - 1] = plataforma;
    valoresFila[COLUMNAS_SERVICIOS.SERVICIO - 1] = numeroServicio;
    valoresFila[COLUMNAS_SERVICIOS.POLIZA - 1] = poliza;
    valoresFila[COLUMNAS_SERVICIOS.OBSERVACIONES - 1] = observaciones;
    valoresFila[COLUMNAS_SERVICIOS.DIRECCION - 1] = direccion;
    valoresFila[COLUMNAS_SERVICIOS.LOCALIDAD - 1] = localidad;
    valoresFila[COLUMNAS_SERVICIOS.DESCRIPCION - 1] = descripcion;
    valoresFila[COLUMNAS_SERVICIOS.ASEGURADO - 1] = asegurado;
    valoresFila[COLUMNAS_SERVICIOS.INQUILINO - 1] = inquilino;
    valoresFila[COLUMNAS_SERVICIOS.DNI - 1] = dni;
    valoresFila[COLUMNAS_SERVICIOS.NOMBRE - 1] = nombre;
    valoresFila[COLUMNAS_SERVICIOS.PDF - 1] = urlArchivo; 
    valoresFila[COLUMNAS_SERVICIOS.FOTOS - 1] = idCarpetaFotos;

    // 5. LÓGICA DE INSERCIÓN O ACTUALIZACIÓN
    Logger.log("5. Decidiendo si insertar o actualizar...");
    let filaExistente = -1;
    const ultimaFila = hojaServicios.getLastRow();
    if (ultimaFila > 1) {
      const datosColumnaServicio = hojaServicios.getRange(2, COLUMNAS_SERVICIOS.SERVICIO, ultimaFila - 1, 1).getValues();
      const indice = datosColumnaServicio.flat().indexOf(numeroServicio);
      if (indice !== -1) filaExistente = indice + 2;
    }

    if (filaExistente !== -1) {
      // CASO: ACTUALIZAR (REAPERTURA)
      Logger.log(` -> Servicio ${numeroServicio} encontrado en fila ${filaExistente}. Actualizando como REAPERTURA.`);
      const fechaAltaOriginal = hojaServicios.getRange(filaExistente, COLUMNAS_SERVICIOS.ALTA).getDisplayValue();
      valoresFila[COLUMNAS_SERVICIOS.ALTA - 1] = fechaAltaOriginal;
      const hoy = new Date();
      let diasIncremento = 1;
      const diaSemanaHoy = hoy.getDay();
      if (diaSemanaHoy === 5) diasIncremento = 3; 
      else if (diaSemanaHoy === 6) diasIncremento = 2;
      const fechaReapertura = new Date(new Date().setDate(hoy.getDate() + diasIncremento));
      valoresFila[COLUMNAS_SERVICIOS.ESTADO - 1] = "CITAR";
      valoresFila[COLUMNAS_SERVICIOS.OBSERVACIONES - 1] = "REAPERTURA";
      valoresFila[COLUMNAS_SERVICIOS.FECHA - 1] = fechaReapertura;
      valoresFila[COLUMNAS_SERVICIOS.HORA - 1] = "";
      valoresFila[COLUMNAS_SERVICIOS.OPERARIO - 1] = "";
      hojaServicios.getRange(filaExistente, 1, 1, valoresFila.length).setValues([valoresFila]);
      Logger.log(` -> Fila ${filaExistente} actualizada.`);
    } else {
      // CASO: INSERTAR (NUEVO SERVICIO)
      Logger.log(` -> Servicio ${numeroServicio} es nuevo. Insertando nueva fila.`);
      let estadoCita = "CITAR";
      let fechaFinalCitaObj;
      if (fechaCita) {
        const partesFecha = fechaCita.split('-');
        fechaFinalCitaObj = new Date(parseInt(partesFecha[2], 10), parseInt(partesFecha[1], 10) - 1, parseInt(partesFecha[0], 10));
        if (horaCita) estadoCita = "CITADO";
      } else {
          const hoy = new Date();
          if (observaciones.includes("URGENCIA")) {
            fechaFinalCitaObj = hoy;
          } else {
            let diasIncremento = 1;
            const diaSemanaHoy = hoy.getDay();
            if (diaSemanaHoy === 5) diasIncremento = 3; else if (diaSemanaHoy === 6) diasIncremento = 2;
            fechaFinalCitaObj = new Date(new Date().setDate(hoy.getDate() + diasIncremento));
          }
      }
      valoresFila[COLUMNAS_SERVICIOS.ALTA - 1] = new Date();
      valoresFila[COLUMNAS_SERVICIOS.FECHA - 1] = fechaFinalCitaObj;
      valoresFila[COLUMNAS_SERVICIOS.HORA - 1] = horaCita;
      valoresFila[COLUMNAS_SERVICIOS.ESTADO - 1] = estadoCita;
      const filaDestino = ultimaFila + 1;
      hojaServicios.getRange(filaDestino, 1, 1, valoresFila.length).setValues([valoresFila]);
      Logger.log(` -> Servicio NUEVO añadido en la fila ${filaDestino}.`);
      if (ultimaFila >= 2) {
        hojaServicios.getRange(2, 1, 1, hojaServicios.getLastColumn()).copyTo(
          hojaServicios.getRange(filaDestino, 1, 1, hojaServicios.getLastColumn()),
          SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
      }
    }
    hilo.markRead();
    Logger.log(`-> Hilo de correo para ${numeroServicio} marcado como leído.`);
  } catch (error) {
    Logger.log(`❌ ERROR al cargar servicio (Asunto: "${asunto}"): ${error.message}\nStack: ${error.stack}`);
  } finally {
    if (docIdTemporal) {
      try { DriveApp.getFileById(docIdTemporal).setTrashed(true); } catch (e) { /* Ignorar */ }
    }
    Logger.log(`--- Fin procesamiento para servicio: ${numeroServicio} ---`);
  }
}

function extraerValor(texto, patron) {
  const match = texto.match(patron);
  return match && match[1] ? match[1].trim() : '';
}

function obtenerCodigoDestinatario(textoDestinatario) {
  const textoLower = textoDestinatario ? textoDestinatario.toLowerCase() : "";
  if (textoLower.includes("(bri)")) return "BRI980Y";
  if (textoLower.includes("(brico)")) return "BRI980B";
  if (textoLower.includes("(cx)")) return "COX980F";
  if (textoLower.includes("(albacete)")) return "UPR980A";
  return "UPR980M";
}

function extraerFechaHoraCita(textoCompletoDoc) {
  let fechaCita = '';
  let horaCita = '';
  const seccionNotasIndex = textoCompletoDoc.indexOf("NOTAS AL PROVEEDOR");
  const textoBusqueda = seccionNotasIndex !== -1 ?
    textoCompletoDoc.slice(seccionNotasIndex + "NOTAS AL PROVEEDOR".length) :
    textoCompletoDoc;
  const citaRegex = /(\d{2}-\d{2}-\d{4})\s*(?:de\s*)?(\d{2}:\d{2})/gi;
  const coincidencias = [...textoBusqueda.matchAll(citaRegex)];
  if (coincidencias.length > 0) {
    const ultimaCoincidencia = coincidencias[coincidencias.length - 1];
    fechaCita = ultimaCoincidencia[1];
    horaCita = ultimaCoincidencia[2];
    let [horas, minutos] = horaCita.split(':').map(Number);
    const fechaHoraOriginal = new Date();
    fechaHoraOriginal.setHours(horas, minutos, 0, 0);
    const fechaHoraAjustada = new Date(fechaHoraOriginal.getTime() - 90 * 60 * 1000);
    horaCita = `${String(fechaHoraAjustada.getHours()).padStart(2, '0')}:${String(fechaHoraAjustada.getMinutes()).padStart(2, '0')}`;
  }
  return { fechaCita, horaCita };
}

function extraerDireccion(textoDocumento) {
  const domicilioMatch = textoDocumento.match(/DOMICILIO:\s*(.+)/i);
  if (!domicilioMatch || !domicilioMatch[1]) return "";
  let lineaDomicilioCompleta = domicilioMatch[1].trim();
  const firstCommaIndex = lineaDomicilioCompleta.indexOf(',');
  if (firstCommaIndex !== -1) return lineaDomicilioCompleta.substring(0, firstCommaIndex).trim().toUpperCase();
  const fechaOcurrenciaIndex = lineaDomicilioCompleta.indexOf("FECHA OCURRENCIA");
  if (fechaOcurrenciaIndex !== -1) return lineaDomicilioCompleta.substring(0, fechaOcurrenciaIndex).trim().toUpperCase();
  return lineaDomicilioCompleta.toUpperCase();
}

function extraerLocalidad(textoDocumento) {
  const domicilioMatch = textoDocumento.match(/DOMICILIO:\s*(.+)/i);
  if (!domicilioMatch || !domicilioMatch[1]) return "";
  let lineaDomicilioCompleta = domicilioMatch[1].trim();
  const firstCommaIndex = lineaDomicilioCompleta.indexOf(',');
  let parteLocalidadProvincia = firstCommaIndex !== -1 ? lineaDomicilioCompleta.substring(firstCommaIndex + 1).trim() : "";
  if (!parteLocalidadProvincia) return "";
  parteLocalidadProvincia = parteLocalidadProvincia.replace(/,?\s*\b\d{5}\b/g, '').trim();
  const fechaOcurrenciaIndex = parteLocalidadProvincia.indexOf("FECHA OCURRENCIA");
  if (fechaOcurrenciaIndex !== -1) parteLocalidadProvincia = parteLocalidadProvincia.substring(0, fechaOcurrenciaIndex).trim();
  const provincias = ["albacete", "ciudad real", "cuenca", "guadalajara", "toledo"];
  let localidadFinal = parteLocalidadProvincia;
  let provinciaDetectada = "";
  for (const prov of provincias) {
    if (parteLocalidadProvincia.toLowerCase().endsWith(prov)) {
      provinciaDetectada = prov;
      localidadFinal = parteLocalidadProvincia.substring(0, parteLocalidadProvincia.length - prov.length).replace(/,$/, '').trim();
      break;
    }
  }
  if (!localidadFinal || localidadFinal.toLowerCase() === provinciaDetectada.toLowerCase()) return provinciaDetectada.toUpperCase();
  return localidadFinal.toUpperCase();
}

/**
 * Carga nuevos servicios desde correos no leídos con la etiqueta 'integraval-servicios'.
 * Extrae información del adjunto PDF y la guarda en la hoja "INTEGRAVAL", evitando duplicados.
 */
function cargarIntegraval() {
  const NOMBRES_HOJAS = {
    INTEGRAVAL: "INTEGRAVAL",
  };
  const HOJA_INTEGRAVAL = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOMBRES_HOJAS.INTEGRAVAL);

  if (!HOJA_INTEGRAVAL) {
    Logger.log(`Error: La hoja "${NOMBRES_HOJAS.INTEGRAVAL}" no se encontró.`);
    return;
  }
  
  const GMAIL_QUERY = 'label:integraval-servicios is:unread has:attachment';
  const hilos = GmailApp.search(GMAIL_QUERY, 0, 1);

  if (hilos.length === 0) {
    Logger.log("No se encontraron nuevos correos de servicio de Integraval para procesar.");
    return;
  }

  const hilo = hilos[0];
  const mensaje = hilo.getMessages()[0];
  const asunto = mensaje.getSubject();
  const adjuntos = mensaje.getAttachments();
  let docIdTemporal = null;

  // Se mantiene la lógica de extracción del número de servicio del asunto
  let numeroServicio = null;
  const matchC = asunto.match(/C-\d{2}_202\d-\d{2}[\/\_]?\d{6}/);
  const matchT = asunto.match(/T-\d{4}-\d{6}/);
  const matchSimple = asunto.match(/\b\d{3,}\b/);

  if (matchC) {
    numeroServicio = matchC[0];
  } else if (matchT) {
    numeroServicio = matchT[0];
  } else if (matchSimple) {
    numeroServicio = matchSimple[0];
  }

  if (!numeroServicio) {
    Logger.log(`No se pudo extraer el número de expediente del asunto: "${asunto}". Marcando como leído.`);
    hilo.markRead();
    return;
  }
  
  Logger.log(`--- Iniciando procesamiento para expediente: ${numeroServicio} (Asunto: ${asunto}) ---`);

  if (!adjuntos || adjuntos.length === 0) {
    Logger.log(`Advertencia: Hilo de correo sin adjuntos. Asunto: "${asunto}". Marcando como leído.`);
    hilo.markRead();
    return;
  }
  
  // MODIFICACIÓN: Lógica simplificada de búsqueda de PDF
  const PARTE_DE_CONFORMIDAD_NOMBRE = "PARTE DE CONFORMIDAD INTEGRAVAL.pdf";
  
  // Filtra los PDF válidos, descartando el "Parte de Conformidad"
  const pdfsValidos = adjuntos.filter(a => 
    a.getContentType() === MimeType.PDF && 
    a.getName() !== PARTE_DE_CONFORMIDAD_NOMBRE
  );
  
  if (pdfsValidos.length !== 1) {
    Logger.log(`No se encontró exactamente un PDF válido para procesar (se encontraron ${pdfsValidos.length}). Marcando como leído.`);
    hilo.markRead();
    return;
  }
  
  const adjunto = pdfsValidos[0];
  Logger.log(`Procesando adjunto: ${adjunto.getName()}`);

  try {
    const archivoTemporalBlob = adjunto.copyBlob();
    const tempFile = DriveApp.createFile(archivoTemporalBlob);
    const convertedFile = Drive.Files.copy({ mimeType: 'application/vnd.google-apps.document' }, tempFile.getId());
    docIdTemporal = convertedFile.id;
    tempFile.setTrashed(true);

    const doc = DocumentApp.openById(docIdTemporal);
    const textoCompletoDoc = doc.getBody().getText();
    Logger.log(`Texto extraído del documento. Longitud: ${textoCompletoDoc.length} caracteres.`);

    const ultimaFila = HOJA_INTEGRAVAL.getLastRow();
    if (ultimaFila > 1) {
      const referenciasExistentes = HOJA_INTEGRAVAL.getRange(2, 2, ultimaFila - 1, 1).getValues().flat();
      if (referenciasExistentes.includes(numeroServicio)) {
        Logger.log(`Servicio duplicado descartado: ${numeroServicio} ya existe en la hoja "${NOMBRES_HOJAS.INTEGRAVAL}". Marcando correo como leído.`);
        hilo.markRead();
        return;
      }
    }

    // --- Funciones auxiliares para la extracción de datos de Integraval ---
    const extraerValor = (patron, texto = textoCompletoDoc) => {
      const match = texto.match(patron);
      const valor = match && match[1] ? match[1].trim() : '';
      Logger.log(`   ExtraerValor (Patrón: ${patron}): "${valor}"`);
      return valor;
    };

    const extraerDireccion = () => {
      const match = textoCompletoDoc.match(/DIRECCIÓN: (.+)/);
      return match ? match[1].trim() : '';
    };

    const extraerLocalidad = () => {
      const match = textoCompletoDoc.match(/Poblaci(ó|o)n \/ Provincia \/ C\.P\.:\s*([^\n\/]+)/i);
      return match ? match[2].trim().toUpperCase() : '';
    };
    
    const extraerInquilino = () => {
      const match = textoCompletoDoc.match(/INQUILINO:\s*.+?(\d{9,})/);
      return match ? match[1].trim() : '';
    };

    const extraerDescripcion = () => {
      const match = textoCompletoDoc.match(/ACTUACION:([\s\S]*?)(?:SOLUCION PROPUESTA ADMINISTRACION:|TRABAJO A REALIZAR:|INQUILINO:)/);
      return match ? match[1].trim() : '';
    };

    // --- Recopilación de todos los datos para la nueva fila ---
    const fechaAsignacionRaw = extraerValor(/FECHA ASIGNACIÓN:\s*(.+)/i);
    const fechaAsignacion = fechaAsignacionRaw ? Utilities.formatDate(new Date(fechaAsignacionRaw), Session.getScriptTimeZone(), "dd/MM/yyyy") : '';
    
    const direccion = extraerDireccion();
    const localidad = extraerLocalidad();
    const inquilino = extraerInquilino();
    const descripcion = extraerDescripcion();

    const valoresNuevaFila = [];
    valoresNuevaFila[0] = fechaAsignacion; // Columna A
    valoresNuevaFila[1] = numeroServicio;   // Columna B
    valoresNuevaFila[2] = "";               // Columna C
    valoresNuevaFila[3] = "";               // Columna D
    valoresNuevaFila[4] = "";               // Columna E
    valoresNuevaFila[5] = "CITAR";          // Columna F
    valoresNuevaFila[6] = direccion;        // Columna G
    valoresNuevaFila[7] = localidad;        // Columna H
    valoresNuevaFila[8] = descripcion;      // Columna I
    valoresNuevaFila[9] = "";               // Columna J
    valoresNuevaFila[10] = inquilino;       // Columna K

    const filaDestino = HOJA_INTEGRAVAL.getLastRow() + 1;
    Logger.log(`Intentando añadir fila en la hoja "${NOMBRES_HOJAS.INTEGRAVAL}" en la fila ${filaDestino}.`);
    
    HOJA_INTEGRAVAL.getRange(filaDestino, 1, 1, valoresNuevaFila.length).setValues([valoresNuevaFila]);
    Logger.log(`Servicio de Integraval ${numeroServicio} procesado y añadido en la fila ${filaDestino}.`);

    if (HOJA_INTEGRAVAL.getLastRow() >= 2) {
      const rangoFormatoOrigen = HOJA_INTEGRAVAL.getRange(2, 1, 1, HOJA_INTEGRAVAL.getLastColumn());
      const rangoNuevaFilaParaFormato = HOJA_INTEGRAVAL.getRange(filaDestino, 1, 1, HOJA_INTEGRAVAL.getLastColumn());
      rangoFormatoOrigen.copyTo(rangoNuevaFilaParaFormato, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
      Logger.log(`Formato de la fila 2 copiado a la nueva fila ${filaDestino}.`);
    }

    hilo.markRead();
    Logger.log(`Hilo de correo para ${numeroServicio} marcado como leído.`);
    Logger.log(`--- Fin procesamiento para expediente: ${numeroServicio} ---`);

  } catch (error) {
    Logger.log(`ERROR al cargar servicio de Integraval desde correo (Asunto: "${asunto}"): ${error.message}\nStack: ${error.stack}`);
  } finally {
    if (docIdTemporal) {
      try {
        DriveApp.getFileById(docIdTemporal).setTrashed(true);
        Logger.log(`Archivo temporal de Drive ${docIdTemporal} eliminado del cubo de basura.`);
      } catch (e) {
        Logger.log(`Advertencia: No se pudo eliminar el archivo temporal ${docIdTemporal} (quizás ya fue eliminado): ${e.message}`);
      }
    }
  }
}