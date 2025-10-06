/**
 * Carga los datos de un servicio seleccionado en la hoja "SERVICIOS"
 * y los transfiere a la hoja "PRESUPUESTO" para iniciar un nuevo presupuesto.
 * Esta función está diseñada para ser llamada desde el menú personalizado.
 */
function presupuestarServicio() {
  const ui = SpreadsheetApp.getUi();
  const hojaServicios = HOJA_SERVICIOS;
  const hojaPresupuesto = HOJA_PRESUPUESTO;

  if (!hojaServicios) {
    ui.alert("Error", `La hoja "${NOMBRES_HOJAS.SERVICIOS}" no se encontró. Asegúrate de que existe.`, ui.ButtonSet.OK);
    Logger.log(`Error: La hoja "${NOMBRES_HOJAS.SERVICIOS}" no se encontró al intentar presupuestar.`);
    return;
  }
  if (!hojaPresupuesto) {
    ui.alert("Error", `La hoja "${NOMBRES_HOJAS.PRESUPUESTO}" no se encontró. Asegúrate de que existe.`, ui.ButtonSet.OK);
    Logger.log(`Error: La hoja "${NOMBRES_HOJAS.PRESUPUESTO}" no se encontró al intentar presupuestar.`);
    return;
  }

  if (HOJA_ACTIVA.getActiveSheet().getName() !== NOMBRES_HOJAS.SERVICIOS) {
    ui.alert("Advertencia", `¡Atención! Debes estar en la hoja **'${NOMBRES_HOJAS.SERVICIOS}'** para poder iniciar un presupuesto desde un servicio.`, ui.ButtonSet.OK);
    Logger.log(`Advertencia: 'presupuestar()' se intentó ejecutar desde la hoja incorrecta (${HOJA_ACTIVA.getActiveSheet().getName()}).`);
    return;
  }

  const rangoActivo = HOJA_ACTIVA.getActiveSheet().getActiveRange();

  if (!rangoActivo || rangoActivo.getNumRows() !== 1 || rangoActivo.getRow() <= 1) {
    ui.alert("Advertencia", "Por favor, selecciona **una única fila válida** (que no sea la cabecera) en la hoja **'SERVICIOS'** para crear el presupuesto.", ui.ButtonSet.OK);
    Logger.log("Advertencia: Selección inválida para presupuestar (no es una fila única o es la cabecera).");
    return;
  }

  const filaSeleccionada = rangoActivo.getRow();

  try {
    const datosFilaServicio = hojaServicios.getRange(filaSeleccionada, 1, 1, hojaServicios.getLastColumn()).getValues()[0];
    Logger.log(`Datos obtenidos de la fila ${filaSeleccionada} de Servicios para presupuestar.`);

    hojaPresupuesto.getRange("A5").setValue((datosFilaServicio[INDICES_DATOS_SERVICIOS.SERVICIO] || ""));
    hojaPresupuesto.getRange("A3").setValue((datosFilaServicio[INDICES_DATOS_SERVICIOS.DIRECCION] || "") + ", " + (datosFilaServicio[INDICES_DATOS_SERVICIOS.LOCALIDAD] || ""));
    hojaPresupuesto.getRange("A2").setValue((datosFilaServicio[INDICES_DATOS_SERVICIOS.NOMBRE] || "") + " - " + (datosFilaServicio[INDICES_DATOS_SERVICIOS.DNI] || ""));
    hojaPresupuesto.getRange("A4").setValue(datosFilaServicio[INDICES_DATOS_SERVICIOS.ASEGURADO] || "");

    // Cambio aquí: Se reemplaza ui.alert por un aviso temporal (toast)
    SpreadsheetApp.getActiveSpreadsheet().toast("Los datos del servicio han sido cargados en la plantilla de presupuesto. Ahora puedes completarlo.", "Presupuesto Iniciado", 5);
    Logger.log(`Servicio con referencia '${datosFilaServicio[INDICES_DATOS_SERVICIOS.SERVICIO]}' cargado exitosamente en la plantilla de presupuesto.`);

  } catch (error) {
    ui.alert('Error al presupuestar servicio', `Ocurrió un problema al cargar el servicio en la plantilla de presupuesto: ${error.message}. Por favor, inténtalo de nuevo o contacta al administrador.`, ui.ButtonSet.OK);
    Logger.log(`Error en la función 'presupuestar()': ${error.message}\nStack: ${error.stack}`);
  }
}



/**
 * Función para limpiar archivos PDF de servicios cerrados en Drive.
 * Obtiene la lista de servicios cerrados de la hoja "SERVICIOS" y elimina
 * los archivos correspondientes en la carpeta de servicios.
 */
function limpiarArchivosServicios() {
    const hojaServicios = HOJA_SERVICIOS; // Asume que HOJA_SERVICIOS es la hoja de datos
    const FOLDER_ID = IDS_CARPETAS_DRIVE.SERVICIOS; // Asume que esta constante existe

    if (!hojaServicios) {
        console.error('Error: La hoja "SERVICIOS" no se encontró.');
        return;
    }
    if (!FOLDER_ID) {
        console.error('Error: El ID de la carpeta de servicios no está definido.');
        return;
    }

    try {
        console.log('Iniciando limpieza automática de archivos de servicios cerrados...');

        // 1. Obtener los números de servicio CERRADOS
        const ultimaFila = hojaServicios.getLastRow();
        let serviciosCerrados = new Set();
        
        if (ultimaFila > FILA_CABECERA) {
            // Se leen las columnas de SERVICIO y ESTADO
            const rangoDatos = hojaServicios.getRange(
                FILA_CABECERA + 1, 
                Math.min(COLUMNAS_SERVICIOS.SERVICIO, COLUMNAS_SERVICIOS.ESTADO), 
                ultimaFila - FILA_CABECERA, 
                Math.abs(COLUMNAS_SERVICIOS.SERVICIO - COLUMNAS_SERVICIOS.ESTADO) + 1
            ).getValues();

            // Identificar los índices de las columnas SERVICIO y ESTADO dentro del rango leído
            // Esto asume que las columnas se leen de izquierda a derecha.
            // Para mayor seguridad, podríamos leer columna por columna, pero esto es más eficiente:
            const indiceServicio = COLUMNAS_SERVICIOS.SERVICIO < COLUMNAS_SERVICIOS.ESTADO ? 0 : 1;
            const indiceEstado = COLUMNAS_SERVICIOS.SERVICIO < COLUMNAS_SERVICIOS.ESTADO ? 1 : 0;


            rangoDatos.forEach(fila => {
                const numeroServicio = String(fila[indiceServicio]).trim();
                const estado = String(fila[indiceEstado]).trim().toUpperCase();

                if (estado === 'CERRADO' && numeroServicio.match(/\b\d{4}\/\d{4,7}\b/)) {
                    serviciosCerrados.add(numeroServicio);
                }
            });
        }
        
        if (serviciosCerrados.size === 0) {
            console.log('No se encontraron servicios con estado "CERRADO" para limpiar.');
            return;
        }

        console.log(`Se encontraron ${serviciosCerrados.size} servicios cerrados listos para limpieza.`);

        // 2. Buscar archivos en Drive por ID de carpeta
        const folder = DriveApp.getFolderById(FOLDER_ID);
        // Buscamos archivos PDF (o el formato que uses)
        const archivos = folder.getFilesByType(MimeType.PDF); 
        let archivosEliminados = 0;
        let archivosMantenidos = 0;

        // Patrón para extraer el número de servicio del nombre del archivo (ej. "2024/12345.pdf")
        const regexNombreServicio = /(\d{4}\/\d{4,7})/;
        
        // 3. Iterar y eliminar
        while (archivos.hasNext()) {
            const archivo = archivos.next();
            const nombreArchivo = archivo.getName();
            
            const match = nombreArchivo.match(regexNombreServicio);

            if (match) {
                const numeroServicioArchivo = match[1];

                if (serviciosCerrados.has(numeroServicioArchivo)) {
                    try {
                        archivo.setTrashed(true);
                        console.log(`Archivo eliminado (CERRADO): ${nombreArchivo}`);
                        archivosEliminados++;
                    } catch (error) {
                        console.error(`Error eliminando archivo ${nombreArchivo}: ${error.message}`);
                    }
                } else {
                    archivosMantenidos++;
                }
            } else {
                // Opcional: Manejar archivos en la carpeta que no tienen el formato de servicio esperado
                archivosMantenidos++;
            }
        }
        
        console.log(`Limpieza completada. Total de archivos eliminados: ${archivosEliminados}. Archivos mantenidos/ignorados: ${archivosMantenidos}`);
        
    } catch (error) {
        console.error(`Error en limpieza automática: ${error.message}\nStack: ${error.stack}`);
    }
}