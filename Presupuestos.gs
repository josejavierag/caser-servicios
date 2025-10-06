/**
 * Inserta una nueva fila de partida en la hoja 'PRESUPUESTO'
 * rellenándola con los datos y fórmulas de la fila de plantilla.
 * @summary Versión optimizada de la función para mayor velocidad.
 */
function nuevaFilaPresupuesto() {
  const ui = SpreadsheetApp.getUi();
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  const hojaPresupuesto = libro.getSheetByName(NOMBRES_HOJAS.PRESUPUESTO);

  if (!hojaPresupuesto) {
    ui.alert("Error", `La hoja '${NOMBRES_HOJAS.PRESUPUESTO}' no se encontró. Verifica el nombre.`, ui.ButtonSet.OK);
    return;
  }

  if (libro.getActiveSheet().getName() !== hojaPresupuesto.getName()) {
    ui.alert("Advertencia", `Debes ejecutar esta acción desde la hoja '${NOMBRES_HOJAS.PRESUPUESTO}'.`, ui.ButtonSet.OK);
    return;
  }

  try {
    const ultimaFila = hojaPresupuesto.getLastRow();
    const filaNueva = ultimaFila + 1;

    // 1. Insertar la fila de una vez
    hojaPresupuesto.insertRowAfter(ultimaFila);
    
    // 2. Definir los valores y fórmulas a insertar en un arreglo
    const rangoEPIS = libro.getSheetByName("EPIS").getRange("D2:D");
    const formulas = [
      [`=IFERROR(INDEX(EPIS!$C$2:$C;MATCH(A${filaNueva}; EPIS!$D$2:$D;0));0)`, 
       `=PRODUCT(C${filaNueva};D${filaNueva})`]
    ];

    // 3. Escribir todos los datos en un solo lote para agilizar
    // Se inserta el valor '1' y las fórmulas de las columnas D y E
    hojaPresupuesto.getRange(filaNueva, 3, 1, 3).setFormulas([[`1`, formulas[0][0], formulas[0][1]]]);
    
    // 4. Establecer la validación de datos en la columna A
    const reglaValidacion = SpreadsheetApp.newDataValidation()
      .requireValueInRange(rangoEPIS)
      .setAllowInvalid(true)
      .build();
    hojaPresupuesto.getRange(`A${filaNueva}`).setDataValidation(reglaValidacion);
    
    // 5. Activar la celda A y notificar al usuario
    hojaPresupuesto.getRange(`A${filaNueva}`).activate();
    libro.toast('Se ha añadido una nueva fila para la partida.', 'Nueva Fila', 3);
    Logger.log("Nueva fila añadida de forma optimizada.");
  } catch (error) {
    const mensajeError = `No se pudo añadir la nueva fila: ${error.message}`;
    ui.alert("Error", mensajeError, ui.ButtonSet.OK);
    Logger.log("Error en nuevaFilaPresupuesto: " + error.toString());
  }
}

/**
 * Restablece la hoja "PRESUPUESTO" a su estado de plantilla por defecto.
 */
function reiniciarPlantilla() {
    const hojaPresupuesto = HOJA_PRESUPUESTO;
    if (!hojaPresupuesto) {
        Logger.log("Error: La hoja 'PRESUPUESTO' no se encontró. Verifica el nombre.");
        return;
    }
    if (HOJA_ACTIVA.getActiveSheet().getName() !== NOMBRES_HOJAS.PRESUPUESTO) {
        Logger.log("Advertencia: Se debe ejecutar esta acción desde la hoja de presupuesto.");
        return;
    }
    try {
        const hoy = new Date();
        const zonaHoraria = HOJA_ACTIVA.getSpreadsheetTimeZone();
        const valoresDefectoCabecera = {
            "A2": "NOMBRE (CIF/DNI)",
            "A3": "DIRECCION, LOCALIDAD",
            "A4": "(34) 000000000",
            "A5": "REFERENCIA:",
            "A8": "",
            "B2": "PRESUPUESTO",
            "B3": '=CONCATENATE(DEC2HEX(RANDBETWEEN(0;4294967295);8);"-";DEC2HEX(RANDBETWEEN(0;65535);4);"-";DEC2HEX(RANDBETWEEN(0;65535);4);"-";DEC2HEX(RANDBETWEEN(0;65535);4);"-";DEC2HEX(RANDBETWEEN(0;16777215);6);DEC2HEX(RANDBETWEEN(0;16777215);6))', // <<-- FÓRMULA CORREGIDA
            "B4": Utilities.formatDate(hoy, zonaHoraria, "dd/MM/yyyy")
        };
        for (const celda in valoresDefectoCabecera) {
            hojaPresupuesto.getRange(celda).clearDataValidations();
            hojaPresupuesto.getRange(celda).setValue(valoresDefectoCabecera[celda]);
        }
        hojaPresupuesto.getRange("A8").clearContent();
        const ultimaFila = hojaPresupuesto.getLastRow();
        if (ultimaFila > CONFIG_PRESUPUESTO.FILA_PLANTILLA) {
            const numFilasToDelete = ultimaFila - CONFIG_PRESUPUESTO.FILA_PLANTILLA;
            hojaPresupuesto.deleteRows(CONFIG_PRESUPUESTO.FILA_PLANTILLA + 1, numFilasToDelete);
        }
        
        // -------------------------------------------------------------
        // Se añaden las fórmulas para el IVA y el total
        // -------------------------------------------------------------
        hojaPresupuesto.getRange(CONFIG_PRESUPUESTO.CELDA_FORMULA_SUMA).setFormula("=SUM($E$12:$E)");
        hojaPresupuesto.getRange(CONFIG_PRESUPUESTO.CELDA_FORMULA_IVA).setFormula(`=E7*0,21`);
        hojaPresupuesto.getRange(CONFIG_PRESUPUESTO.CELDA_FORMULA_TOTAL).setFormula(`=E7*1,21`);
        
        Logger.log("Plantilla de presupuesto restablecida correctamente.");
    } catch (error) {
        Logger.log("Error en reiniciarPlantillaPresupuesto: " + error.toString());
    }
}

/**
 * Restablece la hoja "PRESUPUESTO" a su estado de plantilla por defecto:
 * limpia datos del cliente, partidas, y restablece fórmulas y formatos.
 */
function reiniciarPlantillaPresupuesto() {
    const ui = SpreadsheetApp.getUi();
    const hojaPresupuesto = HOJA_PRESUPUESTO;
    if (!hojaPresupuesto) {
        ui.alert("Error", `La hoja "${NOMBRES_HOJAS.PRESUPUESTO}" no se encontró. Verifica el nombre.`, ui.ButtonSet.OK);
        return;
    }
    if (HOJA_ACTIVA.getActiveSheet().getName() !== NOMBRES_HOJAS.PRESUPUESTO) {
        ui.alert("Advertencia", `Debes ejecutar esta acción desde la hoja '${NOMBRES_HOJAS.PRESUPUESTO}'.`, ui.ButtonSet.OK);
        return;
    }
    try {
        const hoy = new Date();
        const zonaHoraria = HOJA_ACTIVA.getSpreadsheetTimeZone();
        const valoresDefectoCabecera = {
            "A2": "NOMBRE (CIF/DNI)",
            "A3": "DIRECCION, LOCALIDAD",
            "A4": "(34) 000000000",
            "A5": "REFERENCIA:",
            "A8": "",
            "B2": "PRESUPUESTO",
            "B3": '=CONCATENATE(DEC2HEX(RANDBETWEEN(0;4294967295);8);"-";DEC2HEX(RANDBETWEEN(0;65535);4);"-";DEC2HEX(RANDBETWEEN(0;65535);4);"-";DEC2HEX(RANDBETWEEN(0;65535);4);"-";DEC2HEX(RANDBETWEEN(0;16777215);6);DEC2HEX(RANDBETWEEN(0;16777215);6))', // <<-- FÓRMULA CORREGIDA
            "B4": Utilities.formatDate(hoy, zonaHoraria, "dd/MM/yyyy")
        };
        for (const celda in valoresDefectoCabecera) {
            hojaPresupuesto.getRange(celda).clearDataValidations();
            hojaPresupuesto.getRange(celda).setValue(valoresDefectoCabecera[celda]);
        }
        hojaPresupuesto.getRange("A8").clearContent();
        const ultimaFila = hojaPresupuesto.getLastRow();
        if (ultimaFila > CONFIG_PRESUPUESTO.FILA_PLANTILLA) {
            const numFilasToDelete = ultimaFila - CONFIG_PRESUPUESTO.FILA_PLANTILLA;
            hojaPresupuesto.deleteRows(CONFIG_PRESUPUESTO.FILA_PLANTILLA + 1, numFilasToDelete);
        }
        
        // -------------------------------------------------------------
        // Se añaden las fórmulas para el IVA y el total
        // -------------------------------------------------------------
        hojaPresupuesto.getRange(CONFIG_PRESUPUESTO.CELDA_FORMULA_SUMA).setFormula("=SUM($E$12:$E)");
        hojaPresupuesto.getRange(CONFIG_PRESUPUESTO.CELDA_FORMULA_IVA).setFormula(`=E7*0,21`);
        hojaPresupuesto.getRange(CONFIG_PRESUPUESTO.CELDA_FORMULA_TOTAL).setFormula(`=E7*1,21`);
        
        SpreadsheetApp.getActiveSpreadsheet().toast("La plantilla del presupuesto ha sido restablecida correctamente.", "Plantilla Restablecida", 3);
        Logger.log("Plantilla de presupuesto restablecida.");
    } catch (error) {
        ui.alert('Error', `No se pudo restablecer la plantilla del presupuesto: ${error.message}`, ui.ButtonSet.OK);
        Logger.log("Error en reiniciarPlantillaPresupuesto: " + error.toString());
    }
}

// CONFIGURACIÓN: Ajusta estos valores para que coincidan con tus nombres de hojas y rangos.
const RANGOS = {
  // Celdas donde se encuentran los datos de la cabecera
  CABECERA: {
    NUMERO: "B3",
    FECHA: "B4",
    SERVICIO: "A5",
    NOTAS: "A8",
    FORMATO: "B2",
    NETO: "E7",
    IVA: "E8",
    TOTAL: "D9"
  },
  // Rango donde comienzan los datos de los detalles.
  // Columna A (TARIFA), Fila 10.
  DETALLES_INICIO_FILA: 13,
  DETALLES_NUM_COLUMNAS: 5 // TARIFA, NOTA, UNIDADES, PRECIO, TOTAL_TARIFA
};

// Esta función es la que se ejecutará para guardar el presupuesto
function guardarPresupuestoDesdePlantilla() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  const hojaManual = ss.getSheetByName(NOMBRES_HOJAS.HOJA_MANUAL_PLANTILLA);
  const hojaPresupuestos = ss.getSheetByName(NOMBRES_HOJAS.HOJA_PRESUPUESTOS);
  const hojaDetalles = ss.getSheetByName(NOMBRES_HOJAS.HOJA_DETALLES);

  if (!hojaManual || !hojaPresupuestos || !hojaDetalles) {
    ui.alert("Error", "Falta una de las hojas de cálculo. Verifica los nombres en el script.", ui.ButtonSet.OK);
    return;
  }

  try {
    // 1. Leer los datos de la cabecera del presupuesto
    const numero = hojaManual.getRange(RANGOS.CABECERA.NUMERO).getValue();
    const fecha = hojaManual.getRange(RANGOS.CABECERA.FECHA).getValue();
    const servicio = hojaManual.getRange(RANGOS.CABECERA.SERVICIO).getValue();
    const notas = hojaManual.getRange(RANGOS.CABECERA.NOTAS).getValue();
    const formato = hojaManual.getRange(RANGOS.CABECERA.FORMATO).getValue();
    const neto = hojaManual.getRange(RANGOS.CABECERA.NETO).getValue();
    const iva = hojaManual.getRange(RANGOS.CABECERA.IVA).getValue();
    const total = hojaManual.getRange(RANGOS.CABECERA.TOTAL).getValue();
    
    // Validar datos básicos
    if (!numero || !fecha) {
      ui.alert("Advertencia", "Los campos 'NUMERO' y 'FECHA' no pueden estar vacíos.", ui.ButtonSet.OK);
      return;
    }

    // 2. Generar un ID único para el presupuesto (muy importante para AppSheet)
    const presupuestoID = Utilities.getUuid();

    // 3. Escribir la nueva fila en la hoja de cabecera de AppSheet
    hojaPresupuestos.appendRow([
  presupuestoID, // <-- Es VITAL que este ID esté aquí
  numero,
  formato,
  fecha,
  servicio,
  notas
]);

    // 4. Leer y procesar los detalles del presupuesto
    const ultimaFilaDetallesManual = hojaManual.getLastRow();
    const rangoDatosDetalles = hojaManual.getRange(
      RANGOS.DETALLES_INICIO_FILA, 
      1, 
      ultimaFilaDetallesManual - RANGOS.DETALLES_INICIO_FILA + 1, 
      RANGOS.DETALLES_NUM_COLUMNAS
    );
    const datosDetalles = rangoDatosDetalles.getValues().filter(fila => String(fila[0]).trim() !== "");

    if (datosDetalles.length === 0) {
      ui.alert("Advertencia", "No hay detalles para guardar en el presupuesto.", ui.ButtonSet.OK);
      return;
    }

    // 5. Escribir cada detalle en la hoja de detalles de AppSheet
    const detallesParaEscribir = datosDetalles.map(detalle => {
      return [
        Utilities.getUuid(), // ID único para cada detalle
        presupuestoID, // ID del presupuesto padre (la referencia)
        detalle[0], // TARIFA
        detalle[1], // NOTA
        detalle[2], // UNIDADES
        detalle[3], // PRECIO
        detalle[4], // TOTAL_TARIFA
      ];
    });

    hojaDetalles.getRange(hojaDetalles.getLastRow() + 1, 1, detallesParaEscribir.length, detallesParaEscribir[0].length)
      .setValues(detallesParaEscribir);
      
    SpreadsheetApp.getActiveSpreadsheet().toast("El presupuesto se ha guardado y se reflejará en AppSheet al sincronizar.", "Presupuesto Guardado", 5);
  reiniciarPlantillaPresupuesto();
  } catch (error) {
    ui.alert('Error al guardar presupuesto', `No se pudo guardar el presupuesto: ${error.message}`, ui.ButtonSet.OK);
  }
}

/**
 * Restaura los datos de un presupuesto guardado, a partir de la fila seleccionada
 * en la hoja 'PLANTILLA', y los carga en la hoja de plantilla manual.
 */
function restaurarPresupuesto() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const hojaActiva = ss.getActiveSheet();
  const hojaPresupuestos = ss.getSheetByName(NOMBRES_HOJAS.HOJA_PRESUPUESTOS);
  const hojaDetalles = ss.getSheetByName(NOMBRES_HOJAS.HOJA_DETALLES);
  const hojaManual = ss.getSheetByName(NOMBRES_HOJAS.HOJA_MANUAL_PLANTILLA);

  if (!hojaPresupuestos || !hojaDetalles || !hojaManual) {
    ui.alert("Error", "Falta una de las hojas de cálculo. Verifica los nombres en el script.", ui.ButtonSet.OK);
    return;
  }

  // Comprobar que la hoja activa es 'PLANTILLA'
  if (hojaActiva.getName() !== NOMBRES_HOJAS.HOJA_PRESUPUESTOS) {
    ui.alert("Advertencia", `Debes seleccionar una fila en la hoja '${NOMBRES_HOJAS.HOJA_PRESUPUESTOS}' para restaurar un presupuesto.`, ui.ButtonSet.OK);
    return;
  }

  // Obtener la fila seleccionada
  const rangoSeleccionado = hojaActiva.getActiveRange();
  if (!rangoSeleccionado || rangoSeleccionado.getNumRows() !== 1) {
    ui.alert("Advertencia", "Por favor, selecciona la fila completa de un único presupuesto para restaurar.", ui.ButtonSet.OK);
    return;
  }
  
  // <<-- COMPROBACIÓN AÑADIDA AQUÍ
  if (rangoSeleccionado.getRow() === 1) {
    ui.alert("Advertencia", "Por favor, no selecciones la fila de encabezados. Selecciona una fila con datos de presupuesto.", ui.ButtonSet.OK);
    return;
  }
  
  // Obtener los datos de la fila seleccionada
  const filaPresupuesto = rangoSeleccionado.getValues()[0];
  
  // Obtener los encabezados de la hoja de presupuestos para encontrar las columnas dinámicamente
  const encabezadosPresupuestos = hojaPresupuestos.getRange(1, 1, 1, hojaPresupuestos.getLastColumn()).getValues()[0];

  const columnaID = encabezadosPresupuestos.indexOf("ID");
  const columnaNumero = encabezadosPresupuestos.indexOf("NUMERO");
  const columnaFormato = encabezadosPresupuestos.indexOf("FORMATO");
  const columnaFecha = encabezadosPresupuestos.indexOf("FECHA");
  const columnaServicio = encabezadosPresupuestos.indexOf("SERVICIO");
  const columnaNotas = encabezadosPresupuestos.indexOf("NOTAS");
  const columnaNeto = encabezadosPresupuestos.indexOf("NETO");
  const columnaIVA = encabezadosPresupuestos.indexOf("IVA");
  const columnaTotal = encabezadosPresupuestos.indexOf("TOTAL");

  // Validaciones
  if (columnaID === -1 || columnaNumero === -1) {
      ui.alert("Error", "No se encontró la columna 'ID' o 'NUMERO' en la hoja de presupuestos. Verifica los nombres.", ui.ButtonSet.OK);
      return;
  }
  
  const presupuestoID = filaPresupuesto[columnaID];
  const numeroPresupuesto = filaPresupuesto[columnaNumero];

  if (!presupuestoID || !numeroPresupuesto) {
      ui.alert("Error", "La fila seleccionada no contiene un ID de presupuesto válido.", ui.ButtonSet.OK);
      return;
  }
  
  try {
    // 1. Limpiar la plantilla manual
    reiniciarPlantillaPresupuesto(); 
    
    // 2. Escribir los datos de la cabecera en la plantilla manual
    hojaManual.getRange(RANGOS.CABECERA.NUMERO).setValue(filaPresupuesto[columnaNumero]);    
    hojaManual.getRange(RANGOS.CABECERA.FORMATO).setValue(filaPresupuesto[columnaFormato]);
    hojaManual.getRange(RANGOS.CABECERA.FECHA).setValue(filaPresupuesto[columnaFecha]);
    hojaManual.getRange(RANGOS.CABECERA.SERVICIO).setValue(filaPresupuesto[columnaServicio]);
    hojaManual.getRange(RANGOS.CABECERA.NOTAS).setValue(filaPresupuesto[columnaNotas]);
    hojaManual.getRange(RANGOS.CABECERA.NETO).setValue(filaPresupuesto[columnaNeto]);
    hojaManual.getRange(RANGOS.CABECERA.IVA).setValue(filaPresupuesto[columnaIVA]);
    hojaManual.getRange(RANGOS.CABECERA.TOTAL).setValue(filaPresupuesto[columnaTotal]);

    // 3. Buscar y escribir los detalles
    const datosDetalles = hojaDetalles.getDataRange().getValues();
    const encabezadosDetalles = datosDetalles.shift();

    const columnaReferencia = encabezadosDetalles.indexOf("ID_PRESUPUESTO_PADRE"); 
    if (columnaReferencia === -1) {
      ui.alert("Error", "No se encontró la columna 'ID_PRESUPUESTO_PADRE' en la hoja de detalles.", ui.ButtonSet.OK);
      return;
    }

    const detallesAsociados = datosDetalles.filter(detalle => detalle[columnaReferencia] === presupuestoID);

    if (detallesAsociados.length > 0) {
        const columnaTarifa = encabezadosDetalles.indexOf("TARIFA");
        const columnaNota = encabezadosDetalles.indexOf("NOTA");
        const columnaUnidades = encabezadosDetalles.indexOf("UNIDADES");
        const columnaPrecio = encabezadosDetalles.indexOf("PRECIO");
        const columnaTotalTarifa = encabezadosDetalles.indexOf("TOTAL_TARIFA");

      const detallesParaEscribir = detallesAsociados.map(detalle => [
        detalle[columnaTarifa],
        detalle[columnaNota],
        detalle[columnaUnidades],
        detalle[columnaPrecio],
        detalle[columnaTotalTarifa]
      ]);

      const rangoDestino = hojaManual.getRange(
        RANGOS.DETALLES_INICIO_FILA, 
        1, 
        detallesParaEscribir.length, 
        detallesParaEscribir[0].length
      );
      rangoDestino.setValues(detallesParaEscribir);
    }
    
    SpreadsheetApp.getActiveSpreadsheet().toast(`El presupuesto Nº ${numeroPresupuesto} ha sido cargado correctamente.`, "Presupuesto Cargado", 5);

  } catch (error) {
    ui.alert('Error al cargar presupuesto', `No se pudo cargar el presupuesto: ${error.message}`, ui.ButtonSet.OK);
  }
}

/**
 * Se activa con la acción de AppSheet y coordina el proceso de rellenado.
 * Utiliza un parámetro 'accion' para determinar qué función ejecutar.
 * @param {object} e - El objeto de evento de la solicitud GET.
 */
function doGet(e) {
  try {
    const accion = e.parameter.accion;
    if (!accion) {
      throw new Error("Parámetro 'accion' no proporcionado.");
    }

    switch (accion) {
      case "generarPresupuestoSimple":
        // Lógica para el primer botón
        reiniciarPlantilla();
        
        const datos = e.parameter;
        const servicio = datos.servicio;
        const asegurado = datos.asegurado;
        const direccion = datos.direccion;
        const localidad = datos.localidad;
        const nombre = datos.nombre || "";
        const dni = datos.dni || "";

        const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
        const hojaPresupuesto = spreadsheet.getSheetByName("PRESUPUESTO");
        
        if (!hojaPresupuesto) {
          throw new Error("La hoja 'PRESUPUESTO' no se encontró.");
        }
        
        hojaPresupuesto.getRange("A5").setValue(servicio);
        hojaPresupuesto.getRange("A4").setValue(asegurado);
        hojaPresupuesto.getRange("A3").setValue(direccion + ", " + localidad);
        hojaPresupuesto.getRange("A2").setValue(nombre + " " + dni);
        
        return HtmlService.createHtmlOutput("Presupuesto simple generado con éxito.");

      case "generarPresupuestoCompleto":
        // Lógica para el segundo botón
        const idPlantilla = e.parameter.id;
        if (!idPlantilla) {
          throw new Error("ID de plantilla no proporcionado para la acción 'generarPresupuestoCompleto'.");
        }
        
        const datosPresupuesto = obtenerDatosDeAppSheet(idPlantilla);
        reiniciarPlantilla();
        rellenarPresupuesto(datosPresupuesto);
        
        return HtmlService.createHtmlOutput("Presupuesto completo generado con éxito.");

      case "generarDocumentoServicio":
        // Lógica para generar documento de servicio desde plantilla DOC
        const idRegistro = e.parameter.id;
        if (!idRegistro) {
          throw new Error("ID de registro no proporcionado para la acción 'generarDocumentoServicio'.");
        }
        
        return generarDocumentoServicio(e.parameter);

      default:
        throw new Error(`Acción '${accion}' no reconocida.`);
    }

  } catch (error) {
    Logger.log("Error en la ejecución: " + error.message);
    return HtmlService.createHtmlOutput("Error: " + error.message);
  }
}

/**
 * Genera un documento de Google Docs desde una plantilla con datos del servicio
 * @param {object} params - Parámetros de la solicitud
 */
function generarDocumentoServicio(params) {
  try {
    const TEMPLATE_ID = '1IWMq77wf3_sapsNeezpbfwpc2c0Ceak-oMaSsjHYot8';
    const registroId = params.id;
    
    // Obtener datos del registro de SERVICIOS
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName('SERVICIOS');
    
    if (!sheet) {
      throw new Error("La hoja 'SERVICIOS' no se encontró.");
    }
    
    const data = obtenerDatosServicioPorId(sheet, registroId);
    
    if (!data) {
      throw new Error('No se encontró el registro con ID: ' + registroId);
    }
    
    // Crear documento desde plantilla
    const templateDoc = DriveApp.getFileById(TEMPLATE_ID);
    const newDoc = templateDoc.makeCopy(`Servicio_${data.expediente}_${Utilities.formatDate(new Date(), 'GMT+1', 'yyyyMMdd')}`);
    const doc = DocumentApp.openById(newDoc.getId());
    
// Reemplazar marcadores en la plantilla
const body = doc.getBody();
const fechaHoy = new Date();
fechaHoy.setDate(fechaHoy.getDate() +1);
const fechaManana = Utilities.formatDate(fechaHoy, 'Europe/Madrid', 'dd/MM/yyyy');

// Reemplazar en el cuerpo del documento
body.replaceText('<<\\[SERVICIO\\]>>', data.expediente || '');
body.replaceText('<<\\[DIRECCION\\]>>', data.direccion || '');
body.replaceText('<<\\[LOCALIDAD\\]>>', data.localidad || '');
body.replaceText('<<\\[NOMBRE\\]>>', data.nombre || '');
body.replaceText('<<\\[ASEGURADO\\]>>', data.asegurado || '');
body.replaceText('<<\\[FECHA_HOY\\]>>', fechaManana);

// Reemplazar en el pie de página
const footer = doc.getFooter();
if (footer) {
  footer.replaceText('<<\\[SERVICIO\\]>>', data.expediente || '');
  footer.replaceText('<<\\[DIRECCION\\]>>', data.direccion || '');
  footer.replaceText('<<\\[LOCALIDAD\\]>>', data.localidad || '');
  footer.replaceText('<<\\[NOMBRE\\]>>', data.nombre || '');
  footer.replaceText('<<\\[ASEGURADO\\]>>', data.asegurado || '');
  footer.replaceText('<<\\[FECHA_HOY\\]>>', fechaManana);
}

// Reemplazar en el encabezado también (por si acaso)
const header = doc.getHeader();
if (header) {
  header.replaceText('<<\\[SERVICIO\\]>>', data.expediente || '');
  header.replaceText('<<\\[DIRECCION\\]>>', data.direccion || '');
  header.replaceText('<<\\[LOCALIDAD\\]>>', data.localidad || '');
  header.replaceText('<<\\[NOMBRE\\]>>', data.nombre || '');
  header.replaceText('<<\\[ASEGURADO\\]>>', data.asegurado || '');
  header.replaceText('<<\\[FECHA_HOY\\]>>', fechaManana);
}
    
    // Guardar cambios
    doc.saveAndClose();
    
    // Generar PDF para imprimir
    const blob = DriveApp.getFileById(newDoc.getId()).getAs('application/pdf');
    const pdfFile = DriveApp.createFile(blob);
    pdfFile.setName(`Servicio_${data.expediente}_${fechaManana.replace(/\//g, '')}.pdf`);
    
    // Retornar respuesta que abre el PDF para imprimir
    return HtmlService.createHtmlOutput(`
      <script>
        window.open('${pdfFile.getUrl()}', '_blank');
        setTimeout(() => window.close(), 1000);
      </script>
    `);
    
  } catch (error) {
    Logger.log('Error generando documento de servicio: ' + error.message);
    // Reemplaza la parte final del return por esto:
return HtmlService.createHtmlOutput(`
  <script>
    // Redirigir directamente al PDF sin mostrar la página intermedia
    window.location.replace('${pdfFile.getUrl()}');
  
  </script>
`);
  }
}


/**
 * Obtiene los datos de un servicio específico por su campo SERVICIO
 * @param {Sheet} sheet - La hoja de cálculo SERVICIOS
 * @param {string} servicioId - El valor del campo SERVICIO a buscar
 * @return {object} Los datos del servicio o null si no se encuentra
 */
function obtenerDatosServicioPorId(sheet, servicioId) {
  try {
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const servicioIndex = headers.indexOf('SERVICIO');
    
    if (servicioIndex === -1) {
      throw new Error("No se encontró la columna 'SERVICIO' en la hoja.");
    }
    
    // Buscar el registro por el campo SERVICIO
    for (let i = 1; i < data.length; i++) {
      if (data[i][servicioIndex].toString() === servicioId.toString()) {
        return {
          expediente: data[i][servicioIndex] || '',
          direccion: data[i][headers.indexOf('DIRECCION')] || '', 
          localidad: data[i][headers.indexOf('LOCALIDAD')] || '',
          nombre: data[i][headers.indexOf('NOMBRE')] || '',
          asegurado: data[i][headers.indexOf('ASEGURADO')] || ''
        };
      }
    }
    return null;
  } catch (error) {
    Logger.log('Error obteniendo datos del servicio: ' + error.message);
    return null;
  }
}

/**
 * Obtiene el índice de una columna por su nombre de encabezado.
 * Lanza un error si la columna no se encuentra.
 */
function findHeaderIndex(headers, headerName) {
  const index = headers.indexOf(headerName);
  if (index === -1) {
    throw new Error(`Columna '${headerName}' no encontrada en el encabezado de la hoja.`);
  }
  return index;
}

/**
 * Obtiene todos los datos necesarios de las tablas de AppSheet.
 */
function obtenerDatosDeAppSheet(idPlantilla) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const hojaPlantilla = spreadsheet.getSheetByName(NOMBRES_HOJAS.PLANTILLA);
  const hojaServicios = spreadsheet.getSheetByName(NOMBRES_HOJAS.SERVICIOS);
  const hojaDetalles = spreadsheet.getSheetByName(NOMBRES_HOJAS.DETALLES);

  if (!hojaPlantilla || !hojaServicios || !hojaDetalles) {
    throw new Error("Una o más hojas de datos no se encontraron.");
  }

  const headersPlantilla = hojaPlantilla.getRange(1, 1, 1, hojaPlantilla.getLastColumn()).getValues()[0];
  const headersServicios = hojaServicios.getRange(1, 1, 1, hojaServicios.getLastColumn()).getValues()[0];
  const headersDetalles = hojaDetalles.getRange(1, 1, 1, hojaDetalles.getLastColumn()).getValues()[0];

  // Encuentra los índices de las columnas clave
  const indicePresupuesto = findHeaderIndex(headersPlantilla, "PRESUPUESTO");
  const indiceServicio = findHeaderIndex(headersServicios, "SERVICIO");
  const indiceDetalleRef = findHeaderIndex(headersDetalles, "PRESUPUESTO");

  // Obtiene los datos de cada hoja
  const datosHojaPlantilla = hojaPlantilla.getDataRange().getValues();
  const datosHojaServicios = hojaServicios.getDataRange().getValues();
  const datosHojaDetalles = hojaDetalles.getDataRange().getValues();

  // 1. Obtener datos de la fila de PLANTILLA
  const datosPlantilla = datosHojaPlantilla.find(row => row[indicePresupuesto] === idPlantilla);
  if (!datosPlantilla) {
    throw new Error(`Registro de plantilla con ID '${idPlantilla}' no encontrado.`);
  }
  const plantilla = {};
  headersPlantilla.forEach((header, i) => plantilla[header] = datosPlantilla[i]);
  
  // 2. Obtener datos del servicio asociado
  const idServicio = plantilla.SERVICIO;
  const datosServicio = datosHojaServicios.find(row => row[indiceServicio] === idServicio);
  if (!datosServicio) {
    throw new Error(`Registro de servicio con ID '${idServicio}' no encontrado.`);
  }
  const servicio = {};
  headersServicios.forEach((header, i) => servicio[header] = datosServicio[i]);

  // 3. Obtener todos los detalles asociados
  const detallesFiltrados = datosHojaDetalles.filter(row => row[indiceDetalleRef] === idPlantilla);
  const detalles = detallesFiltrados.map(row => {
    const detalle = {};
    headersDetalles.forEach((header, i) => detalle[header] = row[i]);
    return detalle;
  });

  return {
    plantilla: plantilla,
    servicio: servicio,
    detalles: detalles
  };
}

/**
 * Rellena la hoja de presupuesto con los datos obtenidos.
 * @param {object} datos - El objeto con los datos de plantilla, servicio y detalles.
 */
function rellenarPresupuesto(datos) {
  const hojaPresupuesto = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOMBRES_HOJAS.PRESUPUESTO);
  
  // 1. Rellenar encabezado del presupuesto
  hojaPresupuesto.getRange("A2").setValue(datos.servicio.NOMBRE + " " + datos.servicio.DNI);
  hojaPresupuesto.getRange("A3").setValue(datos.servicio.DIRECCIÓN + ", " + datos.servicio.LOCALIDAD);
  hojaPresupuesto.getRange("A4").setValue(datos.servicio.TELEFONO);
  hojaPresupuesto.getRange("A5").setValue(datos.servicio.SERVICIO);

  hojaPresupuesto.getRange("B2").setValue(datos.plantilla.FORMATO);
  hojaPresupuesto.getRange("B3").setValue(datos.plantilla.NUMERO);
  hojaPresupuesto.getRange("B4").setValue(datos.plantilla.FECHA);
  
  // 2. Insertar y rellenar los detalles de las partidas
  if (datos.detalles.length > 0) {
    const filaInicio = 13;
    const numFilas = datos.detalles.length;
    
    // Si hay más de una fila, inserta las filas necesarias
    if (numFilas > 1) {
      hojaPresupuesto.insertRowsAfter(filaInicio - 1, numFilas - 1);
    }
    
    // Preparar los datos para escribir en el rango
    const datosTabla = datos.detalles.map(detalle => [
      detalle.TARIFA,
      detalle.NOTA,
      detalle.UNIDADES,
      detalle.PRECIO,
      detalle.TOTAL_TARIFA
    ]);
    
    // Escribir los datos en el rango
    hojaPresupuesto.getRange(filaInicio, 1, numFilas, 5).setValues(datosTabla);
  }
}