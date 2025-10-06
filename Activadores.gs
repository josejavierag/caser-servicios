/**
 * Función maestra que se ejecuta cada minuto para llamar a otras funciones en secuencia.
 * Incluye manejo de errores para que un fallo en una función no detenga las demás.
 */
function ejecutarTareasCadaMinuto() {
  Logger.log("✅ Iniciando la ejecución de tareas programadas...");
  // --- Tarea 1: Etiquetar y Archivar Correos de Caser ---
  try {
    Logger.log("  -> Ejecutando: etiquetarYArchivarCorreos");
    etiquetarYArchivarCorreos();
  } catch (error) {
    Logger.log(`  ❌ ERROR en etiquetarYArchivarCorreos: ${error.message}\nStack: ${error.stack}`);
  }

  // --- Tarea 2: Cargar Servicios ---
  try {
    Logger.log("  -> Ejecutando: cargarServicios");
    cargarServicios();
  } catch (error) {
    Logger.log(`  ❌ ERROR en cargarServicios: ${error.message}\nStack: ${error.stack}`);
  }

  // --- Tarea 3: Cargar Citas ---
  try {
    Logger.log("  -> Ejecutando: cargarCitas");
    cargarCitas();
  } catch (error) {
    Logger.log(`  ❌ ERROR en cargarCitas: ${error.message}\nStack: ${error.stack}`);
  }

  // --- Tarea 4: Cargar Seguimientos ---
  try {
    Logger.log("  -> Ejecutando: cargarSeguimientos");
    cargarSeguimientos();
  } catch (error) {
    Logger.log(`  ❌ ERROR en cargarSeguimientos: ${error.message}\nStack: ${error.stack}`);
  }

  Logger.log("🏁 Todas las tareas han sido procesadas.");
}

/**
 * Crea un menú personalizado en Google Sheets para abrir la barra lateral de herramientas
 * y la abre automáticamente al cargar la hoja.
 */
function onOpen() {
    const ui = SpreadsheetApp.getUi();
    
    // Muestra la barra lateral automáticamente al abrir la hoja de cálculo
    mostrarSidebar();
}

/**
 * Muestra la barra lateral de herramientas con los botones de acción.
 */
function mostrarSidebar() {
    const html = HtmlService.createHtmlOutputFromFile('Sidebar')
        .setTitle('Panel de Control');
    SpreadsheetApp.getUi().showSidebar(html);
}