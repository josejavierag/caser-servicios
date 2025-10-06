const ID_HOJA_CALCULO_PRINCIPAL = "18jEmhBlbZUJkqbrWENz_vC8Nhxm1uazZuU0jzEb41hM";
const HOJA_ACTIVA = SpreadsheetApp.getActiveSpreadsheet();
const NOMBRES_HOJAS = {
    SERVICIOS: "SERVICIOS",
    CODIGOS_POSTALES: "CCPP",
    OPERARIOS: "OPERARIOS",
    REGISTRO_PRESUPUESTOS: "REGISTRO",
    FACTURAS: "FACTURAS",
    CERRADOS: "CERRADOS",
    PRESUPUESTO: "PRESUPUESTO",
    EXTERNO: "EXTERNO",
    EPIS: "EPIS",
    PAGOS: "PAGOS",
    SEGUIMIENTOS: "SEGUIMIENTOS",
    HOJA_MANUAL_PLANTILLA: "PRESUPUESTO",
    HOJA_PRESUPUESTOS: "PLANTILLA",
    HOJA_DETALLES: "DETALLES",
    PLANTILLA: "PLANTILLA",
    DETALLES: "DETALLES"
};
const HOJA_SERVICIOS = HOJA_ACTIVA.getSheetByName(NOMBRES_HOJAS.SERVICIOS);
const HOJA_CODIGOS_POSTALES = HOJA_ACTIVA.getSheetByName(NOMBRES_HOJAS.CODIGOS_POSTALES);
const HOJA_OPERARIOS = HOJA_ACTIVA.getSheetByName(NOMBRES_HOJAS.OPERARIOS);
const HOJA_REGISTRO_PRESUPUESTOS = HOJA_ACTIVA.getSheetByName(NOMBRES_HOJAS.REGISTRO_PRESUPUESTOS);
const HOJA_FACTURAS = HOJA_ACTIVA.getSheetByName(NOMBRES_HOJAS.FACTURAS);
const HOJA_CERRADOS = HOJA_ACTIVA.getSheetByName(NOMBRES_HOJAS.CERRADOS);
const HOJA_PRESUPUESTO = HOJA_ACTIVA.getSheetByName(NOMBRES_HOJAS.PRESUPUESTO);
const HOJA_EXTERNO = HOJA_ACTIVA.getSheetByName(NOMBRES_HOJAS.EXTERNO);
const HOJA_EPIS = HOJA_ACTIVA.getSheetByName(NOMBRES_HOJAS.EPIS);
const HOJA_PAGOS = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOMBRES_HOJAS.PAGOS);
const IDS_CARPETAS_DRIVE = {
    SEGUIMIENTOS: "1kSill56rB4BrkHODoKZfu_bHCbREYAL2",
    SERVICIOS: "1GMkb-r0rpjJKerOuIr3JE_95GZ2DSjo1"
};

const COLUMNAS_SERVICIOS = {
    ALTA: 1,
    PLATAFORMA: 2,
    SERVICIO: 3,
    POLIZA: 4,
    OBSERVACIONES: 5,
    FECHA: 6,
    HORA: 7,
    OPERARIO: 8,
    ESTADO: 9,
    DIRECCION: 10,
    LOCALIDAD: 11,
    DESCRIPCION: 12,
    ASEGURADO: 13,
    INQUILINO: 14,
    AFECTADO: 15,
    OTRO: 16,
    OTRO2: 17,
    DNI: 18,
    NOMBRE: 19,
    PDF: 20,
};

COLUMNAS_SEGUIMIENTOS = {
  SEGUIMIENTO: 1,
  CUERPO: 2,
  FECHA: 3,
  SERVICIO: 4
}

const HOJA_SEGUIMIENTOS = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOMBRES_HOJAS.SEGUIMIENTOS);

const ULTIMA_COLUMNA_DATOS_SERVICIOS = 20;

const INDICES_DATOS_SERVICIOS = {
    ALTA: 0,
    PLATAFORMA: 1,
    SERVICIO: 2,
    POLIZA: 3,
    OBSERVACIONES: 4,
    FECHA: 5,
    HORA: 6,
    OPERARIO: 7,
    ESTADO: 8,
    DIRECCION: 9,
    LOCALIDAD: 10,
    DESCRIPCION: 11,
    ASEGURADO: 12,
    INQUILINO: 13,
    AFECTADO: 14,
    OTRO: 15,
    OTRO2: 16,
    DNI: 17,
    NOMBRE: 18,
    PDF: 19,
};

const INDICES_DATOS_PLANTILLA = {
  PRESUPUESTO: 0,
  NUMERO: 1,
  FORMATO: 2,
  FECHA: 3,
  SERVICIO: 4,
  NOTAS: 5,
  IMPORTE_NETO: 6,
  IVA: 7,
  TOTAL: 8
}
const COLUMNAS_FACTURAS = {
    NUMERO: 1,
    PLATAFORMA: 2,
    SERVICIO: 3,
    FECHA: 4,
    ESTADO: 5,
    TOTAL: 6,
    DESCRIPCION: 7,
    DIRECCION: 8,
    LOCALIDAD: 9
};

const COLUMNAS_EPIS = {
    EPIGRAFE: 1,
    DESCRIPCION: 2,
    PRECIO: 3,
    EPI_DESC: 4,
    LEYENDA: 5
};
const COLUMNAS_REGISTRO_PRESUPUESTOS = {
    REFERENCIA: 1,
    NOMBRE: 2,
    DIRECCION: 3,
    ASEGURADO: 4,
    FORMATO: 5,
    NUMERO: 6,
    FECHA: 7,
    NOTAS: 8,
    INICIO_PARTIDAS: 9
};
const COLUMNAS_OPERARIOS = {
    OPERARIO: 1,
    SALARIO: 2,
    VEHICULO: 3,
    DIETA: 4,
    SALDO: 5
};
// Constante que faltaba
const COLUMNAS_PAGOS = {
    OPERARIO: 1,
    FECHA: 2,
    SALARIO: 3,
    VEHICULO: 4,
    DIETA: 5,
    ANTICIPO: 6,
    USA_VEHICULO: 7,
    TOTAL: 8
};
// Constantes adicionales necesarias para que el código funcione
const COL_OPERARIOS = {
    OPERARIO: 1,
    SALARIO: 2,
    VEHICULO: 3,
    DIETA: 4,
    SALDO_TOTAL: 5
};
const FILA_CABECERA = 1;
const CONFIG_PRESUPUESTO = {
    FILA_PLANTILLA: 12, // Fila donde empieza la plantilla de partidas
    FILA_INICIO_PARTIDAS: 13, // Fila donde se insertan las nuevas partidas
    CELDA_FORMULA_SUMA: "E7",
    CELDA_FORMULA_IVA: "E8",
    CELDA_FORMULA_TOTAL: "D9"
};
const ENCABEZADO_PRESUPUESTO = {
    CELDAS: ["A5", "A2", "A3", "A4", "B2", "B3", "B4", "A8"], // Celdas del encabezado en Presupuesto
    INDICES: {
        REFERENCIA: 0,
        NOMBRE_DNI: 1,
        DIRECCION: 2,
        TELEFONO: 3,
        NUMERO_PRESUPUESTO: 4,
        CAMPO_B3: 5,
        FECHA: 6,
        OBSERVACIONES: 7
    }
};