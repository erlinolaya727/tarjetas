/*Desarrollador: Ing. Erlin Olaya Rojao
Funcionalidad: App Script encargado de revisar la capeta del google drive donde son alojados los reportes, realizar transformación de xls a google sheet, leer los datos cargados y validar los tipo de archivos, procesar la información, unificar el consolidado y enviar correo electrónico de confirmación
 */
function reporteTarjetas() {

  //enviarTablaTrabajadores();


  Logger.log("Inicia ejecución de App Script");

  // Se inicializa cuerpo de respuesta
  const body = {}
  let response = {}


  try {


    Logger.log("Inicia procesarReporteGoogleSheet()");
    procesarReporteGoogleSheet();
    Logger.log("Terminó procesarReporteGoogleSheet()");

    var sheetGoogle = DriveApp.getFoldersByName(Params.REPORTES_PROCESADOS.Folder_Name);
    var folder = sheetGoogle.next();
    var archivos = folder.getFiles();

    Logger.log("Se iniciar a Iterar por cada archivo encontrado");
    if (archivos.hasNext()) {
      while (archivos.hasNext()) {
        Logger.log("Inicia obtenerDataGoogleSheet()");
        response = obtenerDataGoogleSheet(archivos);
        Logger.log("Terminó  obtenerDataGoogleSheet()");

        const dataSheet = response.body.dataSheetGoogle;
        const nombreTrabajador = response.body.nombreTrabajador;
        const archivo = response.body.archivoEncontrado;
        lastRow = response.body.lastRow;

        const dataSheetConEmcabezado = response.body.dataSheetGoogleEncabezados;

        Logger.log("Fecha: " + dataSheetConEmcabezado[0][0] + "||" + dataSheetConEmcabezado[1][0] + "||" + dataSheetConEmcabezado[2][0]);
        if (dataSheetConEmcabezado[0][0] == "Fecha" || dataSheetConEmcabezado[1][0] == "Fecha" || dataSheetConEmcabezado[2][0] == "Fecha") {

          Logger.log("Pozo: " + dataSheetConEmcabezado[0][1] + "||" + dataSheetConEmcabezado[1][1] + "||" + dataSheetConEmcabezado[2][1]);
          if (dataSheetConEmcabezado[0][1] == "Pozo" || dataSheetConEmcabezado[1][1] == "Pozo" || dataSheetConEmcabezado[2][1] == "Pozo") {

            Logger.log("Rig: " + dataSheetConEmcabezado[0][2] + "||" + dataSheetConEmcabezado[1][2] + "||" + dataSheetConEmcabezado[2][2]);
            if (dataSheetConEmcabezado[0][2] == "RIG" || dataSheetConEmcabezado[1][2] == "RIG" || dataSheetConEmcabezado[2][2] == "RIG") {
              Logger.log("El archivo corresponde al reporte de tarjetas");
              Logger.log("Se encontrò un archivo: " + archivo + " para el trabajador: " + nombreTrabajador);
              Logger.log("Se valida si el reporte tien màs columnas de las debidas");
              if (dataSheet[0].length < 15) {

                Logger.log("El archivo tiene las columnas adecuadas");
                response = validarRegistrosDuplicados(dataSheet);

                const bln_Duplicado = response.body.bln_Duplicado;
                const mensajeDuplicado = response.body.mensajeDuplicado;

                Logger.log("Se valida si existen duplicados");
                if (bln_Duplicado == 'false') {
                  Logger.log("No existen registros duplicados");

                  //Se llama a funciòn para agregar las nuevas filas
                  Logger.log("Se llama a funciòn para agregar las nuevas consolidarInformacion(dataSheet)");
                  consolidarInformacion(dataSheet);

                  var fecha = new Date();
                  var year = fecha.getFullYear();
                  var mes = fecha.getMonth() + 1;

                  archivo.setName(year + "-" + mes + "-" + archivo.getName());
                  Logger.log("Se cambio el nombre del archivo a: " + year + "-" + mes + "-" + archivo.getName());
                  // let user = Session.getActiveUser().getEmail();
                  // Logger.log("Usuario que cargó reporte: "+user);
                  //Se mueve el archivo a la carpeta de destino
                  Logger.log("Se moverà el archivo a la carpeta de destino: ");
                  archivo.moveTo(DriveApp.getFolderById(Params.REPORTES_PROCESADOS.Folder_ID_FinalSheet));

                  //Se envia correo electrónico notificando el proceso
                  Logger.log("Se enviarà correl electrònico mediante la funciòn  enviarCorreoNotificacion(arg1,arg2)");
                  let tarjetasRegistradas = lastRow;
                  //Se pasa el nombre del trabajador, la cantidad de registros y el primer registro del archivo cargado.
                  enviarCorreoNotificacion(nombreTrabajador, tarjetasRegistradas, dataSheet[0][0]);

                } else {
                  Logger.log("Existen registros duplicados. Archivo será eliminado y se enviara correo a HSEQ");
                  archivo.setTrashed(true);
                  let correo = 'coordinadorhseq@ramdesolidscontrol.com';
                  let libro = SpreadsheetApp.openById(Params.REPORTE.SheetID);
                  let sheet = libro.getSheetByName(Params.REPORTE.NameSheetTrabajadores);
                  let dataTrabajadores = sheet.getDataRange().getValues();

                  const mensaje = response.body.mensajeDuplicado;

                  Logger.log("trabajador: " + nombreTrabajador + " con la siguiente información: " + mensaje);

                  for (let i = 0; i < dataTrabajadores.length; i++) {
                    if (nombreTrabajador == dataTrabajadores[i] || dataTrabajadores[i].includes(nombreTrabajador)) {
                      correo = dataTrabajadores[i][1];
                      Logger.log("Se enviará mensaje al correo: " + correo + " del trabajador: " + nombreTrabajador + " con la siguiente información: " + mensaje);

                      criticalCheckPoint(nombreTrabajador, correo, mensaje);
                      Logger.log("Se envio mensaje al correo: " + correo + " del trabajador: " + nombreTrabajador + " con la siguiente información: " + mensaje);
                      break;
                    }
                  }
                }

              } else {
                let mensajeError = 'Achivo tiene más columnas de las requeridas, verifique que no tenga celdas combinadas';
                Logger.log(mensajeError);
                archivo.setTrashed(true);

                let correo = 'coordinadorhseq@ramdesolidscontrol.com';
                let libro = SpreadsheetApp.openById(Params.REPORTE.SheetID);
                let sheet = libro.getSheetByName(Params.REPORTE.NameSheetTrabajadores);
                let dataTrabajadores = sheet.getDataRange().getValues();

                for (let i = 0; i < dataTrabajadores.length; i++) {
                  if (nombreTrabajador == dataTrabajadores[i] || dataTrabajadores[i].includes(nombreTrabajador)) {
                    correo = dataTrabajadores[i][1];
                    break;
                  }
                }
                criticalCheckPoint(nombreTrabajador, correo, mensajeError);
              }
            } else {
              Logger.log("El archivo no corresponde al reporte de tarjetas y será eliminado: No Exite Columna RIG");
              archivo.setTrashed(true);
              body.message = 'Archivo cargado no corresponde al formato de tarjetas de reporte: Columna RIG no se detecta';
              criticalCheckPoint('Erlin', 'coordinadorhseq@ramdesolidscontrol.com', body.message);
              return buildResponse(ResponseStatus.Unsuccessful, body)
            }
          } else {
            Logger.log("El archivo no corresponde al reporte de tarjetas y será eliminado: No Exite Columna POZO");
            archivo.setTrashed(true);
            body.message = 'Archivo cargado no corresponde al formato de tarjetas de reporte: Columna POZO no se detecta';
            criticalCheckPoint('Erlin', 'coordinadorhseq@ramdesolidscontrol.com', body.message);
            return buildResponse(ResponseStatus.Unsuccessful, body)
          }
        } else {
          Logger.log("El archivo no corresponde al reporte de tarjetas y será eliminado: No Exite Columna FECHA");
          archivo.setTrashed(true);
          body.message = 'Archivo cargado no corresponde al formato de tarjetas de reporte: Columna FECHA no se detecta';
          criticalCheckPoint('Erlin', 'coordinadorhseq@ramdesolidscontrol.com', body.message);
          return buildResponse(ResponseStatus.Successful, body)

        }
      }
    }

    // body.message = 'Proceso ejecutado correctamente. EXITOSO';
    // criticalCheckPoint('Erlin', 'coordinadorhseq@ramdesolidscontrol.com', body.message);
    // return buildResponse(ResponseStatus.Successful, body)

  } catch (error) {
    body.message = 'Error en la funcion principal de Reporte de tarjetas' + `${error.stack} `;
    Logger.log(`ERROR - ${body.message} `);
    criticalCheckPoint('Erlin', 'coordinadorhseq@ramdesolidscontrol.com', body.message);
    response = obtenerDataGoogleSheet(archivos);
    const archivo = response.body.archivoEncontrado;
    archivo.setTrashed(true);
    return buildResponse(ResponseStatus.Unsuccessful, body);
  }
}

//Esta función convierte los archivos exls y slxm en archivos de google sheet y los mueve de carpeta
const procesarReporteGoogleSheet = () => {

  const body = {}

  try {
    var excelFile = DriveApp.getFoldersByName(Params.REPORTES_PROCESAR.Folder_Name);
    var folder = excelFile.next();
    var contents = folder.getFiles();

    if (contents.hasNext()) {
      while (contents.hasNext()) {
        var doc = contents.next();
        var tipoarchivo = doc.getMimeType();
        Logger.log("Existe un archivo en la carpeta " + Params.REPORTES_PROCESAR.Folder_Name);
        var excelFiles = DriveApp.getFilesByName(doc);
        var fileId = excelFiles.next();
        var filin = fileId.getId();
        var folderId = Drive.Files.get(filin).parents[0].id;
        var blob = fileId.getBlob();
        var resource = {
          title: fileId.getName(),
          mimeType: MimeType.GOOGLE_SHEETS,
          parents: [{ id: Params.REPORTES_PROCESADOS.Folder_ID }],
        };
        Logger.log("Se validará tipo de archivo cargado");
        if (tipoarchivo == 'application/vnd.ms-excel' || tipoarchivo == 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet') {
          Drive.Files.insert(resource, blob);
          doc.setTrashed(true);
          Logger.log("Archivo contiene macros o no corresponde al tipo adecuado");
        } else if (tipoarchivo == 'application/vnd.google-apps.spreadsheet') {
          Logger.log("El archivo tiene el formato adecuado");
          doc.moveTo(DriveApp.getFolderById(Params.REPORTES_PROCESADOS.Folder_ID));
          body.message = "Archivo tiene formato adecuado";
          return buildResponse(ResponseStatus.Successful, body)
        }
        else {
          body.message = "Archivo: " + doc + "eliminado porque no cumple con las reglas definidas o no Existe";
          Logger.log("Archivo: " + doc + "eliminado porque no cumple con las reglas definidas o no Existe");
          doc.setTrashed(true);
          criticalCheckPoint('Erlin', 'coordinadorhseq@ramdesolidscontrol.com', body.message);
          return buildResponse(ResponseStatus.Unsuccessful, body)
        }
      }
    } else {
      body.message = 'No existe Archivo con tarjetas de observación para ser procesados';
      Logger.log(`INFO - ${body.message} `);
      return buildResponse(ResponseStatus.Successful, body)
    }

  } catch (error) {
    body.message = 'Se genero un error en al procesar los archivos en la carpeta' + Params.REPORTES_PROCESAR.Folder_Name;
    Logger.log(`Error - ${body.message}.Description: ${error.stack} `);
    return buildResponse(ResponseStatus.Unsuccessful, body);
  }
}

//Esta función obtiene los datos del reporte de tarjetas, ya transformado como documento de google sheet.
const obtenerDataGoogleSheet = (archivos) => {

  Logger.log("Se contruye objeto para la función obtenerDataGoogleSheet()");
  // Se inicializa cuerpo de respuesta
  const body = {
    dataSheetGoogle: "",
    dataSheetGoogleEncabezados: "",
    nombreTrabajador: "",
    archivoEncontrado: "",
  }

  try {
    Logger.log("Se validará si existe archivo en la ruta");
    if (archivos.hasNext()) {
      while (archivos.hasNext()) {
        var archivo = archivos.next();
        Logger.log('INFO - La carpeta contiene un archivo llamado: ' + archivo.getName());
        var idArchivo = archivo.getId();
        let libro = SpreadsheetApp.openById(idArchivo);
        var indexHoja = libro.getSheets()[0];
        var nameSheet = indexHoja.getName();
        let sheet = libro.getSheetByName(nameSheet);
        let lastColumn = sheet.getLastColumn();
        let lastRow = sheet.getLastRow();
        let dataSheet = [];
        let dataSheetConEmcabezado = [];

        Logger.log("Se obtendrá dataSheet y dataSheetConEmcabezado");
        dataSheet = sheet.getRange(Params.REPORTE.Fila_Inicio_Sheet, Params.REPORTE.Columna_Inicio_Sheet, lastRow - 9, lastColumn - 1).getValues();

        dataSheetConEmcabezado = sheet.getRange(Params.REPORTE.Fila_Inicio_Sheet_conEncabezado, Params.REPORTE.Columna_Inicio_Sheet, 3, lastColumn - 1).getValues();

        Logger.log('INFO - Ejecuciòn realizada con EXITO');
        body.dataSheetGoogle = dataSheet;
        body.dataSheetGoogleEncabezados = dataSheetConEmcabezado;
        body.nombreTrabajador = dataSheet[0][13];
        body.archivoEncontrado = archivo;
        body.message = 'Achivo Sheet e información obtenida';
        body.lastRow = lastRow - 9;
        return buildResponse(ResponseStatus.Successful, body);
      }
    } else {
      Logger.log('ERROR - Ejecución Falida porque no se encontrò archivo');
      body.message = 'No existe archivo de google Sheet para procesar';
      Logger.log(`INFO - ${body.message} `);
      return buildResponse(ResponseStatus.Successful, body)
    }
  } catch (error) {
    body.message = 'Se presentó un problema al momento de ejecutar la funcion obtenerDataGoogleSheet';
    Logger.log(`Error - ${body.message}.Description: ${error.stack} `);
    body.archivoEncontrado = "NO EXITOSO. Archivo no encontrado";
    return buildResponse(ResponseStatus.Unsuccessful, body);
  }
}

const consolidarInformacion = (newRows) => {

  const body = {}

  try {
    let libro = SpreadsheetApp.openById(Params.REPORTE.SheetID);
    let sheet = libro.getSheetByName(Params.REPORTE.NameSheet);
    var rows = newRows.length;
    var lastRowSheet = sheet.getLastRow();
    var columns = newRows[0].length;
    sheet.insertRowsAfter(lastRowSheet, rows);
    var rangeNewRows = sheet.getRange(lastRowSheet + 1, 2, rows, columns);
    rangeNewRows.setValues(newRows);
    Logger.log("Filas agregadas con éxito a proyecto app script");

  } catch (error) {
    body.message = 'Se genero un error al consolidar la información en la hoja de google sheet: Consolidado';
    Logger.log(`Error - ${body.message}.Description: ${error.stack} `);
    return buildResponse(ResponseStatus.Unsuccessful, body);
  }

}

const enviarCorreoNotificacion = (nombreCargadoPorTrabajador, registros, primerRegistroArchivoCargado) => {
  const body = {};

  try {
    let correo = 'coordinadorhseq@ramdesolidscontrol.com';
    let listDataTrabajadores = [];
    let listNombresCompletos = [];
    let nombreTrabajadorData = [];

    listNombresCompletos = nombreCargadoPorTrabajador.split(" ");
    let primerNombre = listNombresCompletos[0].toLowerCase();

    let libro = SpreadsheetApp.openById(Params.REPORTE.SheetID);
    let sheet = libro.getSheetByName(Params.REPORTE.NameSheetTrabajadores);

    listDataTrabajadores = sheet.getDataRange().getValues();
    Logger.log("Se enviará correo de notificación. Se obtuvo data de trabajadores - hoja trabajadores");

    for (let i = 0; i < listDataTrabajadores.length; i++) {
      nombreTrabajadorData = listDataTrabajadores[i][0].split(" ");
      let primerNombreData = nombreTrabajadorData[0].toLowerCase();

      Logger.log(primerNombre + "==" + primerNombreData + " || " + primerNombreData + "=include(" + primerNombre) + ")";
      if (primerNombre == primerNombreData || primerNombreData.includes(primerNombre)) {
        correo = listDataTrabajadores[i][1];
        Logger.log("Se encontró correo de trabajador de la lista");
        break;
      } else {
        Logger.log("NO se encontró correo de trabajador de la lista en iteración: " + i);
      }
    }
    Logger.log("Se obtuvo el correo del trabajador que subio el archivo: " + correo);

    verificarRegistro(primerNombre, primerRegistroArchivoCargado);

    tablaInformativa = getData(Params.REPORTE.SheetID, 4, 1, 1, false);

    novelty(nombreCargadoPorTrabajador, registros, correo, tablaInformativa);

  } catch (error) {
    body.message = 'Se generó un error en la función enviarCorreoNotificacion()';
    Logger.log(`Error - ${body.message}. Description: ${error.stack}`);
    return buildResponse(ResponseStatus.Unsuccessful, body);
  }
};


const validarRegistrosDuplicados = (registrosNuevos) => {

  // Se inicializa cuerpo de respuesta
  const body = {
  }

  try {

    let observacionNueva = "";

    let libro = SpreadsheetApp.openById(Params.REPORTE.SheetID);
    let sheet = libro.getSheetByName(Params.REPORTE.NameSheet);
    var lastColumn = sheet.getLastColumn();
    var lastRowSheet = sheet.getLastRow();
    let dataSheetActuales = [];
    dataSheetActuales = sheet.getRange(Params.REPORTE.Fila_Inicio_Sheet, Params.REPORTE.Columna_Inicio_Sheet, lastRowSheet - 9, lastColumn - 1).getValues();

    for (let nuevo = 0; nuevo < registrosNuevos.length; nuevo++) {

      let pozoNueva = registrosNuevos[nuevo][1];
      observacionNueva = registrosNuevos[nuevo][7];
      let recomendacionNueva = registrosNuevos[nuevo][8];
      let mejoraNueva = registrosNuevos[nuevo][9];
      let trabajadorNuevo = registrosNuevos[nuevo][13];

      for (let actuales = 0; actuales < dataSheetActuales.length; actuales++) {

        let pozoActual = dataSheetActuales[actuales][1];
        let observacionActual = dataSheetActuales[actuales][7];
        let recomendacionActual = dataSheetActuales[actuales][8];
        let nejoraActual = dataSheetActuales[actuales][9];
        let trabajadorActual = dataSheetActuales[actuales][13];

        if (pozoNueva == pozoActual) {
          if (observacionNueva == observacionActual) {
            Logger.log(observacionNueva + "==" + observacionActual)
            if (recomendacionNueva == recomendacionActual) {
              if (mejoraNueva == nejoraActual) {
                if (trabajadorNuevo == trabajadorActual) {
                  body.bln_Duplicado = "true";
                  Logger.log("Este registro está duplicado: " + observacionNueva);
                }

              }
            }
          }
        }
      }
    }

    if (body.bln_Duplicado == "true") {
      body.mensajeDuplicado = "NO EXITOSO. El registro con observación: " + observacionNueva + " ya se encuentra duplicado. Por esa razón el archivo no podrá ser cargado";
      return buildResponse(ResponseStatus.Unsuccessful, body);
    } else {
      body.bln_Duplicado = 'false';
      body.mensajeDuplicado = "EXITOSO";
      return buildResponse(ResponseStatus.Successful, body);
    }
  } catch (error) {
    body.message = 'Se presentó un problema al momento de verificar duplicados';
    Logger.log(`Error - ${body.message}.Description: ${error.stack} `);
    body.mensajeDuplicado = "NO EXITOSO";
    return buildResponse(ResponseStatus.Unsuccessful, body);
  }
}

//Obtiene data de cualquier google sheet 
const getData = (idHoja, hoja, filaInicial, columnaInicial, encabezado) => {
  let dataTable = [];
  let spreadsheet = SpreadsheetApp.openById(idHoja);
  const nameSheet = spreadsheet.getSheets()[hoja].getSheetName();
  let sheet = spreadsheet.getSheetByName(nameSheet);
  let lastColunm = sheet.getLastColumn();
  let lasRow = sheet.getLastRow();
  if (encabezado) {
    dataTable = sheet.getRange(filaInicial, columnaInicial, lasRow - filaInicial, lastColunm - columnaInicial + 1).getValues();
  } else {
    dataTable = sheet.getRange(filaInicial, columnaInicial, lasRow, lastColunm).getValues();
  }
  return dataTable;
}

const verificarRegistro = (primerNombre, primerRegistroArchivoCargado) => {
  Logger.log("Se iniciará proceso de verificación para el trabajador: " + primerNombre);
  let body = {
    message: "",
  };
  let fila;
  let columna;
  let celda;
  let nombreTra;
  let mesesColumnas = {
    1: "B",
    2: "C",
    3: "D",
    4: "E",
    5: "F",
    6: "G",
    7: "H",
    8: "I",
    9: "J",
    10: "K",
    11: "L",
    12: "M",
  };
  let mesesReportados = [];
  let fechaCargue = new Date(primerRegistroArchivoCargado);

  let mesReporte = fechaCargue.getMonth() + 1; // Utiliza getMonth() en lugar de fechaCargue.month()

  try {
    columna = mesesColumnas[mesReporte];
    Logger.log("Se obtendrá la data de la hoja de trabajadores SI / NO");
    mesesReportados = getData(Params.REPORTE.SheetID, 4, 1, 1, false);

    for (let i = 0; i < mesesReportados.length; i++) {
      nombreTra = mesesReportados[i][0];
      let nombreTraData = nombreTra.toLowerCase(); // Agrega "let" para declarar la variable
      Logger.log("SE VALIDARA: " + nombreTraData + "Includes(" + primerNombre + ")");
      if (nombreTraData.includes(primerNombre)) {
        fila = i + 1;
        break;
      } else {
        fila = i + 1;
      }
    }
    celda = columna + fila;
    Logger.log("Se obtubo la celda a registrar: " + celda);

    let spreadsheet = SpreadsheetApp.openById(Params.REPORTE.SheetID);
    let nameSheet = spreadsheet.getSheets()[4].getSheetName();
    let sheet = spreadsheet.getSheetByName(nameSheet);

    if (mesReporte > 12) {
      sheet.getRange(celda).setValue("NO");
    } else {
      sheet.getRange(celda).setValue("SI");
    }
    body.message = "Se ejecutó correctamente la función verificarRegistro()";
    return buildResponse(ResponseStatus.Successful, body);
  } catch (error) {
    body.message = "Se generó un error en la función verificarRegistro(). Description: " + error.stack; 
    criticalCheckPoint('Erlin', 'coordinadorhseq@ramdesolidscontrol.com', body.message);
    Logger.log(`Error - ${body.message}. Description: ${error.stack}`);
    return buildResponse(ResponseStatus.Unsuccessful, body);
  }
};

