/**
 * @enum
 * Parametros de ejecucion del apps script
 */
const Params = {

  REPORTE: {
    SheetID: '1m_GOWO12Kf0v-Ts84fGLEo2pa0Rc8_F2xgdPZ4apM-A',
    NameSheet: 'Consolidado',
    NameSheetTrabajadores: 'Trabajadores',
    Fila_Inicio_Sheet: '10',
    Columna_Inicio_Sheet: '2',
    Fila_Inicio_Sheet_conEncabezado: '8',
  },
  REPORTES_PROCESAR: {
    Folder_Name: 'Cargar Reporte tarjetas -  Excel  (File responses)',
    
  },
  REPORTES_PROCESADOS: {
    Folder_ID: '1DMyjFeU9-PsAo9Cd_OXMnJAnhn_0ceQyo8CstQHauVCnZ4T1ssPvr_p4eVHEGxscqRRH-71_',    
    Folder_Name:'Reporte Tarjetas (File responses)',
    Folder_ID_FinalSheet: '1lsezdGYLHU3Dq5vuo7vzk32T2l6cAsYU',
  },
  EMAIL: {
    Destinatarios: '',
    GerenteGeneral:'',
    Asunto: 'Reporte de Tarjetas Cargado con Exito',
    Gerente: '',
    Cargo: 'Gerente de Proyectos',
  }
}
