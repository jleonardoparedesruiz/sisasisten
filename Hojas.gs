/**
 * Crea o actualiza las hojas y establece los encabezados y datos iniciales.
 * Utiliza operaciones en bloque para optimizar la escritura y añade validación de datos.
 */
function crearHojasYColumnas() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojas = [
      {
        nombre: "Usuarios",
        columnas: [
          "USUARIO",           // DNI
          "Nombre Completo",
          "Área",
          "Cargo",
          "Email",
          "Contraseña",
          "Nivel",
          "HorasExtras"
        ],
        datos: [
          ["74047479", "Julio Leonardo Paredes Ruiz", "administrativa", "supervisor", "jleonardoparedesruiz@gmail.com", "74047479", 1, 0],
          ["12345678", "prueba", "trabajador", "soldador", "asasd@gmail.com", "12345678", 0, 1]
        ]
      },
      {
        nombre: "BDregistros",
        columnas: [
          "DNI",
          "Nombre",
          "Fecha",
          "Hora",
          "Tipo",
          "Observaciones",
          "Ubicación",
          "Link Imagen"
        ]
      },
      {
        nombre: "Correos",
        columnas: ["Nombre", "Email"]
      },
      {
        nombre: "Bdsinhorario",
        columnas: ["Fecha", "Hora", "Observaciones"]
      },
      {
        nombre: "Horarios",
        columnas: [
          "DIA",
          "Hora ingreso",
          "Hora salida",
          "Hora refrigerio inicio",
          "Hora refrigerio salida",
          "Tolerancia en minutos"
        ],
        datos: [
          ["Lunes",     "08:00", "17:00", "12:00", "13:00", 15],
          ["Martes",    "08:00", "17:00", "12:00", "13:00", 15],
          ["Miercoles", "08:00", "17:00", "12:00", "13:00", 15],
          ["Jueves",    "08:00", "17:00", "12:00", "13:00", 15],
          ["Viernes",   "08:00", "17:00", "12:00", "13:00", 15],
          ["Sabado",    "08:00", "13:00", "00:00", "00:00", 15],
          ["Domingo",   "08:00", "13:00", "00:00", "00:00", 15]
        ]
      }
    ];
    
    hojas.forEach(config => {
      // Obtener o crear la hoja
      let hoja = ss.getSheetByName(config.nombre);
      if (!hoja) {
        hoja = ss.insertSheet(config.nombre);
      } else {
        hoja.clearContents();
      }
      
      const headers = config.columnas;
      // Escribir encabezados en una sola operación
      hoja.getRange(1, 1, 1, headers.length).setValues([headers]);
      
      if (config.datos && config.datos.length > 0) {
        // Filtrar filas que coincidan en longitud con los encabezados
        const validDatos = config.datos.filter(row => row.length === headers.length);
        if (validDatos.length > 0) {
          hoja.getRange(2, 1, validDatos.length, headers.length).setValues(validDatos);
        }
      }
    });
  } catch (error) {
    Logger.log('Error en crearHojasYColumnas: ' + error);
    throw error;
  }
}


