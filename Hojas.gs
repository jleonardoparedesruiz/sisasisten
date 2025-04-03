function crearHojasYColumnas() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojas = [
      {
        nombre: "Usuarios",
        columnas: [
          "USUARIO", "Nombre Completo", "Área", "Cargo",
          "Email", "Contraseña", "Nivel", "HorasExtras"
        ],
        datos: [
          ["74047479", "Julio Leonardo Paredes Ruiz", "administrativa", "supervisor", "jleonardoparedesruiz@gmail.com", "74047479", 1, 0],
          ["12345678", "prueba", "trabajador", "soldador", "asasd@gmail.com", "12345678", 0, 1]
        ]
      },
      {
        nombre: "BDregistros",
        columnas: ["DNI", "Nombre", "Fecha", "Hora", "Tipo", "Observaciones", "Ubicación", "Lugar", "Link Imagen"]
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
          "DIA", "Hora ingreso", "Hora salida",
          "Hora refrigerio inicio", "Hora refrigerio salida", "Tolerancia en minutos"
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
      },
      {
        nombre: "geoballa",
        columnas: ["Lugar", "Ubicacion", "Radio (m)"]
      }
    ];

    hojas.forEach(config => {
      let hoja = ss.getSheetByName(config.nombre);
      if (!hoja) {
        hoja = ss.insertSheet(config.nombre);
      } else {
        hoja.clearContents();
      }

      const headers = config.columnas;
      hoja.getRange(1, 1, 1, headers.length).setValues([headers]);

      if (config.datos && config.datos.length > 0) {
        const validDatos = config.datos.filter(row => row.length === headers.length);
        if (validDatos.length > 0) {
          hoja.getRange(2, 1, validDatos.length, headers.length).setValues(validDatos);
        }
      }

      // 👉 Aplicar formato especial si es la hoja de horarios
      if (config.nombre === "Horarios") {
        // Formato hora para columnas B a E (2 a 5)
        hoja.getRange("B2:E").setNumberFormat("hh:mm");
        // Formato número entero para columna F (6)
        hoja.getRange("F2:F").setNumberFormat("0");
      }
    });
  } catch (error) {
    Logger.log('Error en crearHojasYColumnas: ' + error);
    throw error;
  }
}

