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
      },
      {
        nombre: "Frases",
        columnas: ["Tipo", "Frase"],
        datos: [
          // Puntual
          ["Puntual", "¡Crack, llegaste a tiempo! 🚀"],
          ["Puntual", "¡Así se hace, puntualito como reloj suizo! ⏱️"],
          ["Puntual", "¡Buen inicio de jornada! 🌅"],
          ["Puntual", "¡Eres puntual como el sol! ☀️"],
          ["Puntual", "¡Excelente disciplina! 💼"],
          ["Puntual", "¡A tiempo y con actitud! 💪"],
          ["Puntual", "¡La puntualidad es tu superpoder! 🦸‍♂️"],
          ["Puntual", "¡Ya estás dejando huella desde temprano! 👣"],
          ["Puntual", "¡Hoy también ganaste al reloj! 🕒"],
          ["Puntual", "¡Impecable llegada, sigue así! ✨"],

          // Tarde
          ["Tarde", "Uy... ¿te ganó la almohada? 😴"],
          ["Tarde", "¡Llegaste, pero justito! 🕗"],
          ["Tarde", "¡No pasa nada! Mañana será mejor. 🌤️"],
          ["Tarde", "¡Vamos que sí se puede mejorar! 💥"],
          ["Tarde", "¡Despierta campeón, que ya es hora! 🛌⏰"],
          ["Tarde", "¡La próxima más temprano, tú puedes! 💡"],
          ["Tarde", "¡No pierdas tu ritmo! 🎵"],
          ["Tarde", "¡Llegaste tarde, pero llegaste! 😅"],
          ["Tarde", "¡Mañana rompemos el récord! 🏁"],
          ["Tarde", "¡Un nuevo intento cada día! 💫"],

          // Salida
          ["Salida", "¡Buen trabajo hoy! 🛠️"],
          ["Salida", "Hora de descansar, lo hiciste bien ✨"],
          ["Salida", "¡Día completado como un pro! ✔️"],
          ["Salida", "¡Desconéctate y disfruta! 🎉"],
          ["Salida", "¡Misión cumplida! 🕶️"],
          ["Salida", "¡Gracias por tu esfuerzo! 🙌"],
          ["Salida", "¡Otro día productivo en el bolsillo! 📈"],
          ["Salida", "¡Ahora sí, a recargar energías! 🔋"],
          ["Salida", "¡Gran jornada! Nos vemos mañana. 👋"],
          ["Salida", "¡Hora de apagar motores! 🧠💤"]
        ]
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

      if (config.nombre === "Horarios") {
        hoja.getRange("B2:E").setNumberFormat("hh:mm");
        hoja.getRange("F2:F").setNumberFormat("0");
      }
    });
  } catch (error) {
    Logger.log('Error en crearHojasYColumnas: ' + error);
    throw error;
  }
}


