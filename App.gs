/**
 * doGet: Retorna el HTML del login para iniciar el sistema.
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile("Login")
    .setTitle("Sistema de Asistencia SOLINPA");
}

/************** CONTROL DE SESIÓN Y LOGIN **************/
function validarSesionActiva() {
  try {
    var usuario = PropertiesService.getUserProperties().getProperty("usuarioActivo");
    return !!usuario;
  } catch (error) {
    Logger.log("Error en validarSesionActiva: " + error);
    return { error: true, mensaje: "Error en validarSesionActiva: " + error.message };
  }
}

function validarLogin(usuario, contrasena) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName("Usuarios");
    var datos = hoja.getDataRange().getValues();
    var userTrimmed = usuario.trim();
    var passTrimmed = contrasena.trim();
    for (var i = 1; i < datos.length; i++) {
      var dni = datos[i][0].toString().trim();
      var clave = datos[i][5].toString().trim();
      if (dni === userTrimmed && clave === passTrimmed) {
        PropertiesService.getUserProperties().setProperty("usuarioActivo", userTrimmed);
        return true;
      }
    }
    return false;
  } catch (error) {
    Logger.log("Error en validarLogin: " + error);
    return { error: true, mensaje: "Error en validarLogin: " + error.message };
  }
}

function obtenerNombreUsuario() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName("Usuarios");
    var usuario = PropertiesService.getUserProperties().getProperty("usuarioActivo");
    if (!usuario) return "Usuario";
    var datos = hoja.getDataRange().getValues();
    var userTrimmed = usuario.trim();
    for (var i = 1; i < datos.length; i++) {
      if (datos[i][0].toString().trim() === userTrimmed) {
        return datos[i][1];
      }
    }
    return "Usuario";
  } catch (error) {
    Logger.log("Error en obtenerNombreUsuario: " + error);
    return { error: true, mensaje: "Error en obtenerNombreUsuario: " + error.message };
  }
}

function obtenerNivelUsuario() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName("Usuarios");
    var usuario = PropertiesService.getUserProperties().getProperty("usuarioActivo");
    if (!usuario) return 0;
    var datos = hoja.getDataRange().getValues();
    var userTrimmed = usuario.trim();
    for (var i = 1; i < datos.length; i++) {
      if (datos[i][0].toString().trim() === userTrimmed) {
        return Number(datos[i][6]);
      }
    }
    return 0;
  } catch (error) {
    Logger.log("Error en obtenerNivelUsuario: " + error);
    return { error: true, mensaje: "Error en obtenerNivelUsuario: " + error.message };
  }
}

/************** MARCACIONES Y REGISTRO DE ASISTENCIA **************/
function obtenerMarcacionesHoy() {
  try {
    var usuario = PropertiesService.getUserProperties().getProperty("usuarioActivo");
    if (!usuario) return [];
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName("BDregistros");
    var datos = hoja.getDataRange().getValues();
    var timeZone = ss.getSpreadsheetTimeZone();
    var fechaHoy = Utilities.formatDate(new Date(), timeZone, "yyyy-MM-dd");
    var resultados = [];
    var userTrimmed = usuario.trim();
    
    for (var i = 1; i < datos.length; i++) {
      var dniFila = datos[i][0].toString().trim();
      var valorFecha = datos[i][2];
      var fechaFila = (valorFecha instanceof Date && !isNaN(valorFecha))
            ? Utilities.formatDate(valorFecha, timeZone, "yyyy-MM-dd")
            : valorFecha.toString().trim();
      var valorHora = datos[i][3];
      var horaFila = (valorHora instanceof Date && !isNaN(valorHora))
            ? Utilities.formatDate(valorHora, timeZone, "HH:mm:ss")
            : valorHora.toString().trim();
      var tipo = datos[i][4].toString().trim();

      if (dniFila === userTrimmed && fechaFila === fechaHoy) {
        resultados.push({ tipo: tipo, fecha: fechaFila, hora: horaFila });
      }
    }
    return resultados;
  } catch (error) {
    Logger.log("Error en obtenerMarcacionesHoy: " + error);
    return { error: true, mensaje: "Error en obtenerMarcacionesHoy: " + error.message };
  }
}

function obtenerRegistrosUsuario(fechaInicio, fechaFin) {
  try {
    var usuario = PropertiesService.getUserProperties().getProperty("usuarioActivo");
    if (!usuario) return [];
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName("BDregistros");
    var datos = hoja.getDataRange().getValues();
    var timeZone = ss.getSpreadsheetTimeZone();
    var userTrimmed = usuario.trim();
    var resultado = [];
    
    for (var i = 1; i < datos.length; i++) {
      var fila = datos[i];
      if (fila[0].toString().trim() !== userTrimmed) continue;
      
      var valorFecha = fila[2];
      var fechaStr = (valorFecha instanceof Date && !isNaN(valorFecha))
            ? Utilities.formatDate(valorFecha, timeZone, "yyyy-MM-dd")
            : valorFecha.toString().trim();
      
      if (fechaInicio && fechaFin) {
        if (fechaStr < fechaInicio || fechaStr > fechaFin) continue;
      }
      
      var valorHora = fila[3];
      var horaStr = (valorHora instanceof Date && !isNaN(valorHora))
            ? Utilities.formatDate(valorHora, timeZone, "HH:mm:ss")
            : valorHora.toString().trim();
      
      var tipo = fila[4] ? fila[4].toString().trim() : "";
      var nombre = fila[1] ? fila[1].toString() : "";
      var observaciones = fila[5] ? fila[5].toString() : "";
      var ubicacion = fila[6] ? fila[6].toString() : "";
      var lugar = fila[7] ? fila[7].toString() : "";
      var linkImagen = fila[8] ? fila[8].toString() : "";
      var id = (fila.length >= 10 && fila[9]) ? fila[9].toString() : "";
      
      resultado.push({
        fecha: fechaStr,
        hora: horaStr,
        tipo: tipo,
        nombre: nombre,
        observaciones: observaciones,
        ubicacion: ubicacion,
        lugar: lugar,
        foto: linkImagen,
        id: id
      });
    }
    resultado.sort(function(a, b) {
      var compFecha = a.fecha.localeCompare(b.fecha);
      if (compFecha !== 0) return compFecha;
      return a.hora.localeCompare(b.hora);
    });
    return resultado;
  } catch (error) {
    Logger.log("Error en obtenerRegistrosUsuario: " + error);
    return { error: true, mensaje: "Error en obtenerRegistrosUsuario: " + error.message };
  }
}

function obtenerReporteIndividual(dni, fechaInicio, fechaFin, tipo) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName("BDregistros");
    var datos = hoja.getDataRange().getValues();
    var timeZone = ss.getSpreadsheetTimeZone();
    var resultado = [];

    for (var i = 1; i < datos.length; i++) {
      var fila = datos[i];
      // Filtrar por DNI si se especifica
      var dniFila = fila[0].toString().trim();
      if (dni && dni.trim() !== "" && dniFila !== dni.trim()) continue;
      
      var valorFecha = fila[2];
      var fechaStr = (valorFecha instanceof Date && !isNaN(valorFecha))
            ? Utilities.formatDate(valorFecha, timeZone, "yyyy-MM-dd")
            : valorFecha.toString().trim();
      
      if (fechaInicio && fechaFin) {
        if (fechaStr < fechaInicio || fechaStr > fechaFin) continue;
      }
      
      var valorHora = fila[3];
      var horaStr = (valorHora instanceof Date && !isNaN(valorHora))
            ? Utilities.formatDate(valorHora, timeZone, "HH:mm:ss")
            : valorHora.toString().trim();
      
      var tipoFila = fila[4] ? fila[4].toString().trim() : "";
      if (tipo && tipo.trim() !== "" && tipoFila !== tipo.trim()) continue;
      
      var nombre = fila[1] ? fila[1].toString() : "";
      var observaciones = fila[5] ? fila[5].toString() : "";
      var ubicacion = fila[6] ? fila[6].toString() : "";
      var lugar = fila[7] ? fila[7].toString() : "";
      var linkImagen = fila[8] ? fila[8].toString() : "";
      var id = (fila.length >= 10 && fila[9]) ? fila[9].toString() : "";
      
      resultado.push({
        dni: dniFila,
        fecha: fechaStr,
        hora: horaStr,
        tipo: tipoFila,
        nombre: nombre,
        observaciones: observaciones,
        ubicacion: ubicacion,
        lugar: lugar,
        foto: linkImagen,
        id: id
      });
    }
    resultado.sort(function(a, b) {
      var compFecha = a.fecha.localeCompare(b.fecha);
      if (compFecha !== 0) return compFecha;
      return a.hora.localeCompare(b.hora);
    });
    return resultado;
  } catch (error) {
    Logger.log("Error en obtenerReporteIndividual: " + error);
    return { error: true, mensaje: "Error en obtenerReporteIndividual: " + error.message };
  }
}

/************** SUBIR Y REGISTRAR ASISTENCIA **************/
function subirYRegistrarAsistencia(imagenBase64, ubicacion, tipoEvento) {
  try {
    const usuario = PropertiesService.getUserProperties().getProperty("usuarioActivo");
    if (!usuario) {
      return { mensaje: "Usuario no autenticado." };
    }

    if (!ubicacion || ubicacion === "No disponible" || ubicacion === "No soportado") {
      return { mensaje: "Registro geolocalizado obligatorio. Asegúrate de tener el GPS activado." };
    }

    const validacion = obtenerValidacionHorario(tipoEvento);
    if (!validacion.permitido || validacion.permitido === false) {
      return { mensaje: validacion.mensaje };
    }

    const lugar = verificarGeoballa(ubicacion);
    if (!lugar) {
      return { mensaje: "No se encontró ninguna zona geográfica autorizada para marcar." };
    } else if (!lugar.dentro) {
      return { 
        mensaje: `Estás a ${Math.round(lugar.distancia)} metros de "${lugar.lugar}".\nRadio permitido: ${lugar.radio} m.\nNo puedes marcar asistencia.` 
      };
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hojaUsuarios = ss.getSheetByName("Usuarios");
    var hojaRegistros = ss.getSheetByName("BDregistros");
    var datosUsuarios = hojaUsuarios.getDataRange().getValues();
    var timeZone = ss.getSpreadsheetTimeZone();
    var now = new Date();
    var fechaHoy = Utilities.formatDate(now, timeZone, "yyyy-MM-dd");
    var horaAhora = Utilities.formatDate(now, timeZone, "HH:mm:ss");

    let nombre = "Desconocido";
    let horasExtrasActivas = 0;
    const userTrimmed = usuario.trim();
    for (let i = 1; i < datosUsuarios.length; i++) {
      if (datosUsuarios[i][0].toString().trim() === userTrimmed) {
        nombre = datosUsuarios[i][1];
        horasExtrasActivas = Number(datosUsuarios[i][7]) || 0;
        break;
      }
    }

    // Evitar duplicar registros si el usuario no tiene horas extras
    var registros = hojaRegistros.getDataRange().getValues();
    if (horasExtrasActivas === 0) {
      for (let i = 1; i < registros.length; i++) {
        let fila = registros[i];
        let dniFila = fila[0].toString().trim();
        let tipoFila = fila[4].toString().trim();
        let valorFecha = fila[2];
        let fechaFila = (valorFecha instanceof Date && !isNaN(valorFecha))
              ? Utilities.formatDate(valorFecha, timeZone, "yyyy-MM-dd")
              : valorFecha.toString().trim();

        if (dniFila === userTrimmed && fechaFila === fechaHoy && tipoFila === tipoEvento) {
          return { mensaje: `Ya has registrado ${tipoEvento} hoy.` };
        }
      }
    }

    // Subir la imagen y hacerla pública
    var carpeta = DriveApp.getFolderById("1fhycG_U-hatF-VqPmxEhD4JEhl2MCgWv");
    var blob = Utilities.newBlob(
      Utilities.base64Decode(imagenBase64),
      MimeType.JPEG,
      `${userTrimmed}_${fechaHoy}_${horaAhora}.jpg`
    );
    var archivo = carpeta.createFile(blob);
    archivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var linkImagen = archivo.getUrl();

    // Registrar asistencia
    hojaRegistros.appendRow([
      userTrimmed,
      nombre,
      fechaHoy,
      horaAhora,
      tipoEvento,
      "Ninguna",
      ubicacion,
      lugar.lugar,
      linkImagen
    ]);

    // Registrar horas extra si aplica
    if (horasExtrasActivas === 1 && tipoEvento === "Salida") {
      let ultimaEntrada = null;
      for (let i = 1; i < registros.length; i++) {
        let fila = registros[i];
        if (fila[0].toString().trim() === userTrimmed && fila[4].toString().trim() === "Entrada") {
          ultimaEntrada = fila;
        }
      }
      let horaEntrada = ultimaEntrada ? ultimaEntrada[3] : horaAhora;
      let entradaDate = new Date(fechaHoy + " " + horaEntrada);
      let salidaDate = new Date(fechaHoy + " " + horaAhora);
      let horasTrabajadas = (salidaDate - entradaDate) / (1000 * 60 * 60);

      let hojaHorarios = ss.getSheetByName("Horarios");
      let horarios = hojaHorarios.getDataRange().getValues();
      let dias = ["Domingo", "Lunes", "Martes", "Miercoles", "Jueves", "Viernes", "Sábado"];
      let diaSemana = dias[now.getDay()];
      let horaSalidaProgramada = null;
      for (let i = 1; i < horarios.length; i++) {
        if (horarios[i][0].toString().toLowerCase() === diaSemana.toLowerCase()) {
          horaSalidaProgramada = horarios[i][2];
          break;
        }
      }
      let horasExtra = 0;
      if (horaSalidaProgramada) {
        let salidaProgramadaDate = new Date(fechaHoy + " " + horaSalidaProgramada);
        if (salidaDate > salidaProgramadaDate) {
          horasExtra = (salidaDate - salidaProgramadaDate) / (1000 * 60 * 60);
        }
      }
      let hojaHorasExtra = ss.getSheetByName("HorasExtra");
      hojaHorasExtra.appendRow([
        userTrimmed,
        nombre,
        fechaHoy,
        horaEntrada,
        horaAhora,
        horasTrabajadas,
        horasExtra,
        ""
      ]);
    }

    // Determinar mensaje motivacional
    let tipoFrase = "puntual";
    if (tipoEvento === "Salida") {
      tipoFrase = "salida";
    } else if (tipoEvento === "Entrada") {
      let hojaHorarios = ss.getSheetByName("Horarios");
      let horarios = hojaHorarios.getDataRange().getValues();
      let dias = ["Domingo", "Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado"];
      let diaSemana = dias[now.getDay()];
      for (let i = 1; i < horarios.length; i++) {
        if (horarios[i][0].toString().toLowerCase() === diaSemana.toLowerCase()) {
          let horaIngreso = horarios[i][1];
          let toleranciaMin = parseInt(horarios[i][5]) || 0;
          let horaEsperada = new Date(now);
          let [h, m] = horaIngreso.toString().split(":");
          horaEsperada.setHours(parseInt(h), parseInt(m), 0, 0);
          let limite = new Date(horaEsperada.getTime() + toleranciaMin * 60000);
          if (now > limite) {
            tipoFrase = "tarde";
          }
          break;
        }
      }
    }
    let frase = obtenerFraseMotivacional(tipoFrase);
    return {
      mensaje: `✅ Se registró su ${tipoEvento.toLowerCase()} en: ${lugar.lugar}.\n${frase}`,
      evento: tipoEvento,
      fecha: fechaHoy,
      hora: horaAhora,
      lugar: lugar.lugar
    };
  } catch (error) {
    Logger.log("Error en subirYRegistrarAsistencia: " + error);
    return { error: true, mensaje: "Error al registrar asistencia: " + error.message };
  }
}

/************** GESTIÓN DE USUARIOS **************/
function guardarUsuarioEnHoja(userObj) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hojaUsuarios = ss.getSheetByName("Usuarios");
    var datos = hojaUsuarios.getDataRange().getValues();
    var userTrimmed = userObj.dni.toString().trim();
    var filaEncontrada = -1;

    for (var i = 1; i < datos.length; i++) {
      if (datos[i][0].toString().trim() === userTrimmed) {
        filaEncontrada = i;
        break;
      }
    }
    if (filaEncontrada === -1) {
      hojaUsuarios.appendRow([
        userObj.dni,
        userObj.nombre,
        userObj.area,
        userObj.cargo,
        userObj.email,
        userObj.contrasena,
        userObj.nivel,
        userObj.horasExtras
      ]);
    } else {
      var filaHoja = filaEncontrada + 1;
      var nuevosValores = [
        userObj.dni,
        userObj.nombre,
        userObj.area,
        userObj.cargo,
        userObj.email,
        userObj.contrasena,
        userObj.nivel,
        userObj.horasExtras
      ];
      hojaUsuarios.getRange(filaHoja, 1, 1, nuevosValores.length).setValues([nuevosValores]);
    }
  } catch (error) {
    Logger.log("Error en guardarUsuarioEnHoja: " + error);
    return { error: true, mensaje: "Error en guardarUsuarioEnHoja: " + error.message };
  }
}

function obtenerListaUsuarios() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName("Usuarios");
    var datos = hoja.getDataRange().getValues();
    var lista = [];

    for (var i = 1; i < datos.length; i++) {
      var row = datos[i];
      if (row[0]) {
        lista.push({
          dni: row[0],
          nombre: row[1],
          area: row[2],
          cargo: row[3],
          email: row[4],
          contrasena: row[5],
          nivel: row[6],
          horasExtras: row[7]
        });
      }
    }
    return lista;
  } catch (error) {
    Logger.log("Error en obtenerListaUsuarios: " + error);
    return { error: true, mensaje: "Error en obtenerListaUsuarios: " + error.message };
  }
}

function eliminarUsuario(dni) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName("Usuarios");
    var datos = hoja.getDataRange().getValues();
    var dniTrimmed = dni.toString().trim();

    for (var i = 1; i < datos.length; i++) {
      var dniFila = datos[i][0].toString().trim();
      if (dniFila === dniTrimmed) {
        hoja.deleteRow(i + 1);
        return true;
      }
    }
    return false;
  } catch (error) {
    Logger.log("Error en eliminarUsuario: " + error);
    return { error: true, mensaje: "Error en eliminarUsuario: " + error.message };
  }
}

/************** REPORTES Y ESTADÍSTICAS **************/
function obtenerEstadisticas() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hojaRegistros = ss.getSheetByName("BDregistros");
    var datos = hojaRegistros.getDataRange().getValues();
    var totalRegistros = 0, totalEntradas = 0, totalSalidas = 0;
    var registrosPorUsuario = {};

    for (var i = 1; i < datos.length; i++) {
      if (!datos[i][0]) continue;
      totalRegistros++;
      var tipo = datos[i][4].toString().trim();
      if (tipo === "Entrada") totalEntradas++;
      if (tipo === "Salida") totalSalidas++;
      var dni = datos[i][0].toString().trim();
      registrosPorUsuario[dni] = (registrosPorUsuario[dni] || 0) + 1;
    }
    return {
      totalRegistros: totalRegistros,
      totalEntradas: totalEntradas,
      totalSalidas: totalSalidas,
      registrosPorUsuario: registrosPorUsuario
    };
  } catch (error) {
    Logger.log("Error en obtenerEstadisticas: " + error);
    return { 
      error: true,
      mensaje: "Error en obtenerEstadisticas: " + error.message,
      totalRegistros: 0,
      totalEntradas: 0,
      totalSalidas: 0,
      registrosPorUsuario: {}
    };
  }
}

/************** CIERRE DE SESIÓN **************/
function cerrarSesion() {
  try {
    PropertiesService.getUserProperties().deleteProperty("usuarioActivo");
    return true;
  } catch (error) {
    Logger.log("Error en cerrarSesion: " + error);
    return { error: true, mensaje: "Error en cerrarSesion: " + error.message };
  }
}

/************** VALIDACIÓN AUTOMÁTICA DE ENTRADA/SALIDA **************/
function verificarEntradaSinSalida() {
  try {
    var usuario = PropertiesService.getUserProperties().getProperty("usuarioActivo");
    if (!usuario) return false;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName("BDregistros");
    var datos = hoja.getDataRange().getValues();
    var timeZone = ss.getSpreadsheetTimeZone();
    var fechaHoy = Utilities.formatDate(new Date(), timeZone, "yyyy-MM-dd");
    var entradaEncontrada = false;

    for (var i = 1; i < datos.length; i++) {
      var dniFila = datos[i][0].toString().trim();
      if (dniFila !== usuario.trim()) continue;
      var valorFecha = datos[i][2];
      var fechaFila = (valorFecha instanceof Date && !isNaN(valorFecha))
            ? Utilities.formatDate(valorFecha, timeZone, "yyyy-MM-dd")
            : valorFecha.toString().trim();

      if (fechaFila !== fechaHoy) continue;
      var tipoFila = datos[i][4].toString().trim();
      if (tipoFila === "Entrada") {
        entradaEncontrada = true;
      } else if (tipoFila === "Salida") {
        return false;
      }
    }
    return entradaEncontrada;
  } catch (error) {
    Logger.log("Error en verificarEntradaSinSalida: " + error);
    return { error: true, mensaje: "Error en verificarEntradaSinSalida: " + error.message };
  }
}

/************** FUNCIONES DE GEOBALLAS **************/
function verificarGeoballa(ubicacion) {
  try {
    var parts = ubicacion.split(",");
    if (parts.length < 2) return null;
    var latUser = parseFloat(parts[0].trim());
    var lngUser = parseFloat(parts[1].trim());
    if (isNaN(latUser) || isNaN(lngUser)) return null;

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var geoSheet = ss.getSheetByName("geoballa");
    if (!geoSheet) return null;

    var data = geoSheet.getDataRange().getValues();
    var geoballaMasCercana = null;
    var distanciaMinima = Infinity;

    for (var i = 1; i < data.length; i++) {
      var lugar = data[i][0];
      var ubicacionGeo = data[i][1];
      var radio = parseFloat(data[i][2]);
      if (!ubicacionGeo || isNaN(radio)) continue;
      var geoParts = ubicacionGeo.split(",");
      var latGeo = parseFloat(geoParts[0].trim());
      var lngGeo = parseFloat(geoParts[1].trim());
      if (isNaN(latGeo) || isNaN(lngGeo)) continue;
      var distancia = calcularDistancia(latUser, lngUser, latGeo, lngGeo);
      Logger.log("🛰️ Revisando " + lugar + " - Distancia: " + distancia + " m | Radio: " + radio + " m");

      if (distancia <= radio) {
        return {
          lugar: lugar,
          dentro: true,
          distancia: Math.round(distancia),
          radio: radio
        };
      }
      if (distancia < distanciaMinima) {
        distanciaMinima = distancia;
        geoballaMasCercana = {
          lugar: lugar,
          dentro: false,
          distancia: Math.round(distancia),
          radio: radio
        };
      }
    }
    return geoballaMasCercana;
  } catch (error) {
    Logger.log("⚠️ Error en verificarGeoballa: " + error);
    return null;
  }
}

function calcularDistancia(lat1, lng1, lat2, lng2) {
  var R = 6371000; // Radio de la Tierra en metros
  var dLat = toRad(lat2 - lat1);
  var dLng = toRad(lng2 - lng1);
  var a = Math.sin(dLat / 2) * Math.sin(dLat / 2) +
          Math.cos(toRad(lat1)) * Math.cos(toRad(lat2)) *
          Math.sin(dLng / 2) * Math.sin(dLng / 2);
  var c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
  return R * c;
}
function toRad(value) {
  return value * Math.PI / 180;
}

function guardarGeoballa(geoObj) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName("geoballa");
    if (!hoja) {
      return { mensaje: "La hoja 'geoballa' no existe." };
    }
    var data = hoja.getDataRange().getValues();
    var filaEncontrada = -1;
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().trim() === geoObj.lugar.trim()) {
        filaEncontrada = i;
        break;
      }
    }
    if (filaEncontrada === -1) {
      hoja.appendRow([geoObj.lugar, geoObj.ubicacion, geoObj.radio]);
    } else {
      hoja.getRange(filaEncontrada + 1, 1).setValue(geoObj.lugar);
      hoja.getRange(filaEncontrada + 1, 2).setValue(geoObj.ubicacion);
      hoja.getRange(filaEncontrada + 1, 3).setValue(geoObj.radio);
    }
    return { mensaje: "Geoballa guardada correctamente." };
  } catch (error) {
    Logger.log("Error en guardarGeoballa: " + error);
    return { error: true, mensaje: "Error en guardarGeoballa: " + error.message };
  }
}

function eliminarGeoballa(lugar) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName("geoballa");
    var data = hoja.getDataRange().getValues();
    var lugarTrimmed = lugar.toString().trim();

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().trim() === lugarTrimmed) {
        hoja.deleteRow(i + 1);
        return { mensaje: "Geoballa eliminada correctamente." };
      }
    }
    return { mensaje: "No se encontró la geoballa para el lugar especificado." };
  } catch (error) {
    Logger.log("Error en eliminarGeoballa: " + error);
    return { error: true, mensaje: "Error en eliminarGeoballa: " + error.message };
  }
}

function obtenerGeoballas() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName("geoballa");
    var data = hoja.getDataRange().getValues();
    var geoballas = [];

    for (var i = 1; i < data.length; i++) {
      if (data[i][0]) {
        geoballas.push({
          lugar: data[i][0],
          ubicacion: data[i][1],
          radio: data[i][2]
        });
      }
    }
    return geoballas;
  } catch (error) {
    Logger.log("Error en obtenerGeoballas: " + error);
    return { error: true, mensaje: "Error en obtenerGeoballas: " + error.message, geoballas: [] };
  }
}

function validarHorario(tipoEvento) {
  try {
    const usuario = PropertiesService.getUserProperties().getProperty("usuarioActivo");
    if (!usuario) {
      return { permitido: false, mensaje: "Usuario no autenticado." };
    }
    return obtenerValidacionHorario(tipoEvento);
  } catch (e) {
    Logger.log("Error en validarHorario: " + e);
    return { permitido: false, mensaje: "Error al validar horario: " + e.message };
  }
}

function obtenerValidacionHorario(tipoEvento) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaHorarios = ss.getSheetByName("Horarios");
  const horarios = hojaHorarios.getDataRange().getValues();
  const timeZone = ss.getSpreadsheetTimeZone();
  const now = new Date();
  const dias = ["Domingo", "Lunes", "Martes", "Miercoles", "Jueves", "Viernes", "Sábado"];
  const diaSemana = dias[now.getDay()];
  let horaPermitida = null;
  let toleranciaMin = 0;

  for (let i = 1; i < horarios.length; i++) {
    if (horarios[i][0].toString().toLowerCase() === diaSemana.toLowerCase()) {
      if (tipoEvento === "Entrada") {
        horaPermitida = horarios[i][1];
      } else if (tipoEvento === "Salida") {
        horaPermitida = horarios[i][2];
      }
      toleranciaMin = parseInt(horarios[i][5]) || 0;
      break;
    }
  }
  if (!horaPermitida) {
    return { permitido: false, mensaje: `No hay horario configurado para ${diaSemana}.` };
  }
  if (horaPermitida instanceof Date) {
    horaPermitida = Utilities.formatDate(horaPermitida, timeZone, "HH:mm");
  }
  const horaActual = now.getHours() + now.getMinutes() / 60;
  const [h, m] = horaPermitida.toString().split(":");
  const horaEsperada = parseInt(h) + ((parseInt(m) || 0) / 60);
  const horaLimite = horaEsperada + (toleranciaMin / 60);

  if (tipoEvento === "Entrada") {
    const margenAnticipado = 15 / 60;
    const horaMinima = horaEsperada - margenAnticipado;
    if (horaActual < horaMinima) {
      const minutosFaltantes = Math.round((horaMinima - horaActual) * 60);
      return {
        permitido: false,
        mensaje: `Aún no puedes marcar. Espera ${minutosFaltantes} minuto(s) más (desde las ${Utilities.formatDate(new Date(now.getTime() + minutosFaltantes * 60000), timeZone, "HH:mm")}).`
      };
    }
    if (horaActual > horaLimite) {
      const minutosTarde = Math.round((horaActual - horaEsperada) * 60);
      return {
        permitido: false,
        mensaje: `Llegaste con ${minutosTarde} minuto(s) de retraso respecto al horario permitido (${horaPermitida} + ${toleranciaMin} min).`
      };
    }
  }
  if (tipoEvento === "Salida" && horaActual < horaEsperada) {
    return {
      permitido: "confirm",
      mensaje: `Aún no es hora de salida (normalmente a las ${horaPermitida}). ¿Deseas marcar la salida antes de lo establecido?`
    };
  }
  return { permitido: true };
}

function obtenerFraseMotivacional(tipoFrase) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("Frases");
    if (!hoja) return "¡Buen trabajo!";
    const datos = hoja.getDataRange().getValues();
    const encabezado = datos[0];
    const idxFrase = encabezado.indexOf("Frase");
    const idxTipo = encabezado.indexOf("Tipo");
    if (idxFrase === -1 || idxTipo === -1) return "¡Buen trabajo!";

    const frasesFiltradas = datos.slice(1).filter(fila =>
      fila[idxFrase] && fila[idxTipo] &&
      fila[idxTipo].toString().toLowerCase() === tipoFrase.toLowerCase()
    ).map(fila => fila[idxFrase]);

    if (frasesFiltradas.length === 0) return "¡Buen trabajo!";
    const index = Math.floor(Math.random() * frasesFiltradas.length);
    return frasesFiltradas[index];
  } catch (e) {
    Logger.log("Error en obtenerFraseMotivacional: " + e);
    return "¡Buen trabajo!";
  }
}

/************** FUNCIONES NUEVAS PARA REGISTRO MANUAL **************/
function obtenerNombrePorDNI(dni) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName("Usuarios");
    var datos = hoja.getDataRange().getValues();
    for (var i = 1; i < datos.length; i++) {
      if (datos[i][0].toString().trim() === dni.trim()) {
        return datos[i][1];
      }
    }
    return "";
  } catch (error) {
    Logger.log("Error en obtenerNombrePorDNI: " + error);
    return { error: true, mensaje: "Error en obtenerNombrePorDNI: " + error.message };
  }
}

function guardarRegistroManual(registro) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName("BDregistros");
    var timeZone = ss.getSpreadsheetTimeZone();
    var fecha = registro.fecha || Utilities.formatDate(new Date(), timeZone, "yyyy-MM-dd");
    var regId = Utilities.getUuid();

    hoja.appendRow([
      registro.dni,
      registro.nombre,
      fecha,
      registro.hora,
      registro.tipo,
      registro.observaciones || "",
      registro.ubicacion,
      registro.lugar,
      "", // Link Imagen vacío
      regId  // Columna para el ID
    ]);
    return { mensaje: "Registro manual guardado correctamente." };
  } catch (error) {
    Logger.log("Error en guardarRegistroManual: " + error);
    return { error: true, mensaje: "Error al guardar registro manual: " + error.message };
  }
}

function eliminarRegistroManual(regId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName("BDregistros");
    var data = hoja.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i].length >= 10 && data[i][9] && data[i][9].toString() === regId) {
        hoja.deleteRow(i + 1);
        return { mensaje: "Registro eliminado correctamente." };
      }
    }
    return { mensaje: "Registro no encontrado." };
  } catch (error) {
    Logger.log("Error en eliminarRegistroManual: " + error);
    return { error: true, mensaje: "Error al eliminar el registro: " + error.message };
  }
}

function actualizarRegistroManual(registroEditado) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName("BDregistros");
    var data = hoja.getDataRange().getValues();
    var timeZone = ss.getSpreadsheetTimeZone();

    for (var i = 1; i < data.length; i++) {
      if (data[i].length >= 10 && data[i][9] && data[i][9].toString() === registroEditado.id) {
        hoja.getRange(i + 1, 4).setValue(registroEditado.hora);
        hoja.getRange(i + 1, 5).setValue(registroEditado.tipo);
        hoja.getRange(i + 1, 6).setValue(registroEditado.observaciones);
        hoja.getRange(i + 1, 7).setValue(registroEditado.ubicacion);
        hoja.getRange(i + 1, 8).setValue(registroEditado.lugar);
        return { mensaje: "Registro actualizado con éxito." };
      }
    }
    return { mensaje: "Registro no encontrado." };
  } catch (error) {
    Logger.log("Error en actualizarRegistroManual: " + error);
    return { error: true, mensaje: "Error al actualizar el registro: " + error.message };
  }
}

/************** REGISTRO DE FALTAS **************/
function registrarFaltasAutomaticas() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hojaUsuarios = ss.getSheetByName("Usuarios");
    var hojaRegistros = ss.getSheetByName("BDregistros");
    var hojaFaltas = ss.getSheetByName("Faltas");
    var usuarios = hojaUsuarios.getDataRange().getValues();
    var registros = hojaRegistros.getDataRange().getValues();
    var timeZone = ss.getSpreadsheetTimeZone();
    var fechaHoy = Utilities.formatDate(new Date(), timeZone, "yyyy-MM-dd");

    // Objeto para saber quién registró entrada hoy
    var marcadosHoy = {};
    for (var i = 1; i < registros.length; i++) {
      var dni = registros[i][0].toString().trim();
      var valorFecha = registros[i][2];
      var fechaRegistro = (valorFecha instanceof Date && !isNaN(valorFecha))
          ? Utilities.formatDate(valorFecha, timeZone, "yyyy-MM-dd")
          : valorFecha.toString().trim();
      var tipo = registros[i][4].toString().trim();

      if (fechaRegistro === fechaHoy && tipo === "Entrada") {
        marcadosHoy[dni] = true;
      }
    }

    // Recorrer Usuarios y registrar faltas para quienes no marcaron entrada
    for (var i = 1; i < usuarios.length; i++) {
      var dniUsuario = usuarios[i][0].toString().trim();
      var nombreUsuario = usuarios[i][1] ? usuarios[i][1].toString().trim() : "";

      if (!marcadosHoy[dniUsuario]) {
        var datosFaltas = hojaFaltas.getDataRange().getValues();
        var existeFalta = false;
        for (var j = 1; j < datosFaltas.length; j++) {
          var dniFalta = datosFaltas[j][0].toString().trim();
          var fechaFalta = datosFaltas[j][2].toString().trim();
          if (dniUsuario === dniFalta && fechaFalta === fechaHoy) {
            existeFalta = true;
            break;
          }
        }
        if (!existeFalta) {
          hojaFaltas.appendRow([dniUsuario, nombreUsuario, fechaHoy, ""]);
        }
      }
    }
    return { mensaje: "Faltas automáticas registradas para " + fechaHoy };
  } catch (error) {
    Logger.log("Error en registrarFaltasAutomaticas: " + error);
    return { error: true, mensaje: "Error al registrar faltas: " + error.message };
  }
}

function obtenerFaltas() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName("Faltas");
    var data = hoja.getDataRange().getValues();
    var faltas = [];

    for (var i = 1; i < data.length; i++) {
      if (data[i][0]) {
        faltas.push({
          dni: data[i][0],
          nombre: data[i][1],
          fecha: data[i][2],
          observaciones: data[i][3] || "",
          id: i
        });
      }
    }
    return faltas;
  } catch (error) {
    Logger.log("Error en obtenerFaltas: " + error);
    return { error: true, mensaje: "Error en obtenerFaltas: " + error.message, faltas: [] };
  }
}

function actualizarFalta(faltaId, observaciones) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName("Faltas");
    var row = parseInt(faltaId) + 1;
    hoja.getRange(row, 4).setValue(observaciones);
    return { mensaje: "Falta actualizada correctamente." };
  } catch (error) {
    Logger.log("Error en actualizarFalta: " + error);
    return { error: true, mensaje: "Error al actualizar la falta: " + error.message };
  }
}

// NUEVA FUNCIÓN: OBTENER FALTAS POR USUARIO
function obtenerFaltasPorUsuario(dni) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName("Faltas");
    var data = hoja.getDataRange().getValues();
    var faltas = [];

    var usuarioDNI = dni || PropertiesService.getUserProperties().getProperty("usuarioActivo");
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().trim() === usuarioDNI.toString().trim()) {
        faltas.push({
          dni: data[i][0],
          nombre: data[i][1],
          fecha: data[i][2],
          observaciones: data[i][3] || "",
          id: i
        });
      }
    }
    return faltas;
  } catch (error) {
    Logger.log("Error en obtenerFaltasPorUsuario: " + error);
    return { error: true, mensaje: "Error en obtenerFaltasPorUsuario: " + error.message };
  }
}

/**
 * Función de prueba para ver en Logs las faltas del usuario activo.
 */
function probarFaltas() {
  var faltas = obtenerFaltasPorUsuario();
  Logger.log(JSON.stringify(faltas, null, 2));
}

/************** NUEVA FUNCIÓN PARA FILTRAR FALTAS POR FECHAS (Panel Admin) **************/
function obtenerFaltasPorFechas(fechaInicio, fechaFin) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName("Faltas");
    var datos = hoja.getDataRange().getValues();
    var faltas = [];
    var timeZone = ss.getSpreadsheetTimeZone();
    for (var i = 1; i < datos.length; i++) {
      var fecha = datos[i][2];
      var fechaStr;
      if (fecha instanceof Date && !isNaN(fecha)) {
        fechaStr = Utilities.formatDate(fecha, timeZone, "yyyy-MM-dd");
      } else {
        fechaStr = fecha.toString().trim();
      }
      // Filtrar por fechas si se han especificado ambas
      if (fechaInicio && fechaFin) {
        if (fechaStr < fechaInicio || fechaStr > fechaFin) continue;
      }
      if (datos[i][0]) {
        faltas.push({
          dni: datos[i][0],
          nombre: datos[i][1],
          fecha: fechaStr,
          observaciones: datos[i][3] || "",
          id: i
        });
      }
    }
    return faltas;
  } catch (error) {
    Logger.log("Error en obtenerFaltasPorFechas: " + error);
    return { error: true, mensaje: "Error en obtenerFaltasPorFechas: " + error.message, faltas: [] };
  }
}
