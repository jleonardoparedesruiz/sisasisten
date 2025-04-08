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
    return false;
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
    return false;
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
    return "Usuario";
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
    return 0;
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
    return [];
  }
}

/**
 * obtenerRegistrosUsuario: Retorna un arreglo de objetos con todos los registros
 * del usuario logueado (filtrados por fecha si se proveen fechaInicio y fechaFin).
 * Se incluyen: fecha, hora, tipo, nombre, observaciones, ubicación, lugar, linkImagen y id (si existe).
 */
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
    return [];
  }
}

/**
 * subirYRegistrarAsistencia: Sube la imagen, verifica la geolocalización (geoballa) y registra la asistencia.
 * Evita registros duplicados para el mismo tipo en el mismo día (para usuarios sin horas extras).
 */
function subirYRegistrarAsistencia(imagenBase64, ubicacion, tipoEvento) {
  try {
    const usuario = PropertiesService.getUserProperties().getProperty("usuarioActivo");
    if (!usuario) return { mensaje: "Usuario no autenticado." };

    if (!ubicacion || ubicacion === "No disponible" || ubicacion === "No soportado") {
      return { mensaje: "Registro geolocalizado obligatorio. Asegúrate de tener el GPS activado." };
    }

    const validacion = obtenerValidacionHorario(tipoEvento);
    if (!validacion.permitido || validacion.permitido === false) return { mensaje: validacion.mensaje };

    const lugar = verificarGeoballa(ubicacion);
    if (!lugar || !lugar.dentro) {
      return { mensaje: `Estás a ${Math.round(lugar.distancia)} metros de "${lugar.lugar}".\nRadio permitido: ${lugar.radio} m.\nNo puedes marcar asistencia.` };
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
    var blob = Utilities.newBlob(Utilities.base64Decode(imagenBase64), MimeType.JPEG, `${userTrimmed}_${fechaHoy}_${horaAhora}.jpg`);
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
      let dias = ["Domingo", "Lunes", "Martes", "Miercoles", "Viernes", "Sabado", "Domingo"];
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
    return { mensaje: "Error al registrar asistencia: " + error };
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
    throw error;
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
    return [];
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
    return false;
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
    return { totalRegistros: 0, totalEntradas: 0, totalSalidas: 0, registrosPorUsuario: {} };
  }
}

/************** CIERRE DE SESIÓN **************/
function cerrarSesion() {
  try {
    PropertiesService.getUserProperties().deleteProperty("usuarioActivo");
    return true;
  } catch (error) {
    Logger.log("Error en cerrarSesion: " + error);
    return false;
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
    return false;
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
    return { mensaje: "Error al guardar geoballa." };
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
    return { mensaje: "Error al eliminar geoballa." };
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
    return [];
  }
}

function validarHorario(tipoEvento) {
  try {
    const usuario = PropertiesService.getUserProperties().getProperty("usuarioActivo");
    if (!usuario) return { permitido: false, mensaje: "Usuario no autenticado." };
    return obtenerValidacionHorario(tipoEvento);
  } catch (e) {
    Logger.log("Error en validarHorario: " + e);
    return { permitido: false, mensaje: "Error al validar horario: " + e };
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
  const horaEsperada = parseInt(h) + (parseInt(m) || 0) / 60;
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
/**
 * obtenerNombrePorDNI: Busca en la hoja "Usuarios" y retorna el nombre asociado al DNI.
 */
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
    return "";
  }
}

/**
 * guardarRegistroManual: Guarda un registro manual en la hoja "BDregistros".
 * Se agrega un UUID en una nueva columna (posición 10) para identificar el registro.
 */
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
    return { mensaje: "Error al guardar registro manual: " + error };
  }
}

/**
 * eliminarRegistroManual: Elimina un registro manual de la hoja "BDregistros" buscando el ID en la columna 10.
 */
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
    return { mensaje: "Error al eliminar el registro: " + error };
  }
}

/************** GESTIÓN DE USUARIOS (Funciones ya existentes) **************/
// (guardarUsuarioEnHoja, obtenerListaUsuarios, eliminarUsuario) ya definidas

/************** ESTADÍSTICAS Y REPORTES (Funciones ya existentes) **************/
// (obtenerEstadisticas) ya definida

/************** CIERRE DE SESIÓN (Función ya existente) **************/
// (cerrarSesion) ya definida

/************** VALIDACIÓN AUTOMÁTICA DE ENTRADA/SALIDA (Funciones ya existentes) **************/
// (verificarEntradaSinSalida) ya definida

/************** FUNCIONES DE GEOBALLAS (Funciones ya existentes) **************/
// (verificarGeoballa, calcularDistancia, toRad, guardarGeoballa, eliminarGeoballa, obtenerGeoballas) ya definidas

/************** VALIDACIÓN DE HORARIOS Y FRASES MOTIVACIONALES (Funciones ya existentes) **************/
// (validarHorario, obtenerValidacionHorario, obtenerFraseMotivacional) ya definidas

// Nota: Puedes extender este archivo con funciones de edición manual de registros, por ejemplo:

/**
 * actualizarRegistroManual: Actualiza un registro manual en la hoja "BDregistros"
 * buscando el registro por su ID (almacenado en la columna 10) y actualizando
 * tipo (columna 5), hora (columna 4), observaciones (columna 6), lugar (columna 8) y ubicación (columna 7).
 */
function actualizarRegistroManual(registroEditado) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName("BDregistros");
    var data = hoja.getDataRange().getValues();
    var timeZone = ss.getSpreadsheetTimeZone();
    for (var i = 1; i < data.length; i++) {
      if (data[i].length >= 10 && data[i][9] && data[i][9].toString() === registroEditado.id) {
        // Actualizamos: hora (col 4), tipo (col 5), observaciones (col 6), ubicación (col 7), lugar (col 8)
        hoja.getRange(i+1, 4).setValue(registroEditado.hora);
        hoja.getRange(i+1, 5).setValue(registroEditado.tipo);
        hoja.getRange(i+1, 6).setValue(registroEditado.observaciones);
        hoja.getRange(i+1, 7).setValue(registroEditado.ubicacion);
        hoja.getRange(i+1, 8).setValue(registroEditado.lugar);
        return { mensaje: "Registro actualizado con éxito." };
      }
    }
    return { mensaje: "Registro no encontrado." };
  } catch (error) {
    Logger.log("Error en actualizarRegistroManual: " + error);
    return { mensaje: "Error al actualizar el registro: " + error };
  }
}
