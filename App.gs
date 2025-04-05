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
      var fechaFila = "";
      var valorFecha = datos[i][2];
      if (valorFecha instanceof Date && !isNaN(valorFecha)) {
        fechaFila = Utilities.formatDate(valorFecha, timeZone, "yyyy-MM-dd");
      } else {
        fechaFila = valorFecha.toString().trim();
      }
      var horaFila = "";
      var valorHora = datos[i][3];
      if (valorHora instanceof Date && !isNaN(valorHora)) {
        horaFila = Utilities.formatDate(valorHora, timeZone, "HH:mm:ss");
      } else {
        horaFila = valorHora.toString().trim();
      }
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
 * subirYRegistrarAsistencia: Sube la imagen, verifica la geolocalización (geoballa) y registra la asistencia.
 * Evita registros duplicados para el mismo tipo en el mismo día.
 */
function subirYRegistrarAsistencia(imagenBase64, ubicacion, tipoEvento) {
  try {
    const usuario = PropertiesService.getUserProperties().getProperty("usuarioActivo");
    if (!usuario) return { mensaje: "Usuario no autenticado." };

    if (!ubicacion || ubicacion === "No disponible" || ubicacion === "No soportado") {
      return { mensaje: "Registro geolocalizado obligatorio. Asegúrate de tener el GPS activado." };
    }

    const validacion = obtenerValidacionHorario(tipoEvento);
    if (!validacion.permitido) return { mensaje: validacion.mensaje };

    const lugar = verificarGeoballa(ubicacion);
    if (!lugar || !lugar.dentro) {
      return { mensaje: `Estás a ${Math.round(lugar.distancia)} metros de "${lugar.lugar}".\nRadio permitido: ${lugar.radio} m.\nNo puedes marcar asistencia.` };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojaUsuarios = ss.getSheetByName("Usuarios");
    const hojaRegistros = ss.getSheetByName("BDregistros");
    const datosUsuarios = hojaUsuarios.getDataRange().getValues();
    const timeZone = ss.getSpreadsheetTimeZone();
    const now = new Date();

    const fechaHoy = Utilities.formatDate(now, timeZone, "yyyy-MM-dd");
    const horaAhora = Utilities.formatDate(now, timeZone, "HH:mm:ss");

    let nombre = "Desconocido";
    const userTrimmed = usuario.trim();
    for (let i = 1; i < datosUsuarios.length; i++) {
      if (datosUsuarios[i][0].toString().trim() === userTrimmed) {
        nombre = datosUsuarios[i][1];
        break;
      }
    }

    const registros = hojaRegistros.getDataRange().getValues();
    for (let i = 1; i < registros.length; i++) {
      const fila = registros[i];
      const dniFila = fila[0].toString().trim();
      const tipoFila = fila[4].toString().trim();
      const fechaFila = fila[2] instanceof Date
        ? Utilities.formatDate(fila[2], timeZone, "yyyy-MM-dd")
        : fila[2];
      if (dniFila === userTrimmed && fechaFila === fechaHoy && tipoFila === tipoEvento) {
        return { mensaje: `Ya has registrado ${tipoEvento} hoy.` };
      }
    }

    const carpeta = DriveApp.getFolderById("1fhycG_U-hatF-VqPmxEhD4JEhl2MCgWv");
    const blob = Utilities.newBlob(Utilities.base64Decode(imagenBase64), MimeType.JPEG, `${userTrimmed}_${fechaHoy}_${horaAhora}.jpg`);
    const archivo = carpeta.createFile(blob);
    const linkImagen = archivo.getUrl();

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

    // 🎯 Elegir tipo de frase: puntual / tarde / salida
    let tipoFrase = "puntual";
    if (tipoEvento === "Salida") {
      tipoFrase = "salida";
    } else if (tipoEvento === "Entrada") {
      const hojaHorarios = ss.getSheetByName("Horarios");
      const horarios = hojaHorarios.getDataRange().getValues();
      const dias = ["Domingo", "Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado"];
      const diaSemana = dias[now.getDay()];
      for (let i = 1; i < horarios.length; i++) {
        if (horarios[i][0].toString().toLowerCase() === diaSemana.toLowerCase()) {
          const horaIngreso = horarios[i][1];
          const toleranciaMin = parseInt(horarios[i][5]) || 0;
          const horaEsperada = new Date(now);
          const [h, m] = horaIngreso.toString().split(":");
          horaEsperada.setHours(parseInt(h), parseInt(m), 0, 0);
          const limite = new Date(horaEsperada.getTime() + toleranciaMin * 60000);
          if (now > limite) {
            tipoFrase = "tarde";
          }
          break;
        }
      }
    }

    const frase = obtenerFraseMotivacional(tipoFrase);

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
    return { totalRegistros: totalRegistros, totalEntradas: totalEntradas, totalSalidas: totalSalidas, registrosPorUsuario: registrosPorUsuario };
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
      var fechaFila = "";
      var valorFecha = datos[i][2];
      if (valorFecha instanceof Date && !isNaN(valorFecha)) {
        fechaFila = Utilities.formatDate(valorFecha, timeZone, "yyyy-MM-dd");
      } else {
        fechaFila = valorFecha.toString().trim();
      }
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
/**
 * Verifica si la ubicación está dentro de una geoballa válida
 * Devuelve un objeto con: { dentro: true/false, lugar: "nombre", distancia: X, radio: Y }
 */
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

      Logger.log(`🛰️ Revisando ${lugar} - Distancia: ${distancia} m | Radio: ${radio} m`);

      if (distancia <= radio) {
        return {
          lugar: lugar,
          dentro: true,
          distancia: Math.round(distancia),
          radio: radio
        };
      }

      // Guardamos la más cercana (aunque esté fuera del radio)
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

    return geoballaMasCercana; // Si no estuvo dentro, igual devolvemos info
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

  const dias = ["Domingo", "Lunes", "Martes", "Miercoles", "Jueves", "Viernes", "Sabado"];
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

  // ✅ Normalizar hora si es Date
  if (horaPermitida instanceof Date) {
    horaPermitida = Utilities.formatDate(horaPermitida, timeZone, "HH:mm");
  }

  const horaActual = now.getHours() + now.getMinutes() / 60;
  const [h, m] = horaPermitida.toString().split(":");
  const horaEsperada = parseInt(h) + (parseInt(m) || 0) / 60;
  const horaLimite = horaEsperada + (toleranciaMin / 60);

  if (tipoEvento === "Entrada") {
    const margenAnticipado = 15 / 60; // 15 minutos antes
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
