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
    const usuario = PropertiesService.getUserProperties().getProperty("usuarioActivo");
    return !!usuario;
  } catch (error) {
    Logger.log("Error en validarSesionActiva: " + error);
    return false;
  }
}

function validarLogin(usuario, contrasena) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("Usuarios");
    const datos = hoja.getDataRange().getValues();
    const userTrimmed = usuario.trim();
    const passTrimmed = contrasena.trim();
    for (let i = 1; i < datos.length; i++) {
      const dni = datos[i][0].toString().trim();
      const clave = datos[i][5].toString().trim();
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
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("Usuarios");
    const usuario = PropertiesService.getUserProperties().getProperty("usuarioActivo");
    if (!usuario) return "Usuario";
    const datos = hoja.getDataRange().getValues();
    const userTrimmed = usuario.trim();
    for (let i = 1; i < datos.length; i++) {
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
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("Usuarios");
    const usuario = PropertiesService.getUserProperties().getProperty("usuarioActivo");
    if (!usuario) return 0;
    const datos = hoja.getDataRange().getValues();
    const userTrimmed = usuario.trim();
    for (let i = 1; i < datos.length; i++) {
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
    const usuario = PropertiesService.getUserProperties().getProperty("usuarioActivo");
    if (!usuario) return [];
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("BDregistros");
    const datos = hoja.getDataRange().getValues();
    const timeZone = ss.getSpreadsheetTimeZone();
    const fechaHoy = Utilities.formatDate(new Date(), timeZone, "yyyy-MM-dd");
    const resultados = [];
    const userTrimmed = usuario.trim();
    for (let i = 1; i < datos.length; i++) {
      const dniFila = datos[i][0].toString().trim();
      let fechaFila = "";
      const valorFecha = datos[i][2];
      if (valorFecha instanceof Date && !isNaN(valorFecha)) {
        fechaFila = Utilities.formatDate(valorFecha, timeZone, "yyyy-MM-dd");
      } else {
        fechaFila = valorFecha.toString().trim();
      }
      let horaFila = "";
      const valorHora = datos[i][3];
      if (valorHora instanceof Date && !isNaN(valorHora)) {
        horaFila = Utilities.formatDate(valorHora, timeZone, "HH:mm:ss");
      } else {
        horaFila = valorHora.toString().trim();
      }
      const tipo = datos[i][4].toString().trim();
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
 * subirYRegistrarAsistencia: Sube la imagen, verifica la geolocalización y registra la asistencia.
 * Evita registros duplicados para el mismo tipo en el mismo día.
 */
function subirYRegistrarAsistencia(imagenBase64, ubicacion, tipoEvento) {
  try {
    const usuario = PropertiesService.getUserProperties().getProperty("usuarioActivo");
    if (!usuario) {
      return { mensaje: "Usuario no autenticado." };
    }
    if (!ubicacion || ubicacion === "No disponible" || ubicacion === "No soportado") {
      return { mensaje: "Registro geolocalizado obligatorio. Asegúrate de tener el GPS activado." };
    }
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojaUsuarios = ss.getSheetByName("Usuarios");
    const hojaRegistros = ss.getSheetByName("BDregistros");
    const datosUsuarios = hojaUsuarios.getDataRange().getValues();
    let nombre = "Desconocido";
    const userTrimmed = usuario.trim();
    for (let i = 1; i < datosUsuarios.length; i++) {
      if (datosUsuarios[i][0].toString().trim() === userTrimmed) {
        nombre = datosUsuarios[i][1];
        break;
      }
    }
    const timeZone = ss.getSpreadsheetTimeZone();
    const now = new Date();
    const fechaHoy = Utilities.formatDate(now, timeZone, "yyyy-MM-dd");
    const horaAhora = Utilities.formatDate(now, timeZone, "HH:mm:ss");
    // Evitar duplicados: si ya se registró este tipo hoy, no registrar de nuevo.
    const datosRegistros = hojaRegistros.getDataRange().getValues();
    for (let i = 1; i < datosRegistros.length; i++) {
      const dniFila = datosRegistros[i][0].toString().trim();
      let fechaFila = "";
      const valorFecha = datosRegistros[i][2];
      if (valorFecha instanceof Date && !isNaN(valorFecha)) {
        fechaFila = Utilities.formatDate(valorFecha, timeZone, "yyyy-MM-dd");
      } else {
        fechaFila = valorFecha.toString().trim();
      }
      const tipoFila = datosRegistros[i][4].toString().trim();
      if (dniFila === userTrimmed && fechaFila === fechaHoy && tipoFila === tipoEvento) {
        return { mensaje: `Ya has registrado ${tipoEvento} hoy.` };
      }
    }
    // Subir imagen a Drive
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
      linkImagen
    ]);
    return { mensaje: `Se registró ${tipoEvento} correctamente.`, evento: tipoEvento, fecha: fechaHoy, hora: horaAhora };
  } catch (error) {
    Logger.log("Error en subirYRegistrarAsistencia: " + error);
    return { mensaje: "Error al registrar asistencia." };
  }
}

/************** GESTIÓN DE USUARIOS **************/
function guardarUsuarioEnHoja(userObj) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojaUsuarios = ss.getSheetByName("Usuarios");
    const datos = hojaUsuarios.getDataRange().getValues();
    const userTrimmed = userObj.dni.toString().trim();
    let filaEncontrada = -1;
    for (let i = 1; i < datos.length; i++) {
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
      const filaHoja = filaEncontrada + 1;
      const nuevosValores = [
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
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("Usuarios");
    const datos = hoja.getDataRange().getValues();
    const lista = [];
    for (let i = 1; i < datos.length; i++) {
      const row = datos[i];
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
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("Usuarios");
    const datos = hoja.getDataRange().getValues();
    const dniTrimmed = dni.toString().trim();
    for (let i = 1; i < datos.length; i++) {
      const dniFila = datos[i][0].toString().trim();
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
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojaRegistros = ss.getSheetByName("BDregistros");
    const datos = hojaRegistros.getDataRange().getValues();
    let totalRegistros = 0, totalEntradas = 0, totalSalidas = 0;
    const registrosPorUsuario = {};
    for (let i = 1; i < datos.length; i++) {
      if (!datos[i][0]) continue;
      totalRegistros++;
      const tipo = datos[i][4].toString().trim();
      if (tipo === "Entrada") totalEntradas++;
      if (tipo === "Salida") totalSalidas++;
      const dni = datos[i][0].toString().trim();
      registrosPorUsuario[dni] = (registrosPorUsuario[dni] || 0) + 1;
    }
    return { totalRegistros, totalEntradas, totalSalidas, registrosPorUsuario };
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
/**
 * Verifica si existe una "Entrada" sin "Salida" para el usuario activo en el día de hoy.
 * Retorna true si ya hay una Entrada sin Salida, de lo contrario false.
 */
function verificarEntradaSinSalida() {
  try {
    const usuario = PropertiesService.getUserProperties().getProperty("usuarioActivo");
    if (!usuario) return false;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("BDregistros");
    const datos = hoja.getDataRange().getValues();
    const timeZone = ss.getSpreadsheetTimeZone();
    const fechaHoy = Utilities.formatDate(new Date(), timeZone, "yyyy-MM-dd");
    let entradaEncontrada = false;
    for (let i = 1; i < datos.length; i++) {
      const dniFila = datos[i][0].toString().trim();
      if (dniFila !== usuario.trim()) continue;
      let fechaFila = "";
      const valorFecha = datos[i][2];
      if (valorFecha instanceof Date && !isNaN(valorFecha)) {
        fechaFila = Utilities.formatDate(valorFecha, timeZone, "yyyy-MM-dd");
      } else {
        fechaFila = valorFecha.toString().trim();
      }
      if (fechaFila !== fechaHoy) continue;
      const tipoFila = datos[i][4].toString().trim();
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























