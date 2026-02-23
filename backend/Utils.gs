function ss() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function hoja(nombre) {
  return ss().getSheetByName(nombre);
}

function ahora() {
  return new Date();
}

// ==========================================
// CONTROL DE ENTORNOS (PROD / TEST)
// ==========================================
const DB_PROD_ID = "1zCxn5Cvuvfs29Hbpp58W6VCvV6AczGMG1o7CkhS8d2E";
const DB_TEST_ID = "1N7ofFjp98-B3-QvTBNEUzeA7G0V6rMVQVPlEyZ_ZJT4";

const ENV_MODE_DEFAULT = "PROD";

// NUEVA FUNCIÓN: Verifica si el usuario es admin leyendo la pestaña PERMISOS
function esAdminEnPermisos(userEmail) {
  if (!userEmail) return false;

  const db = SpreadsheetApp.openById(DB_PROD_ID);
  const sPermisos = db.getSheetByName("PERMISOS");
  if (!sPermisos) return false;

  const data = sPermisos.getDataRange().getValues();
  const emailBuscado = String(userEmail).trim().toLowerCase();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === emailBuscado) {
      return data[i][8] === true || String(data[i][8]).toUpperCase() === 'TRUE';
    }
  }
  return false;
}

function getCurrentUserEmail_() {
  return String(Session.getActiveUser().getEmail() || "").trim().toLowerCase();
}

// Permite evaluar si se puede cambiar de entorno usando el correo del Frontend o el de la sesión
function canToggleEnvironment(userEmail) {
  return esAdminEnPermisos(userEmail || getCurrentUserEmail_());
}

function getEnvScopeKey_() {
  // Clave anónima por usuario/sesión para evitar interferencia entre operadores
  const userKey = Session.getTemporaryActiveUserKey();
  return userKey ? "ENV_MODE__" + userKey : "ENV_MODE__GLOBAL";
}

function getActiveDbId() {
  const env = getCurrentEnvironment();
  return env === "TEST" ? DB_TEST_ID : DB_PROD_ID;
}

// MODIFICADO: Ahora recibe el userEmail desde el Frontend
function switchEnvironment(mode, userEmail) {
  if (mode !== "PROD" && mode !== "TEST") return;

  const esAdmin = esAdminEnPermisos(userEmail);

  if (!esAdmin && mode === "TEST") {
    throw new Error("No autorizado: solo administradores pueden habilitar entorno TEST. Revisa tu hoja de PERMISOS.");
  }

  const scriptProps = PropertiesService.getScriptProperties();
  const safeMode = esAdmin ? mode : "PROD";
  scriptProps.setProperty(getEnvScopeKey_(), safeMode);
  PropertiesService.getUserProperties().setProperty("ENV_MODE", safeMode);
  return safeMode;
}

function getCurrentEnvironment() {
  const scriptProps = PropertiesService.getScriptProperties();
  const userProps = PropertiesService.getUserProperties();

  const resolved =
    scriptProps.getProperty(getEnvScopeKey_()) ||
    userProps.getProperty("ENV_MODE") ||
    scriptProps.getProperty("ENV_MODE") ||
    ENV_MODE_DEFAULT;

  // CORRECCIÓN: Usamos esAdminEnPermisos en lugar de la función antigua
  if (!esAdminEnPermisos(getCurrentUserEmail_())) return "PROD";

  return resolved;
}

function getEnvironmentDiagnostics() {
  const env = getCurrentEnvironment();
  const targetDbId = getActiveDbId();
  const db = SpreadsheetApp.openById(targetDbId);

  return {
    env: env,
    targetDbId: targetDbId,
    connectedDbId: db.getId(),
    connectedDbName: db.getName(),
    isIsolated: db.getId() === targetDbId,
    scopeKey: getEnvScopeKey_(),
    canToggle: canToggleEnvironment(),
  };
}

// ==========================================
// GESTIÓN DE SESIÓN POR CORREO Y PIN
// ==========================================
function procesarLoginEmail(email, pin) {
  // Siempre validamos contra Producción
  const db = SpreadsheetApp.openById(DB_PROD_ID);
  const sPermisos = db.getSheetByName("PERMISOS");
  
  if (!sPermisos) throw new Error("Falta la pestaña PERMISOS en la base de datos.");

  const data = sPermisos.getDataRange().getValues();
  const emailInput = String(email).trim().toLowerCase();
  
  for (let i = 1; i < data.length; i++) {
    let emailDb = String(data[i][0]).trim().toLowerCase();
    
    // Si encuentra el correo en la base de datos
    if (emailDb === emailInput && emailDb !== "") {
      let esAdmin = data[i][8] === true || String(data[i][8]).toUpperCase() === 'TRUE';
      let pinDb = String(data[i][1]).trim(); // Columna B (El PIN secreto)
      
      // 1. Si es admin y NO mandó PIN, el backend avisa que requiere la contraseña
      if (esAdmin && (!pin || String(pin).trim() === "")) {
         return { requiresPin: true };
      }
      
      // 2. Si es admin y mandó PIN, lo validamos
      if (esAdmin && String(pin).trim() !== pinDb) {
         return { success: false, error: "PIN incorrecto. Acceso de administrador denegado." };
      }
      
      // 3. Login Exitoso (Para Operadores normales, o Admins con PIN correcto)
      let env = PropertiesService.getUserProperties().getProperty('ENV_MODE') || 'PROD';
      if (!esAdmin) {
        env = 'PROD'; // Los operadores siempre van a Producción forzosamente
        PropertiesService.getUserProperties().setProperty('ENV_MODE', 'PROD');
      }

      return {
        success: true,
        nombre: emailDb.split("@")[0], // Usa la primera parte del correo como nombre
        esAdmin: esAdmin,
        entorno: env,
        permisos: {
          entradas: data[i][2] === true || String(data[i][2]).toUpperCase() === 'TRUE',
          salidas: data[i][3] === true || String(data[i][3]).toUpperCase() === 'TRUE',
          ubicaciones: data[i][4] === true || String(data[i][4]).toUpperCase() === 'TRUE',
          productos: data[i][5] === true || String(data[i][5]).toUpperCase() === 'TRUE',
          envios: data[i][6] === true || String(data[i][6]).toUpperCase() === 'TRUE',
          bajas: data[i][7] === true || String(data[i][7]).toUpperCase() === 'TRUE'
        }
      };
    }
  }
  
  return { success: false, error: "Este correo no está registrado en el sistema." };
}

// ==========================================
// MÓDULO DE ADMINISTRACIÓN DE USUARIOS
// ==========================================
function obtenerListaUsuarios() {
  const db = SpreadsheetApp.openById(DB_PROD_ID);
  const s = db.getSheetByName("PERMISOS");
  if (!s) return [];
  
  const data = s.getDataRange().getValues();
  let usuarios = [];
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) { // Si hay correo
      usuarios.push({
        correo: String(data[i][0]).trim(),
        pin: data[i][1],
        entradas: data[i][2] === true || String(data[i][2]).toUpperCase() === 'TRUE',
        salidas: data[i][3] === true || String(data[i][3]).toUpperCase() === 'TRUE',
        ubicaciones: data[i][4] === true || String(data[i][4]).toUpperCase() === 'TRUE',
        productos: data[i][5] === true || String(data[i][5]).toUpperCase() === 'TRUE',
        envios: data[i][6] === true || String(data[i][6]).toUpperCase() === 'TRUE',
        bajas: data[i][7] === true || String(data[i][7]).toUpperCase() === 'TRUE',
        esAdmin: data[i][8] === true || String(data[i][8]).toUpperCase() === 'TRUE'
      });
    }
  }
  return usuarios;
}

function guardarUsuario(u) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const db = SpreadsheetApp.openById(DB_PROD_ID);
    const s = db.getSheetByName("PERMISOS");
    const data = s.getDataRange().getValues();
    
    let fila = -1;
    const correoBusqueda = String(u.correo).trim().toLowerCase();
    
    // Buscamos si el usuario ya existe para actualizarlo
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim().toLowerCase() === correoBusqueda) {
        fila = i + 1;
        break;
      }
    }
    
    const rowData = [
      u.correo, u.pin || "", 
      u.entradas, u.salidas, u.ubicaciones, 
      u.productos, u.envios, u.bajas, u.esAdmin
    ];
    
    if (fila > 0) {
      s.getRange(fila, 1, 1, 9).setValues([rowData]); // Actualiza
    } else {
      s.appendRow(rowData); // Nuevo usuario
    }
    return { success: true };
  } catch(e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

function eliminarUsuario(correoAEliminar, correoActualAdmin) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    if (String(correoAEliminar).trim().toLowerCase() === String(correoActualAdmin).trim().toLowerCase()) {
       throw new Error("Sistema de seguridad: No puedes eliminar tu propio usuario administrador.");
    }

    const db = SpreadsheetApp.openById(DB_PROD_ID);
    const s = db.getSheetByName("PERMISOS");
    const data = s.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim().toLowerCase() === String(correoAEliminar).trim().toLowerCase()) {
        s.deleteRow(i + 1);
        return { success: true };
      }
    }
    throw new Error("Usuario no encontrado.");
  } catch(e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// BITÁCORA DE ACTIVIDAD (TRAZABILIDAD)
// ==========================================
function registrarEnBitacora(usuario, accion, detalle) {
  try {
    // Usamos la base de datos activa (Prod o Test)
    const db = SpreadsheetApp.openById(getActiveDbId());
    let sBitacora = db.getSheetByName("BITACORA_ACTIVIDAD");
    
    // Si la hoja está vacía, le pone encabezados automáticos
    if (sBitacora.getLastRow() === 0) {
      sBitacora.appendRow(["FECHA", "USUARIO", "ACCIÓN", "DETALLE"]);
      sBitacora.getRange("A1:D1").setFontWeight("bold").setBackground("#f3f3f3");
    }
    
    // Escribimos el movimiento
    sBitacora.appendRow([new Date(), usuario, accion, detalle]);
    return true;
  } catch(e) {
    console.error("Error en bitácora: " + e.message);
    return false;
  }
}