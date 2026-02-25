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

function canToggleEnvironment(userEmail) {
  return esAdminEnPermisos(userEmail);
}

// Genera una llave única temporal y anónima por cada navegador que abre la app
function getEnvScopeKey_() {
  const userKey = Session.getTemporaryActiveUserKey();
  return userKey ? "ENV_MODE__" + userKey : "ENV_MODE__GLOBAL";
}

function getCurrentEnvironment() {
  // SOLO lee la clave temporal del navegador actual, evitando cruces de usuarios
  const scriptProps = PropertiesService.getScriptProperties();
  return scriptProps.getProperty(getEnvScopeKey_()) || "PROD";
}

function getActiveDbId() {
  const env = getCurrentEnvironment();
  return env === "TEST" ? DB_TEST_ID : DB_PROD_ID;
}

function switchEnvironment(mode, userEmail) {
  if (mode !== "PROD" && mode !== "TEST") return;

  const esAdmin = esAdminEnPermisos(userEmail);

  if (!esAdmin && mode === "TEST") {
    throw new Error("No autorizado: solo administradores pueden habilitar entorno TEST.");
  }

  const safeMode = esAdmin ? mode : "PROD";
  
  // Guarda el entorno SOLO para la sesión actual del navegador
  PropertiesService.getScriptProperties().setProperty(getEnvScopeKey_(), safeMode);
  return safeMode;
}


// ==========================================
// GESTIÓN DE SESIÓN POR CORREO Y PIN
// ==========================================
function procesarLoginEmail(email, pin) {
  const db = SpreadsheetApp.openById(DB_PROD_ID);
  const sPermisos = db.getSheetByName("PERMISOS");
  if (!sPermisos) throw new Error("Falta la pestaña PERMISOS en la base de datos.");
  
  const data = sPermisos.getDataRange().getValues();
  const emailInput = String(email).trim().toLowerCase();
  
  for (let i = 1; i < data.length; i++) {
    let emailDb = String(data[i][0]).trim().toLowerCase();
    if (emailDb === emailInput && emailDb !== "") {
      let esAdmin = data[i][8] === true || String(data[i][8]).toUpperCase() === 'TRUE';
      let pinDb = String(data[i][1]).trim();
      
      if (esAdmin && (!pin || String(pin).trim() === "")) return { requiresPin: true };
      if (esAdmin && String(pin).trim() !== pinDb) return { success: false, error: "PIN incorrecto." };
      
      PropertiesService.getScriptProperties().setProperty(getEnvScopeKey_(), "PROD");

      // --- REGISTRO SILENCIOSO (NUNCA BLOQUEA) ---
      const scriptProps = PropertiesService.getScriptProperties();
      let rawSessions = scriptProps.getProperty('ACTIVE_SESSIONS');
      let sessions = rawSessions ? JSON.parse(rawSessions) : {};
      let now = new Date().getTime();

      // Limpiar fantasmas
      for (let user in sessions) {
         if (now - sessions[user].lastPing > 1800000) delete sessions[user];
      }

      sessions[emailDb] = { lastPing: now, rol: esAdmin ? "admin" : "user", env: "PROD" };
      scriptProps.setProperty('ACTIVE_SESSIONS', JSON.stringify(sessions));

      return {
        success: true,
        nombre: emailDb.split("@")[0], 
        esAdmin: esAdmin,
        entorno: "PROD",
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
  return { success: false, notFound: true, error: "Correo no registrado." };
}

function pingSesion(email, rol, entornoActual, vistaActual) {
  if (!email) return { sigueActivo: false };
  
  const scriptProps = PropertiesService.getScriptProperties();
  let raw = scriptProps.getProperty('ACTIVE_SESSIONS');
  if (!raw) return { sigueActivo: false }; 
  
  let sessions = JSON.parse(raw);
  let now = new Date().getTime();
  let correoNormalizado = String(email).trim().toLowerCase();
  
  // Rescatar la hora original de entrada para no perder su turno en la fila.
  let tiempoEntrada = (sessions[correoNormalizado] && sessions[correoNormalizado].loginTime) 
                      ? sessions[correoNormalizado].loginTime 
                      : now;
  
  sessions[correoNormalizado] = { 
     loginTime: tiempoEntrada, 
     lastPing: now, 
     rol: rol || "user", 
     env: entornoActual || "PROD", 
     vista: vistaActual || "desconocida" 
  };
  
  let adminsEnProd = [];
  let normalesEnProd = [];
  let usuariosConectados = [];
  
  // Limpiamos inactivos (25 seg)
  for (let user in sessions) {
     if (now - sessions[user].lastPing < 25000) { 
        usuariosConectados.push({ email: user, ...sessions[user] });
        if (sessions[user].env === "PROD") {
           if (sessions[user].rol === "admin") adminsEnProd.push(user);
           else normalesEnProd.push({ e: user, t: sessions[user].loginTime });
        }
     } else {
        delete sessions[user]; 
     }
  }
  scriptProps.setProperty('ACTIVE_SESSIONS', JSON.stringify(sessions));
  
  let tengoPermiso = false;
  let lider = "";

  if (entornoActual === "TEST") {
      tengoPermiso = true;
      lider = correoNormalizado;
  } else {
      if (rol === "admin") {
          tengoPermiso = true; // Admins en PROD SIEMPRE mandan
          lider = correoNormalizado; 
      } else {
          if (adminsEnProd.length > 0) {
              tengoPermiso = false;
              lider = adminsEnProd[0]; 
          } else {
              // Ordenamos por hora de llegada (el primero que entró gana)
              normalesEnProd.sort((a,b) => a.t - b.t);
              if (normalesEnProd.length > 0) {
                  lider = normalesEnProd[0].e;
                  tengoPermiso = (lider === correoNormalizado);
              }
          }
      }
  }
  
  return {
     sigueActivo: true,
     tengoPermiso: tengoPermiso,
     escritorActual: lider.split('@')[0], 
     listaUsuarios: usuariosConectados
  };
}

// NUEVO: Función para registrar un usuario sin permisos operativos
function registrarUsuarioPendiente(email) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const db = SpreadsheetApp.openById(DB_PROD_ID);
    const s = db.getSheetByName("PERMISOS");
    const data = s.getDataRange().getValues();
    
    const emailBuscado = String(email).trim().toLowerCase();
    
    // Verificamos por seguridad que no exista ya
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim().toLowerCase() === emailBuscado) {
        return { success: false, error: "El usuario ya existe." };
      }
    }
    
    // Agregamos a la hoja: [CORREO, PIN(vacío), Entradas, Salidas, Ubic, Prod, Envios, Bajas, EsAdmin] (Todo en false)
    s.appendRow([emailBuscado, "", false, false, false, false, false, false, false]); 
    
    return { success: true };
  } catch(e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
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
    if (data[i][0]) { 
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
      s.getRange(fila, 1, 1, 9).setValues([rowData]); 
    } else {
      s.appendRow(rowData); 
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
    // Busca la base de datos activa (Prod o Test)
    const db = SpreadsheetApp.openById(getActiveDbId());
    let sBitacora = db.getSheetByName("BITACORA_ACTIVIDAD");
    
    // MEJORA: Si la hoja NO existe en este Excel, ¡la crea automáticamente!
    if (!sBitacora) {
      sBitacora = db.insertSheet("BITACORA_ACTIVIDAD");
    }
    
    // Si la hoja está vacía (recién creada), le pone encabezados automáticos
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

function registrarCierreSesion(email) {
  if(!email) return;
  try {
    const scriptProps = PropertiesService.getScriptProperties();
    let raw = scriptProps.getProperty('ACTIVE_SESSIONS');
    if(raw) {
        let sessions = JSON.parse(raw);
        delete sessions[email.toLowerCase()];
        scriptProps.setProperty('ACTIVE_SESSIONS', JSON.stringify(sessions));
    }
  } catch(e) {}
}

// --- BOTÓN DE PÁNICO: CORRER ESTA FUNCIÓN DESDE EL EDITOR PARA DESTRABAR EL SISTEMA ---
function LIBERAR_SISTEMA() {
   PropertiesService.getScriptProperties().deleteProperty('ACTIVE_SESSIONS');
   console.log("✅ Sistema liberado exitosamente. Todos los usuarios fueron desconectados.");
}

// --- NUEVO: HEARTBEAT (LATIDO) PARA MANTENER Y VERIFICAR LA SESIÓN ---
function pingSesion(email, rol, entornoActual, vistaActual) {
  if (!email) return { sigueActivo: false };

  // --- SEMÁFORO DE CONCURRENCIA ---
  // Evita que dos navegadores se borren mutuamente si hacen ping en el mismo milisegundo
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(3000); 
  } catch (e) {
    return { sigueActivo: true, ignorarPing: true }; // Si hay tráfico, espera al siguiente latido
  }

  try {
      const scriptProps = PropertiesService.getScriptProperties();
      let raw = scriptProps.getProperty('ACTIVE_SESSIONS');
      if (!raw) return { sigueActivo: false }; 

      let sessions = JSON.parse(raw);
      let now = new Date().getTime();
      let correoNormalizado = String(email).trim().toLowerCase();

      // MANTENER ANTIGÜEDAD INTACTA
      let tiempoEntrada = (sessions[correoNormalizado] && sessions[correoNormalizado].loginTime) 
                          ? sessions[correoNormalizado].loginTime 
                          : now;

      sessions[correoNormalizado] = { 
         loginTime: tiempoEntrada, 
         lastPing: now, 
         rol: rol || "user", 
         env: entornoActual || "PROD", 
         vista: vistaActual || "desconocida" 
      };

      let adminsEnProd = [];
      let normalesEnProd = [];
      let usuariosConectados = [];

      // LIMPIAR INACTIVOS (Menos de 25 segundos para que suelte rápido el permiso al salir)
      for (let user in sessions) {
         if (now - sessions[user].lastPing < 25000) { 
            usuariosConectados.push({ email: user, ...sessions[user] });
            if (sessions[user].env === "PROD") {
               if (sessions[user].rol === "admin") adminsEnProd.push(user);
               else normalesEnProd.push({ e: user, t: sessions[user].loginTime });
            }
         } else {
            delete sessions[user]; 
         }
      }
      
      scriptProps.setProperty('ACTIVE_SESSIONS', JSON.stringify(sessions));

      // LÓGICA DE HERENCIA DE PERMISOS
      let tengoPermiso = false;
      let lider = "";

      if (entornoActual === "TEST") {
          tengoPermiso = true;
          lider = correoNormalizado;
      } else {
          if (rol === "admin") {
              tengoPermiso = true; // Admin en PROD SIEMPRE manda
              lider = correoNormalizado; 
          } else {
              if (adminsEnProd.length > 0) {
                  tengoPermiso = false;
                  lider = adminsEnProd[0]; 
              } else {
                  // ORDENAMOS POR HORA DE LLEGADA (El más antiguo toma el control)
                  normalesEnProd.sort((a, b) => a.t - b.t);
                  if (normalesEnProd.length > 0) {
                      lider = normalesEnProd[0].e;
                      tengoPermiso = (lider === correoNormalizado);
                  }
              }
          }
      }

      return {
         sigueActivo: true,
         tengoPermiso: tengoPermiso,
         escritorActual: lider.split('@')[0], 
         listaUsuarios: usuariosConectados
      };
      
  } finally {
      lock.releaseLock();
  }
}

// --- BARRERA DE SEGURIDAD BACKEND ---
function verificarAccesoServidor() {
  const raw = PropertiesService.getScriptProperties().getProperty('ACTIVE_SESSIONS');
  
  // Si la memoria está vacía (alguien usó el botón de pánico)
  if (!raw) {
     throw new Error("🔒 SEGURIDAD: El sistema fue liberado o tu sesión fue cerrada remotamente. Recarga la página.");
  }
  
  let sessions = JSON.parse(raw);
  let now = new Date().getTime();
  let hayAlguienActivo = false;
  
  // Verificamos si hay al menos una sesión viva que no haya expirado
  for (let user in sessions) {
     // CORRECCIÓN: Leemos exactamente la propiedad lastPing del nuevo modelo de datos
     let ultimoPing = sessions[user].lastPing || 0; 
     
     if (now - ultimoPing < 1800000) { // 30 minutos (1800000 ms)
        hayAlguienActivo = true;
        break;
     }
  }
  
  // Si pasó el tiempo y caducó, bloqueamos la acción
  if (!hayAlguienActivo) {
     throw new Error("🔒 SEGURIDAD: Tu tiempo de sesión ha expirado por inactividad. Recarga la página e inicia sesión.");
  }
}