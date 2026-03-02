/**
 * Controller.gs - VERSIÓN V27 (FIX CONFLICTO DE VARIABLES)
 * - FIX: Se renombró la variable de ID para evitar conflictos con Repository.gs.
 * - FIX: Conexión forzada por ID para evitar desconexiones.
 * - MANTIENE: Lógica de Fechas, Agrupación y Validaciones.
 */

// ==========================================
// 0. CONFIGURACIÓN Y HERRAMIENTAS
// ==========================================
// Usamos un nombre ÚNICO para evitar choque con otros archivos

// BORRAMOS EL ID FIJO Y USAMOS EL GESTOR DE ENTORNOS
function obtenerSpreadsheet() {
  try {
    return SpreadsheetApp.openById(getActiveDbId());
  } catch (e) {
    throw new Error(
      "No se pudo abrir la base configurada para el entorno actual (" +
        getCurrentEnvironment() +
        "). ID objetivo: " +
        getActiveDbId() +
        ". Detalle: " +
        e.message,
    );
  }
}

function obtenerHojaSegura(nombre) {
  const ss = obtenerSpreadsheet();
  return ss.getSheetByName(nombre);
}

function obtenerHojaOCrear(nombre, encabezados) {
  const ss = obtenerSpreadsheet();
  let sheet = ss.getSheetByName(nombre);
  if (!sheet) {
    sheet = ss.insertSheet(nombre);
    if (encabezados && encabezados.length > 0) sheet.appendRow(encabezados);
  }
  return sheet;
}

// Auxiliar para normalizar fechas visualmente
function _fmtFechaDisplay(valor) {
  if (!valor) return "---";
  if (valor instanceof Date) {
    return Utilities.formatDate(
      valor,
      Session.getScriptTimeZone(),
      "dd/MM/yyyy",
    );
  }
  return String(valor).trim();
}

function _normalizarUnidadLabel(unidadRaw) {
  const u = String(unidadRaw || "")
    .toLowerCase()
    .trim();
  if (u.includes("unid") || u.includes("pza") || u.includes("pieza"))
    return "Pza";
  if (u.includes("kg") || u.includes("kilo")) return "Kg";
  return "L";
}
// ==========================================
// 1. CATÁLOGOS
// ==========================================
function obtenerCatalogos() {
  const sProd = obtenerHojaOCrear("PRODUCTOS", [
    "ID",
    "NOMBRE",
    "DESCRIPCION",
    "UNIDAD",
  ]);
  const sPres = obtenerHojaOCrear("PRESENTACIONES", [
    "ID",
    "NOMBRE",
    "VOLUMEN",
  ]);
  const sUbic = obtenerHojaOCrear("UBICACIONES", ["ID", "NOMBRE"]);
  const sInv = obtenerHojaOCrear("INVENTARIO", []);

  const leer = (s, esPres) => {
    if (!s || s.getLastRow() < 2) return [];
    return s
      .getDataRange()
      .getValues()
      .slice(1)
      .map((r) => ({
        id: String(r[0]).trim(),
        nombre: r[1],
        volumen: esPres ? (Number(String(r[2]).replace(",", ".")) || 0) : 0,
        unidad: esPres ? "" : r[3] || "L", // <--- Aquí ya lee si es L, Kg o Pza
      }))
      .filter((i) => i.id);
  };

  let productos = leer(sProd);
  const presentaciones = leer(sPres, true);
  const ubicaciones = leer(sUbic);

  // --- Calcular stock total por producto ---
  let stockPorProducto = {};
  if (sInv && sInv.getLastRow() > 1) {
    const dataInv = sInv.getDataRange().getValues();
    for (let i = 1; i < dataInv.length; i++) {
      const pId = String(dataInv[i][0]).trim();
      const stock = Number(dataInv[i][3]) || 0;
      if (!stockPorProducto[pId]) stockPorProducto[pId] = 0;
      stockPorProducto[pId] += stock;
    }
  }

  // Adjuntar el stock a cada producto
  productos = productos.map((p) => {
    p.stockTotal = stockPorProducto[p.id] || 0;
    return p;
  });

  return {
    productos: productos,
    presentaciones: presentaciones,
    ubicaciones: ubicaciones,
  };
}

// ==========================================
// 1. CATÁLOGOS
// ==========================================
// ==========================================
// 1. CATÁLOGOS
// ==========================================
const leer = (s, esPres) => {
  if (!s || s.getLastRow() < 2) return [];
  return s
    .getDataRange()
    .getValues()
    .slice(1)
    .map((r) => ({
      id: String(r[0]).trim(),
      nombre: r[1],
      volumen: esPres ? (Number(String(r[2]).replace(",", ".")) || 0) : 0,
      unidad: esPres ? "" : r[3] || "L", // <--- AHORA LEE LA UNIDAD (L, Kg, Pza)
    }))
    .filter((i) => i.id);
};

// 1. REEMPLAZA TU FUNCIÓN ACTUAL 'obtenerListaClientes' POR ESTA:
function obtenerListaClientes() {
  const s = obtenerHojaSegura("CLIENTES");
  if (!s || s.getLastRow() < 2) return [];
  
  const data = s.getDataRange().getDisplayValues();
  let clientes = [];
  
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    
    let dirs = [];
    try {
       // Extrae las direcciones guardadas en la Columna G (Índice 6)
       dirs = JSON.parse(data[i][6] || "[]");
    } catch(e) {
       // Compatibilidad si tenías direcciones viejas
       if (data[i][3] && !data[i][3].includes("terno")) dirs = [data[i][3]]; 
    }

    let tipoActual = String(data[i][3]).trim();
    if (!tipoActual.includes("terno") && !tipoActual.includes("nterno")) tipoActual = "Externo"; 

    clientes.push({
      id: String(data[i][0]).trim(),
      nombre: String(data[i][1]).trim(),
      empresa: String(data[i][2]).trim(),
      tipo: tipoActual,
      telefono: String(data[i][4]).trim(),
      correo: String(data[i][5]).trim(), // <--- Esto enciende los correos en la tarjeta
      email: String(data[i][5]).trim(),
      direcciones: Array.isArray(dirs) ? dirs : []
    });
  }
  return clientes;
}

// ==========================================
// 2. GESTIÓN DE INVENTARIO (CONSOLIDACIÓN)
// ==========================================
function obtenerDatosUbicaciones() {
  const sInv = obtenerHojaOCrear("INVENTARIO", [
    "ID_PROD",
    "ID_PRES",
    "ID_UBIC",
    "STOCK",
    "CADUCIDAD",
    "ELABORACION",
    "LOTE",
    "F_ENTRADA",
    "PROVEEDOR",
  ]);
  const sUbic = obtenerHojaOCrear("UBICACIONES", []);
  const sProd = obtenerHojaOCrear("PRODUCTOS", []);
  const sPres = obtenerHojaOCrear("PRESENTACIONES", []);

  const mapProd = {},
    mapPres = {},
    mapPresVol = {};

  if (sProd.getLastRow() > 1) {
    sProd
      .getDataRange()
      .getValues()
      .slice(1)
      .forEach((r) => {
        mapProd[String(r[0]).trim()] = { nombre: r[1], unidad: r[3] || "L" };
      });
  }

  if (sPres.getLastRow() > 1) {
    sPres
      .getDataRange()
      .getValues()
      .slice(1)
      .forEach((r) => {
        const id = String(r[0]).trim();
        mapPres[id] = r[1];
        mapPresVol[id] = Number(r[2]) || 0;
      });
  }

  let ubicaciones = [];
  if (sUbic.getLastRow() > 1) {
    sUbic
      .getDataRange()
      .getValues()
      .slice(1)
      .forEach((r) => {
        if (r[0])
          ubicaciones.push({
            id: String(r[0]).trim(),
            nombre: r[1] || "S/N",
            items: [],
            totales: { L: 0, Kg: 0, Pza: 0 }, // SEPARAMOS LOS TOTALES
          });
      });
  }

  if (sInv.getLastRow() > 1) {
    const dInv = sInv.getDataRange().getValues();
    for (let i = 1; i < dInv.length; i++) {
      const stock = Number(dInv[i][3]);
      if (stock > 0.001) {
        const uId = String(dInv[i][2]).trim();
        let ubic = ubicaciones.find((u) => u.id === uId);
        if (!ubic) {
          ubic = {
            id: uId,
            nombre: "Ubic: " + uId,
            items: [],
            totales: { L: 0, Kg: 0, Pza: 0 },
          };
          ubicaciones.push(ubic);
        }

        const pId = String(dInv[i][0]).trim();
        const prId = String(dInv[i][1]).trim();
        const lote = String(dInv[i][6]).trim();
        const caducidadStr = _fmtFechaDisplay(dInv[i][4]);

        const prodData = mapProd[pId] || { nombre: pId, unidad: "L" };
        const nProd = prodData.nombre;

        // Estandarizamos la unidad a L, Kg o Pza
        let unidadRaw = String(prodData.unidad).toLowerCase();
        let uLower = "L";
        if (
          unidadRaw.includes("unid") ||
          unidadRaw.includes("pza") ||
          unidadRaw.includes("pieza")
        )
          uLower = "Pza";
        else if (unidadRaw.includes("kg") || unidadRaw.includes("kilo"))
          uLower = "Kg";

        const nPres = mapPres[prId] || prId;

        let itemExistente = ubic.items.find(
          (it) =>
            it.raw_producto_id === pId &&
            it.raw_presentacion_id === prId &&
            it.lote === lote &&
            it.caducidad === caducidadStr,
        );

        if (itemExistente) {
          itemExistente.volumen += stock;
        } else {
          ubic.items.push({
            raw_producto_id: pId,
            producto: nProd,
            raw_presentacion_id: prId,
            presentacion: nPres,
            lote: lote,
            volumen: stock,
            volumen_nominal: mapPresVol[prId] || 0,
            caducidad: caducidadStr,
            nombre_completo: `${nProd} (${nPres})`,
            unidad: uLower, // Guardamos la unidad del item
          });
        }
        // Sumamos a la canasta que corresponde
        if (!ubic.totales[uLower]) ubic.totales[uLower] = 0;
        ubic.totales[uLower] += stock;
      }
    }
  }
  return ubicaciones;
}

function obtenerDatosProductos() {
  const sInv = obtenerHojaOCrear("INVENTARIO", []);
  if (sInv.getLastRow() < 2) return [];
  
  // Agregamos mapPresVol para poder calcular equivalencias
  const mapProd = {}, mapUbic = {}, mapPres = {}, mapPresVol = {};
  const sP = obtenerHojaOCrear("PRODUCTOS", []),
        sU = obtenerHojaOCrear("UBICACIONES", []),
        sPr = obtenerHojaOCrear("PRESENTACIONES", []);
        
  if (sP.getLastRow() > 1) {
    sP.getDataRange().getValues().slice(1).forEach((r) => {
        mapProd[String(r[0]).trim()] = { nombre: r[1], unidad: r[3] || "L" };
    });
  }
  if (sU.getLastRow() > 1) {
    sU.getDataRange().getValues().slice(1).forEach((r) => (mapUbic[String(r[0]).trim()] = r[1]));
  }
  if (sPr.getLastRow() > 1) {
    sPr.getDataRange().getValues().slice(1).forEach((r) => {
        mapPres[String(r[0]).trim()] = r[1];
        mapPresVol[String(r[0]).trim()] = Number(r[2]) || 0; // Extraemos el volumen nominal
    });
  }
  
  const dataInv = sInv.getDataRange().getValues();
  let productosMap = {};

  for (let i = 1; i < dataInv.length; i++) {
    const stock = Number(dataInv[i][3]);
    if (stock > 0.001) {
      const pId = String(dataInv[i][0]).trim();
      const prodObj = mapProd[pId] || { nombre: pId, unidad: "L" };
      const rawName = prodObj.nombre;
      let baseName = rawName;
      let subName = "";
      let match = rawName.match(/(.*)\(([^)]+)\)$/);
      
      if (match) {
        subName = match[1].trim();
        baseName = match[2].trim();
      }

      if (!productosMap[baseName]) {
        productosMap[baseName] = {
          id: pId,
          nombre: baseName,
          unidad: prodObj.unidad,
          totalVolumen: 0,
          lotes: [],
        };
      }
      productosMap[baseName].totalVolumen += stock;

      const uId = String(dataInv[i][2]).trim();
      const prId = String(dataInv[i][1]).trim();
      const uName = mapUbic[uId] || uId;
      let presName = mapPres[prId] || prId;
      const cadStr = _fmtFechaDisplay(dataInv[i][4]);
      
      let loteExistente = productosMap[baseName].lotes.find(
        (l) => l.lote === String(dataInv[i][6]).trim() &&
               l.ubicacion === uName &&
               l.presentacion === presName &&
               l.caducidad === cadStr &&
               l.alias === subName
      );
if (loteExistente) {
          loteExistente.volumen += stock;
        } else {
          productosMap[baseName].lotes.push({
            lote: String(dataInv[i][6]).trim(),
            volumen: stock,
            ubicacion: uName,
            ubicacion_id: uId,           
            presentacion: presName,
            presentacion_id: prId,       
            alias: subName, 
            caducidad: cadStr,
            volumen_nominal: mapPresVol[prId] || 0,
            
            // --- NUEVO: IDENTIDAD REAL DE LA VARIANTE ---
            raw_producto_id: pId,           
            raw_producto_nombre: rawName    
          });
        }
    }
  }

  return Object.values(productosMap).map((p) => {
    let uLower = String(p.unidad).toLowerCase();
    if (uLower.includes("unid") || uLower.includes("pza")) p.unidad = "Pza";
    else if (uLower.includes("kg") || uLower.includes("kilo")) p.unidad = "Kg";
    else p.unidad = "L";
    return p;
  });
}

// ==========================================
// 3. REGISTRO DE MOVIMIENTOS
// ==========================================
function registrarNuevoProducto(datos) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(10000);
    verificarAccesoServidor(); // Bloqueo de concurrencia para evitar registros simultáneos
    const sheet = obtenerHojaOCrear("PRODUCTOS", [
      "ID",
      "NOMBRE",
      "DESCRIPCION",
      "UNIDAD",
    ]);
    const data = sheet.getDataRange().getValues();

    // Nos aseguramos de que esté en mayúsculas
    const nombreNuevo = String(datos.nombre).trim().toUpperCase();

    // SUPER VALIDACIÓN: Quitamos todos los espacios para comparar
    // Esto hace que "PERFECT DUO" y "PERFECTDUO" se detecten como iguales
    const nombreNuevoSinEspacios = nombreNuevo.replace(/\s+/g, "");

    for (let i = 1; i < data.length; i++) {
      let nombreExistente = String(data[i][1]).trim().toUpperCase();
      let existenteSinEspacios = nombreExistente.replace(/\s+/g, "");

      // Comparamos sin espacios
      if (existenteSinEspacios === nombreNuevoSinEspacios) {
        throw new Error(
          `El producto "${nombreExistente}" ya existe en tu catálogo. No puedes registrarlo dos veces.`,
        );
      }
    }

    // Si pasa la validación, lo registramos usando el nombre con espacios correctos
    const nuevoId = Utilities.getUuid();
    sheet.appendRow([nuevoId, nombreNuevo, datos.descripcion, datos.unidad]);

    return true;
  } catch (error) {
    // Si hubo un error (como el duplicado), lo enviamos al frontend
    throw new Error(error.message);
  } finally {
    lock.releaseLock();
  }
}

function registrarNuevaPresentacion(n) {
  let v = 0,
    m = String(n).match(/[\d\.]+/);
  if (m) {
    v = parseFloat(m[0]);
    if (String(n).toLowerCase().includes("ml")) v /= 1000;
  }
  obtenerHojaOCrear("PRESENTACIONES").appendRow([Utilities.getUuid(), n, v]);
  return true;
}
function registrarNuevaUbicacion(n) {
  obtenerHojaOCrear("UBICACIONES").appendRow([Utilities.getUuid(), n]);
  return true;
}

function registrarEntradaUnica(datos) {
  return registrarEntradaMasiva([datos]);
}

function registrarEntradaMasiva(lista) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    verificarAccesoServidor();
    const sInv = obtenerHojaOCrear("INVENTARIO", [
      "ID_PROD",
      "ID_PRES",
      "ID_UBIC",
      "STOCK",
      "CADUCIDAD",
      "ELABORACION",
      "LOTE",
      "F_ENTRADA",
      "PROVEEDOR",
    ]);
    const sLog = obtenerHojaOCrear("REGISTROS_ENTRADA", [
      "FECHA",
      "PROD",
      "PRES",
      "UBIC",
      "VOL",
      "LOTE",
      "PROVEEDOR",
      "TIPO",
    ]);

    lista.forEach((d) => {
      if (d.producto_id.includes("Selecciona") || d.producto_id === "undefined")
        throw new Error("Producto inválido.");

      let fElab = d.fecha_elaboracion;
      let fCad = d.fecha_caducidad;

      if (!fElab) {
        let now = new Date();
        fElab = Utilities.formatDate(
          now,
          Session.getScriptTimeZone(),
          "yyyy-MM-dd",
        );
      }

      if (!fCad || fCad === "") {
        let partes = String(fElab).split("-");
        if (partes.length === 3) {
          let anio = parseInt(partes[0]) + 2;
          fCad = `${anio}-${partes[1]}-${partes[2]}`;
        } else {
          fCad = "SIN-FECHA";
        }
      }

      sInv.appendRow([
        d.producto_id,
        d.presentacion_id,
        d.ubicacion_id,
        Number(d.volumen_L),
        fCad,
        fElab,
        d.lote,
        new Date(),
        d.proveedor,
      ]);

      sLog.appendRow([
        new Date(),
        d.producto_id,
        d.presentacion_id,
        d.ubicacion_id,
        d.volumen_L,
        d.lote,
        d.proveedor,
        "Entrada",
      ]);
    });
    return true;
  } catch (e) {
    throw e;
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// 4. MOVER Y TRANSFORMAR
// ==========================================
function transferirProducto(
  origenId,
  destinoId,
  productoId,
  presentacionId,
  lote,
  cantidadMover,
) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    verificarAccesoServidor();
    const sInv = obtenerHojaOCrear("INVENTARIO", []);
    const data = sInv.getDataRange().getValues();

    let filasOrigen = [],
      stockTotalDisp = 0;
    for (let i = 1; i < data.length; i++) {
      if (
        String(data[i][0]) == String(productoId) &&
        String(data[i][1]) == String(presentacionId) &&
        String(data[i][2]) == String(origenId) &&
        String(data[i][6]) == String(lote)
      ) {
        filasOrigen.push({
          index: i + 1,
          stock: Number(data[i][3]),
          rowData: data[i],
        });
        stockTotalDisp += Number(data[i][3]);
      }
    }

    if (filasOrigen.length === 0) throw new Error("Origen no encontrado");
    if (stockTotalDisp < cantidadMover - 0.001)
      throw new Error("Stock insuficiente");

    let restante = cantidadMover;
    for (let item of filasOrigen) {
      if (restante <= 0) break;
      let restar = Math.min(item.stock, restante);
      let ns = item.stock - restar;
      sInv.getRange(item.index, 4).setValue(ns <= 0.001 ? 0 : ns);
      restante -= restar;
    }

    sInv.appendRow([
      productoId,
      presentacionId,
      destinoId,
      Number(cantidadMover),
      filasOrigen[0].rowData[4],
      filasOrigen[0].rowData[5],
      lote,
      new Date(),
      "Transferencia",
    ]);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

function realizarTransformacion(datos) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    verificarAccesoServidor();
    const sInv = obtenerHojaOCrear("INVENTARIO", []);
    const data = sInv.getDataRange().getValues();

    let filasOrigen = [],
      stockTotalDisp = 0;
    for (let i = 1; i < data.length; i++) {
      if (
        String(data[i][2]) == String(datos.origenId) &&
        String(data[i][0]) == String(datos.productoIdOrigen) &&
        String(data[i][6]).toUpperCase() ==
          String(datos.loteOrigen).toUpperCase()
      ) {
        filasOrigen.push({
          index: i + 1,
          stock: Number(data[i][3]),
          rowData: data[i],
        });
        stockTotalDisp += Number(data[i][3]);
      }
    }

    if (filasOrigen.length === 0) throw new Error("Lote origen no encontrado");
    if (stockTotalDisp < Number(datos.cantidadLitros))
      throw new Error("Stock insuficiente");

    let restante = Number(datos.cantidadLitros);
    for (let item of filasOrigen) {
      if (restante <= 0) break;
      let restar = Math.min(item.stock, restante);
      let ns = item.stock - restar;
      sInv.getRange(item.index, 4).setValue(ns <= 0.001 ? 0 : ns);
      restante -= restar;
    }

    const lFin = datos.nuevoLote || datos.loteOrigen;
    sInv.appendRow([
      datos.nuevoProductoId,
      datos.nuevaPresentacionId,
      datos.origenId,
      Number(datos.cantidadLitros),
      filasOrigen[0].rowData[4],
      filasOrigen[0].rowData[5],
      lFin,
      new Date(),
      "Transformación",
    ]);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// 5. PROCESAMIENTO DE SALIDAS (PEDIDOS)
// ==========================================
function registrarSalidas(lista) {
  return procesarPedidoCompleto(
    {
      idCliente: "GENERICO",
      nombreCliente: "Salida Rápida",
      direccion: "-",
      telefono: "-",
      email: "-",
      tipoEnvio: "Nacional",
      paqueteria: "Mostrador",
      guia: "-",
      costoEnvio: 0,
    },
    lista,
    false,
  );
}

function procesarPedidoCompleto(datosPedido, itemsCarrito, guardarCliente) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    verificarAccesoServidor();
    itemsCarrito.forEach((item) => {
      const p = String(item.producto_id);
      if (p.includes("Selecciona") || p === "undefined" || !p)
        throw new Error("⚠️ Producto inválido.");
      if (!item.lote || item.lote === "undefined")
        throw new Error("⚠️ Lote inválido.");
    });

    const unicos = {};
    itemsCarrito.forEach((i) => {
      const k = `${i.producto_id}|${i.presentacion_id}|${i.lote}|${i.ubicacion_id}`;
      if (unicos[k]) {
        unicos[k].volumen_L += Number(i.volumen_L);
        unicos[k].piezas += Number(i.piezas);
      } else {
        unicos[k] = {
          ...i,
          volumen_L: Number(i.volumen_L),
          piezas: Number(i.piezas),
        };
      }
    });
    const itemsProcesar = Object.values(unicos);

    const idPedido = "PED-" + Math.floor(Date.now() / 1000);
    const fechaHoy = new Date();

if (guardarCliente) {
      const sCli = obtenerHojaOCrear("CLIENTES", [
        "ID",
        "NOMBRE",
        "EMPRESA",
        "DIRECCION",
        "TELEFONO",
        "EMAIL",
      ]);
      let idCli = datosPedido.idCliente;
      if (!idCli || idCli === "nuevo" || idCli === "") {
        idCli = "CLI-" + Math.floor(Math.random() * 10000);
      }
      sCli.appendRow([
        idCli,
        datosPedido.nombreCliente,
        datosPedido.empresa,
        datosPedido.direccion,
        datosPedido.telefono,
        datosPedido.email,
      ]);
      
      // 🔥 CORRECCIÓN CLAVE: Le asignamos el nuevo ID al pedido 
      // para que quede enlazado permanentemente al cliente en la base de datos.
      datosPedido.idCliente = idCli;
    }

    const sPed = obtenerHojaOCrear("PEDIDOS", [
      "ID_PEDIDO",
      "FECHA",
      "ID_CLIENTE",
      "NOMBRE",
      "DIRECCION",
      "TELEFONO",
      "PAQUETERIA",
      "GUIA",
      "TIPO",
      "COSTO",
      "ESTATUS",
      "F_EST",
      "F_REAL",
      "LINK",
    ]);
    sPed.appendRow([
      idPedido,
      fechaHoy,
      datosPedido.idCliente,
      datosPedido.nombreCliente,
      datosPedido.direccion,
      datosPedido.telefono,
      datosPedido.paqueteria,
      datosPedido.guia,
      datosPedido.tipoEnvio,
      datosPedido.costoEnvio,
      "Pendiente",
      "",
      "",
      "",
    ]);

    // <--- CAMBIO 1: Agregamos 'PIEZAS' a los encabezados de la hoja DETALLE_PEDIDOS
    const sInv = obtenerHojaOCrear("INVENTARIO", []),
      sSal = obtenerHojaOCrear("REGISTROS_SALIDA", []),
      sDet = obtenerHojaOCrear("DETALLE_PEDIDOS", [
        "ID_PEDIDO",
        "PRODUCTO",
        "PRESENTACION",
        "LOTE",
        "VOLUMEN",
        "PIEZAS",
        "UNIDAD",
      ]);
    const dataInv = sInv.getDataRange().getValues();

    itemsProcesar.forEach((item) => {
      // <--- CAMBIO 2: Agregamos Number(item.piezas || 0) a la fila que se va a guardar
      sDet.appendRow([
        idPedido,
        item.nombre_producto,
        item.nombre_presentacion || "---",
        item.lote,
        Number(item.volumen_L),
        Number(item.piezas || 0),
        _normalizarUnidadLabel(item.unidad_medida),
      ]);

      let restante = item.volumen_L;
      for (let i = 1; i < dataInv.length; i++) {
        if (restante <= 0.001) break;
        const invProd = String(dataInv[i][0]).trim();
        const invLote = String(dataInv[i][6]).trim().toUpperCase();
        const invUbic = String(dataInv[i][2]).trim();

        if (
          invProd === String(item.producto_id).trim() &&
          invLote === String(item.lote).trim().toUpperCase() &&
          invUbic === String(item.ubicacion_id).trim()
        ) {
          const stock = Number(dataInv[i][3]);
          if (stock > 0) {
            const restar = Math.min(stock, restante);
            sInv.getRange(i + 1, 4).setValue(stock - restar);
            restante -= restar;
          }
        }
      }
      sSal.appendRow([
        item.producto_id,
        item.nombre_producto,
        item.presentacion_id,
        item.nombre_presentacion,
        item.volumen_L,
        item.piezas,
        item.ubicacion_id,
        item.lote,
        datosPedido.nombreCliente,
        Session.getActiveUser().getEmail(),
        fechaHoy,
        idPedido,
      ]);
    });

    return { success: true, idPedido: idPedido };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// 6. HISTORIAL DE ENVÍOS
// ==========================================
function obtenerHistorialPedidos() {
  const sPed = obtenerHojaSegura("PEDIDOS");
  if (!sPed) return [];
  const data = sPed.getDataRange().getDisplayValues();
  if (data.length < 2) return [];

  const sDet = obtenerHojaSegura("DETALLE_PEDIDOS"),
    sCli = obtenerHojaSegura("CLIENTES");
  const mapProd = {},
    mapEmp = {};
  if (sDet) {
    const dDet = sDet.getDataRange().getDisplayValues();
    for (let i = 1; i < dDet.length; i++) {
      let id = String(dDet[i][0]).trim();
      if (id) {
        if (!mapProd[id]) mapProd[id] = [];
        
        let prodName = dDet[i][1];
        let presentacion = String(dDet[i][2]).toLowerCase();
        let volumen = Number(dDet[i][4]) || 0;
        
        // Deducir unidad de la presentación en lugar de usar las piezas
        let uni = "L";
        if (presentacion.includes("kg") || presentacion.includes("kilo") || presentacion.includes("gramo")) {
            uni = "Kg";
        } else if (presentacion.includes("unid") || presentacion.includes("pza") || presentacion.includes("pieza")) {
            uni = "Pza";
        }
        
        // Si es pieza quitamos decimales, si no, lo dejamos normal
        let volMostrar = (uni === "Pza") ? Math.round(volumen) : volumen;
        
        mapProd[id].push(`${prodName} (${volMostrar} ${uni})`);
      }
    }
  }

  if (sCli)
    sCli
      .getDataRange()
      .getDisplayValues()
      .forEach((r, i) => {
        if (i > 0) mapEmp[String(r[0]).trim()] = r[2];
      });

  let pedidos = [];
  for (let i = data.length - 1; i >= 1; i--) {
    let r = data[i],
      id = String(r[0]).trim();
    if (id) {
      let prods = mapProd[id] || [],
        resumen = prods.length > 0 ? prods.join(", ") : "---";
      if (resumen.length > 55) resumen = resumen.substring(0, 55) + "...";
      pedidos.push({
        id: id,
        fecha: r[1],
        cliente: r[3],
        empresa: mapEmp[String(r[2]).trim()] || "",
        resumenProductos: resumen,
        logistica: r[6],
        guia: r[7],
        estatus: r[10] || "Pendiente",
        costo: r[9] ? String(r[9]).replace(/[^0-9.]/g, "") : 0,
      });
    }
  }
  return pedidos;
}

function actualizarPedido(id, fe, fr, st, guia) {
  const s = obtenerHojaSegura("PEDIDOS");
  const d = s.getDataRange().getValues();
  for (let i = 1; i < d.length; i++) {
    if (String(d[i][0]).trim() === String(id).trim()) {
      // 1. Actualizar Estatus y Fechas (Código original)
      s.getRange(i + 1, 11).setValue(st);
      s.getRange(i + 1, 13).setValue(fe);
      s.getRange(i + 1, 14).setValue(fr);

      // 2. --- NUEVO: Actualizar Guía si se envió ---
      if (guia && guia.trim() !== "") {
        s.getRange(i + 1, 8).setValue(guia); // Columna H es la 8
      }
      // -------------------------------------------

      return "OK";
    }
  }
}

// FIFO & AUX
function obtenerSugerenciaFIFO(productoId, carrito) {
  try {
    const prodIdBuscado = String(productoId).trim();
    const sheetInv = obtenerHojaSegura("INVENTARIO");
    if (!sheetInv || sheetInv.getLastRow() < 2)
      return JSON.stringify({ success: false, error: "Sin stock" });
    const data = sheetInv.getDataRange().getValues();

    // 1. Obtener Nombres
    const mapUbic = {}, mapPres = {};
    const sU = obtenerHojaSegura("UBICACIONES"), sP = obtenerHojaSegura("PRESENTACIONES");
    
    if (sU) sU.getDataRange().getValues().slice(1).forEach((r) => (mapUbic[String(r[0]).trim()] = r[1]));
    if (sP) sP.getDataRange().getValues().slice(1).forEach((r) => (mapPres[String(r[0]).trim()] = r[1]));
    
    let lotes = [];

    // 2. Leer y Agrupar
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === prodIdBuscado) {
        const uId = String(data[i][2]).trim();
        const presId = String(data[i][1]).trim();
        const lote = String(data[i][6]).trim();
        const caducidadStr = _fmtFechaDisplay(data[i][4]);
        const stock = Number(data[i][3]);
        
        let existente = lotes.find(
          (l) =>
            l.presentacion_id === presId &&
            l.ubicacion_id === uId &&
            l.lote === lote &&
            _fmtFechaDisplay(l.caducidad) === caducidadStr,
        );
        if (existente) {
          existente.stock_real += stock;
        } else {
          lotes.push({
            producto_id: prodIdBuscado, // <--- GUARDAMOS EL ID DEL PRODUCTO
            presentacion_id: presId,
            ubicacion_id: uId,
            stock_real: stock,
            caducidad: data[i][4],
            elaboracion: data[i][5],
            lote: lote,
            fecha_entrada: data[i][7],
            nombre_ubicacion: mapUbic[uId] || uId,
            nombre_presentacion: mapPres[presId] || presId,
          });
        }
      }
    }

    // 3. Descontar Carrito (LA MAGIA DE LA CORRECCIÓN)
    if (carrito && Array.isArray(carrito)) {
      carrito.forEach((item) => {
        const l = lotes.find(
          (x) =>
            x.producto_id === String(item.producto_id).trim() && // <--- FIX: EXIGIMOS QUE SEA EL MISMO PRODUCTO
            x.lote === String(item.lote).trim() &&
            x.ubicacion_id === String(item.ubicacion_id).trim(),
        );
      
        if (l) l.stock_real -= Number(item.volumen_L);
      });
    }

    const validos = lotes.filter((l) => l.stock_real > 0.001);
    if (validos.length === 0)
      return JSON.stringify({ success: false, error: "Sin stock disponible" });
      
    const getMs = (d) => (d instanceof Date ? d.getTime() : 0);
    validos.sort(
      (a, b) =>
        getMs(a.elaboracion || a.fecha_entrada) -
        getMs(b.elaboracion || b.fecha_entrada),
    );
    const mejor = validos[0];

    return JSON.stringify({
      success: true,
      mejor_candidato: {
        presentacion_id: mejor.presentacion_id,
        ubicacion_id: mejor.ubicacion_id,
        lote: mejor.lote,
        stock_real: mejor.stock_real,
      },
      lista_completa: validos.map((l) => ({
        presentacion_id: l.presentacion_id,
        ubicacion_id: l.ubicacion_id,
        lote: l.lote,
        stock: l.stock_real.toFixed(2),
        caducidad: _fmtFechaDisplay(l.caducidad),
        nombre_ubicacion: l.nombre_ubicacion,
        nombre_presentacion: l.nombre_presentacion,
        es_sugerido:
          l.lote === mejor.lote && l.ubicacion_id === mejor.ubicacion_id,
      })),
    });
  } catch (e) {
    return JSON.stringify({ success: false, error: e.message });
  }
}

function _fmtF(f) {
  try {
    if (!f) return "";
    let s = String(f).trim();

    // FILTRO ANTI-URL: Si encuentra un enlace en el Excel por error, lo ignora.
    if (
      s.toLowerCase().startsWith("http") ||
      s.toLowerCase().includes("drive.google")
    ) {
      return "";
    }

    if (s.includes("-")) {
      let p = s.split("-");
      if (p.length >= 3)
        return p[0].length === 4
          ? s.substring(0, 10)
          : `${p[2].split(" ")[0]}-${p[1]}-${p[0]}`;
    }
    if (s.includes("/")) {
      let p = s.split("/");
      if (p.length >= 3) return `${p[2].split(" ")[0]}-${p[1]}-${p[0]}`;
    }
    return "";
  } catch (e) {
    return ""; // A prueba de balas
  }
}

function formatearFechaInput(f) {
  return _fmtF(f);
}

function EJECUTAR_DIAGNOSTICO_DEBUG() {
  const ss = obtenerSpreadsheet();
  let sD = ss.getSheetByName("DEBUG");
  if (!sD) sD = ss.insertSheet("DEBUG");
  sD.clear();
  let rep = [["HOJA", "EXISTE?", "FILAS"]];
  ["PEDIDOS", "DETALLE_PEDIDOS", "CLIENTES", "INVENTARIO"].forEach((n) => {
    let s = ss.getSheetByName(n);
    rep.push([n, s ? "SÍ" : "NO", s ? s.getLastRow() : "-"]);
  });
  sD.getRange(1, 1, rep.length, 3).setValues(rep);
}

// ==========================================
// 7. MÓDULO DE DESINCORPORACIÓN (BAJAS)
// ==========================================

function obtenerMaterialesCaducados() {
  const sInv = obtenerHojaSegura("INVENTARIO");
  if (!sInv || sInv.getLastRow() < 2) return [];

  const mapProd = {},
    mapPres = {},
    mapUbic = {};
  const sP = obtenerHojaSegura("PRODUCTOS");
  const sPr = obtenerHojaSegura("PRESENTACIONES");
  const sU = obtenerHojaSegura("UBICACIONES");

  if (sP)
    sP.getDataRange()
      .getValues()
      .slice(1)
      .forEach(
        (r) =>
          (mapProd[String(r[0]).trim()] = {
            nombre: r[1],
            unidad: _normalizarUnidadLabel(r[3]),
          }),
      );
  if (sPr)
    sPr
      .getDataRange()
      .getValues()
      .slice(1)
      .forEach((r) => (mapPres[String(r[0]).trim()] = r[1]));
  if (sU)
    sU.getDataRange()
      .getValues()
      .slice(1)
      .forEach((r) => (mapUbic[String(r[0]).trim()] = r[1]));

  const dataInv = sInv.getDataRange().getValues();
  let caducados = [];

  // Establecemos "hoy" a la medianoche para una comparación justa
  const hoy = new Date();
  hoy.setHours(0, 0, 0, 0);

  for (let i = 1; i < dataInv.length; i++) {
    const stock = Number(dataInv[i][3]);

    if (stock > 0.001) {
      const caducidad = dataInv[i][4]; // Columna E (Caducidad)
      let fechaCad = null;

      // Transformar la fecha para poder compararla
      if (caducidad instanceof Date) {
        fechaCad = caducidad;
      } else if (typeof caducidad === "string") {
        if (caducidad.includes("-")) {
          const partes = caducidad.split("T")[0].split("-"); // YYYY-MM-DD
          if (partes.length === 3)
            fechaCad = new Date(partes[0], partes[1] - 1, partes[2]);
        } else if (caducidad.includes("/")) {
          const partes = caducidad.split("/"); // DD/MM/YYYY
          if (partes.length === 3)
            fechaCad = new Date(partes[2], partes[1] - 1, partes[0]);
        }
      }

      // Si tiene fecha válida y ya pasó (es menor a hoy)
      if (fechaCad && fechaCad < hoy) {
        const pId = String(dataInv[i][0]).trim();
        const prId = String(dataInv[i][1]).trim();
        const uId = String(dataInv[i][2]).trim();

        caducados.push({
          producto_id: pId,
          producto: (mapProd[pId] && mapProd[pId].nombre) || pId,
          presentacion_id: prId,
          presentacion: mapPres[prId] || prId,
          ubicacion_id: uId,
          ubicacion: mapUbic[uId] || uId,
          lote: String(dataInv[i][6]).trim(),
          volumen: stock,
          unidad: (mapProd[pId] && mapProd[pId].unidad) || "L",
          caducidadStr: _fmtFechaDisplay(caducidad),
        });
      }
    }
  }
  return caducados;
}

// ==========================================
// FASE 3 (MEJORADA): DRIVE + PLANTILLA
// ==========================================

// ¡IMPORTANTE! PEGA AQUÍ EL ID DE TU CARPETA "REPORTES_BAJAS" DE DRIVE
const ID_CARPETA_BAJAS_DRIVE = "151nk1FvsdYP8eRf8wo7dTfU8PXYEABcY";

function procesarBajaOficial(itemsBaja) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(45000);
    verificarAccesoServidor();

    // --- PASO 1: VERIFICAR ID DE DRIVE ---
    if (
      !ID_CARPETA_BAJAS_DRIVE ||
      ID_CARPETA_BAJAS_DRIVE.includes("drive.google.com") ||
      ID_CARPETA_BAJAS_DRIVE.includes("/")
    ) {
      throw new Error(
        "Paso 1: El ID de la carpeta de Drive es inválido. Asegúrate de poner solo las letras y números del final del link.",
      );
    }

    // --- PASO 2: BUSCAR LA PLANTILLA ---
    const sTemplate = obtenerHojaSegura("TEMPLATE_BAJAS");
    if (!sTemplate)
      throw new Error(
        "Paso 2: No se encontró la pestaña 'TEMPLATE_BAJAS'. Revisa que esté escrita exactamente así en tu Excel.",
      );

    const sInv = obtenerHojaSegura("INVENTARIO");
    const sSal = obtenerHojaSegura("REGISTROS_SALIDA");
    const dataInv = sInv.getDataRange().getValues();
    const fechaHoy = new Date();
    const timestampFile = Utilities.formatDate(
      fechaHoy,
      Session.getScriptTimeZone(),
      "yyyyMMdd_HHmm",
    );

    // --- PASO 3: DESCONTAR DEL INVENTARIO ---
    itemsBaja.forEach((item) => {
      let restante = Number(item.volumen);
      for (let i = 1; i < dataInv.length; i++) {
        if (restante <= 0.001) break;
        const invProd = String(dataInv[i][0]).trim();
        const invLote = String(dataInv[i][6]).trim().toUpperCase();
        const invUbic = String(dataInv[i][2]).trim();

        if (
          invProd === String(item.producto_id).trim() &&
          invLote === String(item.lote).trim().toUpperCase() &&
          invUbic === String(item.ubicacion_id).trim()
        ) {
          const stock = Number(dataInv[i][3]);
          if (stock > 0) {
            const restar = Math.min(stock, restante);
            sInv.getRange(i + 1, 4).setValue(stock - restar);
            restante -= restar;
          }
        }
      }
      sSal.appendRow([
        item.producto_id,
        item.producto,
        item.presentacion_id,
        item.presentacion,
        item.volumen,
        0,
        item.ubicacion_id,
        item.lote,
        "DESINCORPORACIÓN",
        Session.getActiveUser().getEmail(),
        fechaHoy,
        `BAJA-${timestampFile}`,
      ]);
    });

    // --- PASO 4: CONECTAR CON CARPETA DE DRIVE ---
    let folderDestino;
    try {
      folderDestino = DriveApp.getFolderById(ID_CARPETA_BAJAS_DRIVE);
    } catch (e) {
      throw new Error(
        "Paso 4: No se encontró la carpeta en Drive. Revisa el ID y los permisos. (" +
          e.message +
          ")",
      );
    }

    // --- PASO 5: CREAR EXCEL NUEVO Y COPIAR PLANTILLA ---
    let newDoc, newDocId;

    // NOMBRE DEL ARCHIVO ACTUALIZADO
    const docName = `Listado de Materiales para desincorporación_${timestampFile}`;

    try {
      newDoc = SpreadsheetApp.create(docName);
      newDocId = newDoc.getId();
      sTemplate.copyTo(newDoc).setName("Reporte");
      newDoc.deleteSheet(newDoc.getSheets()[0]);
    } catch (e) {
      throw new Error(
        "Paso 5: Falló al crear el Excel o copiar la plantilla. (" +
          e.message +
          ")",
      );
    }

    // --- PASO 6: RELLENAR DATOS Y FORMATO DINÁMICO ---
    const targetSheet = newDoc.getSheetByName("Reporte");
    let rowTotal = 0;

    try {
      // Encontramos la primera fila vacía (asumiendo que los encabezados terminan en la fila 3)
      const startRow = targetSheet.getLastRow() + 1;
      let totalVolumen = 0;

      // Escribimos los datos respetando las columnas (A-E)
      itemsBaja.forEach((item, index) => {
        const currentRow = startRow + index;
        targetSheet.getRange(currentRow, 1).setValue(item.lote);
        targetSheet.getRange(currentRow, 2).setValue(item.producto);
        targetSheet.getRange(currentRow, 3).setValue(""); // Nombre químico
        targetSheet
          .getRange(currentRow, 4)
          .setValue("ENVASE PET DE " + item.presentacion);
        targetSheet.getRange(currentRow, 5).setValue(Number(item.volumen));

        totalVolumen += Number(item.volumen);
      });

      const endRow = startRow + itemsBaja.length - 1;

      // COPIAR FORMATO: Si hay más de 1 item, copiamos el formato de la fila inicial hacia abajo
      if (itemsBaja.length > 1) {
        const maxCols = targetSheet.getMaxColumns();
        const firstRowRange = targetSheet.getRange(startRow, 1, 1, maxCols);
        const targetRange = targetSheet.getRange(
          startRow + 1,
          1,
          itemsBaja.length - 1,
          maxCols,
        );
        firstRowRange.copyTo(
          targetRange,
          SpreadsheetApp.CopyPasteType.PASTE_FORMAT,
          false,
        );
      }

      // LIMPIEZA Y TOTALES: Dejamos 1 fila en blanco, quitamos formato y ponemos totales
      rowTotal = endRow + 2;

      // Quita bordes y fondos de las 2 filas debajo de la tabla para un aspecto limpio
      targetSheet
        .getRange(rowTotal - 1, 1, 2, targetSheet.getMaxColumns())
        .clearFormat();

      // Inserta "TOTAL" y la suma en negritas
      targetSheet
        .getRange(rowTotal, 4)
        .setValue("TOTAL")
        .setFontWeight("bold")
        .setHorizontalAlignment("right");
      targetSheet
        .getRange(rowTotal, 5)
        .setValue(totalVolumen)
        .setFontWeight("bold");

      SpreadsheetApp.flush();
    } catch (e) {
      throw new Error(
        "Paso 6: Falló al escribir los datos en el nuevo archivo. (" +
          e.message +
          ")",
      );
    }

    // --- PASO 7: GENERAR PDF (CON TRUCO DE HOJA TEMPORAL) ---
    let excelFile = DriveApp.getFileById(newDocId);
    let pdfFile;

    // Clonamos la hoja para "destruirla" visualmente sin afectar el Excel
    const pdfSheet = targetSheet.copyTo(newDoc);
    pdfSheet.setName("TMP_PDF");
    targetSheet.hideSheet(); // Ocultamos la original, el PDF solo tomará la visible

    try {
      // En la hoja temporal, ocultamos de la columna F en adelante
      const totalCols = pdfSheet.getMaxColumns();
      if (totalCols > 5) pdfSheet.hideColumns(6, totalCols - 5);

      // Ocultamos las filas vacías que sobran hacia abajo para no imprimir páginas en blanco
      const totalRows = pdfSheet.getMaxRows();
      if (totalRows > rowTotal)
        pdfSheet.hideRows(rowTotal + 1, totalRows - rowTotal);

      SpreadsheetApp.flush();

      // Generamos el PDF basado en la hoja temporal
      const pdfBlob = excelFile.getAs(MimeType.PDF).setName(`${docName}.pdf`);
      pdfFile = folderDestino.createFile(pdfBlob);
      pdfFile.setSharing(
        DriveApp.Access.ANYONE_WITH_LINK,
        DriveApp.Permission.VIEW,
      );
    } catch (e) {
      throw new Error("Paso 7: Falló al generar el PDF. (" + e.message + ")");
    } finally {
      // CRÍTICO: Restauramos el archivo Excel a su estado normal SIEMPRE
      targetSheet.showSheet();
      newDoc.deleteSheet(pdfSheet);
      SpreadsheetApp.flush();
    }

    // --- PASO 8: MOVER EXCEL Y DAR PERMISOS ---
    try {
      excelFile.moveTo(folderDestino);
      excelFile.setSharing(
        DriveApp.Access.ANYONE_WITH_LINK,
        DriveApp.Permission.VIEW,
      );
    } catch (e) {
      throw new Error(
        "Paso 8: Falló al organizar el archivo Excel final. (" +
          e.message +
          ")",
      );
    }

    // --- ÉXITO ---
    return {
      success: true,
      urlExcel: excelFile.getUrl(),
      urlPDF: pdfFile.getUrl(),
    };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Localiza esta función en backend/Controller.gs
 * Se corrigió para que incluya la unidad real de cada producto en el detalle.
 */
/**
 * Obtener detalle de pedido (Con búsqueda inteligente en CRM)
 */
function obtenerDetallePedidoCompleto(idPedido) {
  const sPed = obtenerHojaSegura("PEDIDOS"),
    sDet = obtenerHojaSegura("DETALLE_PEDIDOS"),
    sCli = obtenerHojaSegura("CLIENTES"),
    sProd = obtenerHojaSegura("PRODUCTOS");

  const id = String(idPedido).trim();
  const dP = sPed.getDataRange().getDisplayValues();
  let cab = null, items = [];

  const mapUnidadPorNombre = {};
  if (sProd && sProd.getLastRow() > 1) {
    sProd.getDataRange().getDisplayValues().slice(1).forEach((r) => {
        const nombre = String(r[1] || "").trim().toUpperCase();
        if (nombre) mapUnidadPorNombre[nombre] = _normalizarUnidadLabel(r[3]);
    });
  }

  // 2. Buscar cabecera del pedido
  for (let i = 1; i < dP.length; i++) {
    if (String(dP[i][0]).trim() === id) {
      let r = dP[i];
      let email = "---", empresa = "";
      let telefono = String(r[5] || "").trim(); // Telefono guardado originalmente

      // BÚSQUEDA INTELIGENTE EN EL CRM
      if (sCli) {
        const dC = sCli.getDataRange().getDisplayValues();
        for (let k = 1; k < dC.length; k++) {
          let idCli = String(dC[k][0]).trim();
          let nombreCli = String(dC[k][1]).trim().toLowerCase();
          let empresaCli = String(dC[k][2]).trim().toLowerCase();
          let nombrePedido = String(r[3]).trim().toLowerCase();

          // Magia: Busca coincidencia exacta por ID, o coincidencia por Nombre/Empresa (ideal para los externos)
          if (idCli === String(r[2]).trim() || 
             (nombrePedido !== "" && (nombreCli === nombrePedido || empresaCli === nombrePedido))) {
            
            empresa = dC[k][2];
            email = dC[k][5]; // Rescatamos el correo del CRM
            
            // Si el pedido no tenía teléfono, lo rescatamos del CRM también
            if (!telefono || telefono === "---") telefono = dC[k][4]; 
            break;
          }
        }
      }
      
      cab = {
        id: r[0],
        cliente: r[3],
        direccion: r[4],
        telefono: telefono,
        email: email,
        empresa: empresa,
        paqueteria: r[6],
        guia: r[7],
        costoEnvio: r[9] ? r[9].replace(/[^0-9.]/g, "") : 0,
        estatus: r[10],
        fechaEst: _fmtF(r[11]),
        fechaReal: _fmtF(r[12]),
        comentarios: r[14] || ""
      };
      break;
    }
  }

  if (!cab) throw new Error("Pedido no encontrado");

  // 3. Buscar items e inyectar su unidad real
  if (sDet) {
    const dD = sDet.getDataRange().getDisplayValues();
    for (let i = 1; i < dD.length; i++) {
      if (String(dD[i][0]).trim() === id) {
        let pName = dD[i][1];
        if (pName.includes("Selecciona") || pName === "undefined") pName = "⚠️ Error Datos";
        const unidadFila = _normalizarUnidadLabel(dD[i][6]);
        const unidad = dD[i][6] && String(dD[i][6]).trim() !== "" ? unidadFila : mapUnidadPorNombre[String(pName || "").trim().toUpperCase()] || "L";

        items.push({
          producto: pName,
          presentacion: dD[i][2],
          lote: dD[i][3],
          volumen: dD[i][4],
          piezas: dD[i][5] || 0,
          unidad: unidad,
        });
      }
    }
  }
  return { cabecera: cab, items: items };
}

function actualizarPedido(id, fe, fr, st, guia, comentarios) { // Añadido "comentarios"
  const s = obtenerHojaSegura("PEDIDOS");
  const d = s.getDataRange().getValues();
  for (let i = 1; i < d.length; i++) {
    if (String(d[i][0]).trim() === String(id).trim()) {
      s.getRange(i + 1, 11).setValue(st);
      s.getRange(i + 1, 12).setValue(fe);
      s.getRange(i + 1, 13).setValue(fr);

      if (guia && guia.trim() !== "") {
        s.getRange(i + 1, 8).setValue(guia);
      }
      
      // NUEVO: Guardar comentarios
      if (comentarios !== undefined) {
         s.getRange(i + 1, 15).setValue(comentarios); // Columna O
      }
      return "OK";
    }
  }
}

// ==========================================
// 8. GESTIÓN DE DOCUMENTOS EN DRIVE (PEDIDOS)
// ==========================================

// REEMPLAZA ESTO CON EL ID DE TU CARPETA MAESTRA EN DRIVE
const ID_CARPETA_PADRE_PEDIDOS = "119ZLT4_yRFpkhndI5sPQF6lZCXA0FFUm";

function obtenerOCrearCarpetaPedido(idPedido) {
  const sPed = obtenerHojaSegura("PEDIDOS");
  const d = sPed.getDataRange().getValues();
  let rowIndex = -1;
  let rowData = null;

  for (let i = 1; i < d.length; i++) {
    if (String(d[i][0]).trim() === String(idPedido).trim()) {
      rowIndex = i + 1;
      rowData = d[i];
      break;
    }
  }

  if (rowIndex === -1)
    throw new Error("Pedido no encontrado en la base de datos.");

  // Leer columna N (índice 13) para ver si ya tiene carpeta
  let linkGuardado = rowData[13];
  if (linkGuardado && String(linkGuardado).includes("drive.google.com")) {
    try {
      let folderId = String(linkGuardado).split("/").pop().split("?")[0];
      return DriveApp.getFolderById(folderId);
    } catch (e) {}
  }

  let folderPadre;
  try {
    folderPadre = DriveApp.getFolderById(ID_CARPETA_PADRE_PEDIDOS);
  } catch (e) {
    throw new Error("No se encontró la carpeta principal en Drive.");
  }

  let fechaPed = rowData[1];
  let fechaStr = Utilities.formatDate(
    fechaPed instanceof Date ? fechaPed : new Date(),
    Session.getScriptTimeZone(),
    "dd-MM-yyyy",
  );

  let idCliente = rowData[2];
  let nombreCliente = rowData[3] || "CLIENTE";
  let empresa = "";

  const sCli = obtenerHojaSegura("CLIENTES");
  if (sCli) {
    const dC = sCli.getDataRange().getValues();
    for (let k = 1; k < dC.length; k++) {
      if (dC[k][0] == idCliente && dC[k][2]) {
        empresa = "_" + dC[k][2];
        break;
      }
    }
  }

  let producto = "VARIOS";
  const sDet = obtenerHojaSegura("DETALLE_PEDIDOS");
  if (sDet) {
    const dD = sDet.getDataRange().getValues();
    for (let k = 1; k < dD.length; k++) {
      if (dD[k][0] == idPedido) {
        producto = dD[k][1];
        break;
      }
    }
  }

  let clnNom = String(nombreCliente)
    .replace(/[^a-zA-Z0-9 ]/g, "")
    .trim();
  let clnEmp = String(empresa)
    .replace(/[^a-zA-Z0-9_ ]/g, "")
    .trim();
  let clnProd = String(producto)
    .replace(/[^a-zA-Z0-9 ]/g, "")
    .trim();

  // Nombre Dinámico
  let nombreCarpetaFinal =
    `ENVIO_${clnNom}${clnEmp}_${clnProd}_${fechaStr}`.toUpperCase();

  let folderDestino = folderPadre.createFolder(nombreCarpetaFinal);
  folderDestino.setSharing(
    DriveApp.Access.ANYONE_WITH_LINK,
    DriveApp.Permission.VIEW,
  );

  // Guardar enlace en la Columna N (14)
  sPed.getRange(rowIndex, 14).setValue(folderDestino.getUrl());

  return folderDestino;
}

function subirArchivoPedido(idPedido, nombreArchivo, base64Data, mimeType) {
  try {
    const folder = obtenerOCrearCarpetaPedido(idPedido);
    const blob = Utilities.newBlob(
      Utilities.base64Decode(base64Data),
      mimeType,
      nombreArchivo,
    );
    const file = folder.createFile(blob);
    return {
      success: true,
      url: file.getUrl(),
      nombre: file.getName(),
      urlCarpeta: folder.getUrl(),
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function obtenerArchivosPedido(idPedido) {
  try {
    const sPed = obtenerHojaSegura("PEDIDOS");
    const d = sPed.getDataRange().getValues();
    let linkGuardado = null;

    for (let i = 1; i < d.length; i++) {
      if (String(d[i][0]).trim() === String(idPedido).trim()) {
        linkGuardado = d[i][13]; // Columna N
        break;
      }
    }

    if (!linkGuardado || !String(linkGuardado).includes("drive.google.com")) {
      return { success: true, archivos: [], urlCarpeta: null };
    }

    let folderId = String(linkGuardado).split("/").pop().split("?")[0];
    let folder = DriveApp.getFolderById(folderId);

    const files = folder.getFiles();
    let lista = [];
    while (files.hasNext()) {
      let f = files.next();
      lista.push({ nombre: f.getName(), url: f.getUrl() });
    }

    return { success: true, archivos: lista, urlCarpeta: folder.getUrl() };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ==========================================
// 9. GESTIÓN DE UBICACIONES (EDITAR Y BORRAR)
// ==========================================

function actualizarNombreUbicacion(id, nuevoNombre) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    verificarAccesoServidor();
    const sUbic = obtenerHojaSegura("UBICACIONES");
    const data = sUbic.getDataRange().getValues();

    // Buscar la ubicación por ID y actualizar el nombre
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(id).trim()) {
        sUbic.getRange(i + 1, 2).setValue(nuevoNombre);
        return { success: true };
      }
    }
    throw new Error("Ubicación no encontrada.");
  } catch (e) {
    throw new Error(e.message);
  } finally {
    lock.releaseLock();
  }
}

function borrarUbicacion(id) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    verificarAccesoServidor();

    // 1. VALIDACIÓN DE SEGURIDAD: Comprobar que no tenga inventario
    const sInv = obtenerHojaSegura("INVENTARIO");
    if (sInv && sInv.getLastRow() > 1) {
      const invData = sInv.getDataRange().getValues();
      for (let i = 1; i < invData.length; i++) {
        // Si el producto está en esta ubicación y tiene más de 0 litros, bloqueamos el borrado
        if (
          String(invData[i][2]).trim() === String(id).trim() &&
          Number(invData[i][3]) > 0.001
        ) {
          throw new Error(
            "No se puede borrar: Aún hay productos con stock en esta ubicación.",
          );
        }
      }
    }

    // 2. BORRAR LA UBICACIÓN
    const sUbic = obtenerHojaSegura("UBICACIONES");
    const data = sUbic.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(id).trim()) {
        sUbic.deleteRow(i + 1); // Elimina la fila completa del Excel
        return { success: true };
      }
    }
    throw new Error("Ubicación no encontrada.");
  } catch (e) {
    throw new Error(e.message);
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// TRANSFERENCIA MASIVA
// ==========================================
function transferirMasivo(origenId, destinoId, itemsMover) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    verificarAccesoServidor();
    const sInv = obtenerHojaOCrear("INVENTARIO", []);
    const data = sInv.getDataRange().getValues();
    
    let nuevasFilas = [];

    // Recorremos cada item que seleccionaste en la pantalla
    itemsMover.forEach(item => {
       let restante = Number(item.volumen);
       
       for (let i = 1; i < data.length; i++) {
          if (restante <= 0.001) break;

          if (
             String(data[i][0]) == String(item.productoId) &&
             String(data[i][1]) == String(item.presentacionId) &&
             String(data[i][2]) == String(origenId) &&
             String(data[i][6]) == String(item.lote)
          ) {
             let stockFila = Number(data[i][3]);
             if (stockFila > 0) {
                 let restar = Math.min(stockFila, restante);
                 let nuevoStock = stockFila - restar;
                 
                 // Descontamos del origen
                 sInv.getRange(i + 1, 4).setValue(nuevoStock <= 0.001 ? 0 : nuevoStock);
                 restante -= restar;

                 // Preparamos la fila para insertar en el destino
                 nuevasFilas.push([
                    item.productoId,
                    item.presentacionId,
                    destinoId,
                    Number(restar),
                    data[i][4], // Caducidad
                    data[i][5], // Elaboracion
                    item.lote,
                    new Date(),
                    "Transferencia Masiva"
                 ]);
             }
          }
       }
       
       if (restante > 0.001) {
          throw new Error(`El stock del lote ${item.lote} se acabó antes de completar el movimiento.`);
       }
    });

    // Guardamos todas las nuevas entradas en el destino de un solo golpe
    if (nuevasFilas.length > 0) {
        nuevasFilas.forEach(fila => sInv.appendRow(fila));
    }

    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// MANTENIMIENTO: LIMPIEZA DE INVENTARIO
// ==========================================
function rutinaLimpiezaSemanalCeros() {
  // ATENCIÓN: Forzamos a que limpie la base de datos de PRODUCCIÓN
  // para que el trigger automático no dependa del botón de la interfaz web.
  const DB_PROD = "1zCxn5Cvuvfs29Hbpp58W6VCvV6AczGMG1o7CkhS8d2E"; 
  const db = SpreadsheetApp.openById(DB_PROD);
  const sheet = db.getSheetByName("INVENTARIO");
  
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  let filasBorradas = 0;

  // IMPORTANTE: Recorremos de ABAJO hacia ARRIBA
  for (let i = data.length - 1; i >= 1; i--) {
    let stock = Number(data[i][3]); // La columna D (índice 3) es el Volumen
    
    // Si el stock es 0 (o un decimal minúsculo por error de cálculo)
    if (stock <= 0.001) {
      sheet.deleteRow(i + 1);
      filasBorradas++;
    }
  }
  
  console.log(`🧹 Rutina completada: Se eliminaron ${filasBorradas} registros en 0.`);
}

// ==========================================
// CANCELACIÓN DE PEDIDOS Y LOGÍSTICA INVERSA
// ==========================================
function cancelarPedido(idPedido, tipoCancelacion, idUbicacionDestino, motivo) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    verificarAccesoServidor();
    
    const sPed = obtenerHojaSegura("PEDIDOS");
    const dPed = sPed.getDataRange().getValues();
    let filaPedido = -1;
    let comentariosActuales = "";

    // 1. Buscar el Pedido y actualizar su estatus
    for (let i = 1; i < dPed.length; i++) {
      if (String(dPed[i][0]).trim() === String(idPedido).trim()) {
        filaPedido = i + 1;
        comentariosActuales = dPed[i][14] || ""; // Columna O
        break;
      }
    }
    
    if (filaPedido === -1) throw new Error("Pedido no encontrado");

    // Escribimos el historial automático
    const emailUsuario = Session.getActiveUser().getEmail() || "Sistema";
    const fechaTexto = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
    let notaSistema = `\n\n❌ [${fechaTexto} - ${emailUsuario}]: Pedido CANCELADO. Motivo: ${motivo}. `;
    notaSistema += tipoCancelacion === "retorno" ? `Mercancía enviada a devolución.` : `Mercancía marcada como Merma/Pérdida.`;
    
    sPed.getRange(filaPedido, 11).setValue("Cancelado");
    sPed.getRange(filaPedido, 15).setValue(comentariosActuales + notaSistema);

    // 2. Si es Retorno, devolvemos el stock al Inventario RECUPERANDO LAS FECHAS
    if (tipoCancelacion === "retorno" && idUbicacionDestino) {
        const sSal = obtenerHojaSegura("REGISTROS_SALIDA");
        const sInv = obtenerHojaSegura("INVENTARIO");
        const dSal = sSal.getDataRange().getValues();
        const dInv = sInv.getDataRange().getValues();
        
        // --- MAGIA: MAPEAR LAS FECHAS ORIGINALES DEL INVENTARIO ---
        let mapaFechas = {};
        for(let i = 1; i < dInv.length; i++) {
            let pId = String(dInv[i][0]).trim();
            let lote = String(dInv[i][6]).trim().toUpperCase();
            let key = pId + "|" + lote;
            
            // Guardamos la fecha de elaboración y caducidad asociadas a ese Lote
            if (!mapaFechas[key]) {
                mapaFechas[key] = {
                    caducidad: dInv[i][4],
                    elaboracion: dInv[i][5]
                };
            }
        }
        
        let itemsARetornar = [];
        // Buscamos todo lo que salió con este ID de Pedido
        for (let i = 1; i < dSal.length; i++) {
            if (String(dSal[i][11]).trim() === String(idPedido).trim()) {
                itemsARetornar.push({
                    producto_id: String(dSal[i][0]).trim(),
                    presentacion_id: dSal[i][2],
                    volumen_L: dSal[i][4],
                    lote: String(dSal[i][7]).trim().toUpperCase()
                });
            }
        }
        
        // Lo ingresamos como filas nuevas en el INVENTARIO inyectando las fechas rescatadas
        itemsARetornar.forEach(item => {
            if (Number(item.volumen_L) > 0) {
                
                // Buscamos si el lote existe en el mapa histórico para recuperar sus fechas
                let keyBusqueda = item.producto_id + "|" + item.lote;
                let fechasHistoricas = mapaFechas[keyBusqueda] || { caducidad: "SIN-FECHA", elaboracion: "SIN-FECHA" };

                sInv.appendRow([
                    item.producto_id,
                    item.presentacion_id,
                    idUbicacionDestino,
                    Number(item.volumen_L),
                    fechasHistoricas.caducidad,    // <--- ¡AQUÍ SE RESTAURA LA CADUCIDAD!
                    fechasHistoricas.elaboracion,  // <--- ¡AQUÍ SE RESTAURA LA ELABORACIÓN!
                    item.lote,
                    new Date(),
                    `DEVOLUCIÓN ${idPedido}`
                ]);
            }
        });
    }

    return { success: true };
  } catch (e) {
    throw new Error("Fallo en la cancelación: " + e.message);
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// EJECUTOR DE REORGANIZACIÓN (BIN PACKING)
// ==========================================
function ejecutarReorganizacionBackend() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error("El sistema está ocupado. Intenta en unos segundos.");
  
  try {
    verificarAccesoServidor();

    // --- VERIFICACIÓN DE ROL (Seguridad estricta) ---
    const emailUsuario = Session.getActiveUser().getEmail().toLowerCase();
    const rawSessions = PropertiesService.getScriptProperties().getProperty('ACTIVE_SESSIONS');
    if (rawSessions) {
        let sessions = JSON.parse(rawSessions);
        if (sessions[emailUsuario] && sessions[emailUsuario].rol !== "admin") {
            throw new Error("🔒 SEGURIDAD: Acción bloqueada. Solo un Administrador puede reorganizar el almacén.");
        }
    }
    // ------------------------------------------------

    const sInv = obtenerHojaSegura("INVENTARIO");
    // ... (el resto de tu código sigue igual hacia abajo)
    const sUbic = obtenerHojaSegura("UBICACIONES");
    const sPres = obtenerHojaSegura("PRESENTACIONES");

    const dataInv = sInv.getDataRange().getValues();
    const headersInv = dataInv[0];
    const dataUbic = sUbic.getDataRange().getValues();
    const dataPres = sPres.getDataRange().getValues();

// 1. Mapear presentaciones para saber su capacidad (vNom)
    let mapPres = {};
    for (let i = 1; i < dataPres.length; i++) {
        let id = String(dataPres[i][0]).trim();
        let nombre = String(dataPres[i][1]).toLowerCase();
        let esPieza = nombre.includes("pza") || nombre.includes("unid") || nombre.includes("pieza");
        let vNom = parseFloat(dataPres[i][2]);
        
        if (isNaN(vNom) || vNom <= 0) {
             let m = nombre.match(/[\d\.]+/);
             if (m) {
                 vNom = parseFloat(m[0]);
                 if (nombre.includes("ml") || nombre.includes("gr")) vNom /= 1000;
             }
        }
        if (!vNom || vNom <= 0) vNom = 1;
        
        let pts = 0;
        let esCaja = false;
        
        // REGLA: Si es pieza abstracta, NUNCA va a las cajas. Se queda intacto en la BD.
        if (!esPieza) {
            if (Math.abs(vNom - 1.0) < 0.01) { esCaja = true; pts = 5; }
            else if (Math.abs(vNom - 0.25) < 0.01) { esCaja = true; pts = 1; }
        }
        
        mapPres[id] = { vNom: vNom, pts: pts, esCaja: esCaja };
    }

    // 2. Mapear Ubicaciones existentes
    let mapUbicIds = {};
    for (let i = 1; i < dataUbic.length; i++) {
        mapUbicIds[String(dataUbic[i][1]).trim().toUpperCase()] = String(dataUbic[i][0]).trim();
    }

    let intactos = [];
    let piezasVigentes = [];
    let piezasCaducadas = [];
    const hoy = new Date();
    hoy.setHours(0,0,0,0);

    // 3. Desarmar inventario en piezas físicas
    for (let i = 1; i < dataInv.length; i++) {
        let row = dataInv[i];
        let presId = String(row[1]).trim();
        let infoPres = mapPres[presId];
        
        // Si no es 1L o 0.25L, se queda intacto en su lugar original
        if (!infoPres || !infoPres.esCaja) {
            intactos.push(row);
            continue;
        }
        
        let vol = parseFloat(row[3]);
        let pzasFisicas = Math.floor(vol / infoPres.vNom);
        let restoDecimal = vol - (pzasFisicas * infoPres.vNom);
        
        // Si hay líquido suelto (ej. medias botellas), lo dejamos intacto
        if (restoDecimal > 0.001) {
            let rowResto = [...row];
            rowResto[3] = restoDecimal;
            intactos.push(rowResto);
        }
        
        if (pzasFisicas > 0) {
            let fCadStr = row[4];
            let estaCaducado = false;
            if (fCadStr && fCadStr !== "---" && fCadStr !== "SIN-FECHA") {
                 let partes = String(fCadStr).split('/');
                 if (partes.length === 3) {
                     let d = new Date(partes[2], partes[1]-1, partes[0]);
                     if (d < hoy) estaCaducado = true;
                 }
            }
            
            // Expandimos las botellas conservando su LOTE y FECHAS originales
            for (let p = 0; p < pzasFisicas; p++) {
                let pieza = {
                    prodId: row[0],
                    presId: row[1],
                    vNom: infoPres.vNom,
                    pts: infoPres.pts,
                    caducidad: row[4],
                    elaboracion: row[5],
                    lote: row[6],
                    fecha: row[7],
                    prov: row[8]
                };
                if (estaCaducado) piezasCaducadas.push(pieza);
                else piezasVigentes.push(pieza);
            }
        }
    }

    // 4. Lógica de Bin Packing (Misma regla: 1L primero)
    function empaquetarBD(piezas, prefijoNombre) {
        piezas.sort((a,b) => {
            if (b.pts !== a.pts) return b.pts - a.pts;
            return String(a.prodId).localeCompare(String(b.prodId));
        });
        
        let cajas = [];
        let cajaActual = { nombre: `${prefijoNombre} 1`, puntosUsados: 0, contenido: [] };
        let contador = 1;
        
        for (let p of piezas) {
            if (cajaActual.puntosUsados + p.pts > 60) {
                cajas.push(cajaActual);
                contador++;
                cajaActual = { nombre: `${prefijoNombre} ${contador}`, puntosUsados: 0, contenido: [] };
            }
            cajaActual.contenido.push(p);
            cajaActual.puntosUsados += p.pts;
        }
        if (cajaActual.contenido.length > 0) cajas.push(cajaActual);
        return cajas;
    }

    let todasLasCajas = empaquetarBD(piezasVigentes, "CAJA VIGENTES").concat(empaquetarBD(piezasCaducadas, "CAJA CADUCADOS"));

    // 5. Crear las ubicaciones en la Hoja si no existen
    let nuevasUbicacionesAInsertar = [];
    todasLasCajas.forEach(c => {
        let nombreUpp = c.nombre.toUpperCase();
        if (!mapUbicIds[nombreUpp]) {
            let newId = "UBIC-" + new Date().getTime() + "-" + Math.floor(Math.random()*1000);
            mapUbicIds[nombreUpp] = newId;
            nuevasUbicacionesAInsertar.push([newId, c.nombre]);
        }
        c.idUbicacion = mapUbicIds[nombreUpp];
    });

    if (nuevasUbicacionesAInsertar.length > 0) {
        sUbic.getRange(sUbic.getLastRow() + 1, 1, nuevasUbicacionesAInsertar.length, 2).setValues(nuevasUbicacionesAInsertar);
    }

    // 6. Volver a armar las filas sumando botellas del mismo lote en la misma caja
    let filasEmpacadas = [];
    todasLasCajas.forEach(c => {
        let mapAgrupacion = {};
        c.contenido.forEach(p => {
            let key = p.prodId + "|" + p.presId + "|" + p.lote + "|" + p.caducidad;
            if (!mapAgrupacion[key]) mapAgrupacion[key] = { ...p, pzas: 1 };
            else mapAgrupacion[key].pzas++;
        });
        
        for (let key in mapAgrupacion) {
            let p = mapAgrupacion[key];
            let volumenTotal = p.pzas * p.vNom;
            filasEmpacadas.push([
                p.prodId,
                p.presId,
                c.idUbicacion, 
                volumenTotal,
                p.caducidad,
                p.elaboracion,
                p.lote,
                p.fecha,
                p.prov + " (Auto-Reorg)"
            ]);
        }
    });

    // 7. Sobreescribir Inventario de manera limpia
    let nuevoDataInv = [headersInv].concat(intactos).concat(filasEmpacadas);
    sInv.clearContents();
    sInv.getRange(1, 1, nuevoDataInv.length, nuevoDataInv[0].length).setValues(nuevoDataInv);

    return { success: true, totalCajas: todasLasCajas.length };
    
  } catch (e) {
      throw new Error(e.message);
  } finally {
      lock.releaseLock();
  }
}

// ==========================================
// MÓDULO: HISTORIAL DE ENTRADAS (MEJORADO)
// ==========================================
function obtenerHistorialEntradas() {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    verificarAccesoServidor();
    
    const sheet = obtenerHojaSegura("REGISTROS_ENTRADA"); 
    if (!sheet) return { headers: [], rows: [] };
    
    const data = sheet.getDataRange().getDisplayValues(); 
    if (data.length <= 1) return { headers: data[0] || [], rows: [] };
    
    // --- MAGIA: Diccionarios blindados con toUpperCase() ---
    const mapProd = {}, mapPres = {}, mapUbic = {};
    try {
        obtenerHojaSegura("PRODUCTOS").getDataRange().getValues().forEach(r => mapProd[String(r[0]).trim().toUpperCase()] = r[1]);
        obtenerHojaSegura("PRESENTACIONES").getDataRange().getValues().forEach(r => mapPres[String(r[0]).trim().toUpperCase()] = r[1]);
        obtenerHojaSegura("UBICACIONES").getDataRange().getValues().forEach(r => mapUbic[String(r[0]).trim().toUpperCase()] = r[1]);
    } catch(e) {}

    let headers = ["FECHA", "PRODUCTO", "PRESENTACIÓN", "DESTINO", "CANTIDAD", "LOTE", "PROVEEDOR"];
    let rows = [];
    
    let limite = Math.max(1, data.length - 100);
    for (let i = data.length - 1; i >= limite; i--) {
      let r = data[i];
      
      let fecha = r[0];
      let prod = mapProd[String(r[1]).trim().toUpperCase()] || r[1];
      let pres = mapPres[String(r[2]).trim().toUpperCase()] || r[2];
      
      // Lógica inteligente para Ubicaciones Eliminadas
      let idUbicRaw = String(r[3]).trim();
      let ubic = mapUbic[idUbicRaw.toUpperCase()];
      
      if (!ubic) {
          // Si no existe y parece un ID (es muy largo), le ponemos una etiqueta limpia
          if (idUbicRaw.length > 20 && idUbicRaw.includes("-")) {
              ubic = "Ubic. Eliminada";
          } else {
              ubic = idUbicRaw;
          }
      }

      let cant = r[4]; 
      let lote = r[5];
      let prov = r[6] || "---";

      rows.push([fecha, prod, pres, ubic, cant, lote, prov]);
    }
    
    return { headers: headers, rows: rows };
  } catch (e) {
    throw new Error("Error al leer el historial: " + e.message);
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// MÓDULO: DASHBOARD GERENCIAL (POWER BI STYLE)
// ==========================================
function obtenerEstadisticasDashboard() {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    verificarAccesoServidor();

    const sPed = obtenerHojaSegura("PEDIDOS");
    const sDet = obtenerHojaSegura("DETALLE_PEDIDOS");
    const sInv = obtenerHojaSegura("INVENTARIO");
    const sPres = obtenerHojaSegura("PRESENTACIONES");
    const sProd = obtenerHojaSegura("PRODUCTOS");
    const sCli = obtenerHojaSegura("CLIENTES");

    const hoy = new Date();
    hoy.setHours(0,0,0,0);
    const anioActual = hoy.getFullYear();
    const mesActualStr = `${anioActual}-${String(hoy.getMonth() + 1).padStart(2, '0')}`;
    
    let stats = {
      kpi: {
        gastosMesActual: 0, gastosHistorico: 0, pedidosMesActual: 0,
        enviosEnProceso: 0, enviosCompletados: 0, enviosCancelados: 0,
        litrosActivos: 0, litrosCaducados: 0
      },
      tendenciaAnual: {
        meses: ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"],
        gastos: new Array(12).fill(0), pedidos: new Array(12).fill(0)
      },
      topProductosEnviados: {}, topStock: {}, topEmpresas: {}, geoTipo: {}, geoEstados: {} 
    };

    // MAPEOS
    let mapPres = {};
    if (sPres && sPres.getLastRow() > 1) sPres.getDataRange().getValues().slice(1).forEach(r => mapPres[String(r[0]).trim()] = Number(r[2]) || 0);
    
    let mapProd = {};
    if (sProd && sProd.getLastRow() > 1) {
       sProd.getDataRange().getValues().slice(1).forEach(r => {
           let uLow = String(r[3] || 'L').toLowerCase();
           let uni = 'L';
           if(uLow.includes('kg') || uLow.includes('kilo')) uni = 'Kg';
           else if(uLow.includes('pza') || uLow.includes('unid') || uLow.includes('pieza')) uni = 'Pza';
           mapProd[String(r[0]).trim()] = { nombre: r[1], unidad: uni };
       });
    }

    // --- NUEVA LÓGICA PARA MAPEAR CLIENTES CON SU "TIPO" ---
    let mapEmpresas = {};
    if (sCli && sCli.getLastRow() > 1) {
       sCli.getDataRange().getValues().slice(1).forEach(r => {
           mapEmpresas[String(r[0]).trim()] = {
               nombre: String(r[2]).trim() || String(r[1]).trim(),
               tipo: String(r[3]).trim().toUpperCase() // Columna D (Interno/Externo)
           };
       });
    }

    const obtenerNombrePadre = (rawName) => {
        let match = String(rawName).match(/(.*)\(([^)]+)\)$/);
        return match ? match[2].trim() : String(rawName).trim();
    };

    // 1. INVENTARIO
    if (sInv && sInv.getLastRow() > 1) {
      const dI = sInv.getDataRange().getValues();
      for (let i = 1; i < dI.length; i++) {
        let pId = String(dI[i][0]).trim(); let prId = String(dI[i][1]).trim();
        let stock = Number(dI[i][3]) || 0; let cadVal = dI[i][4];
        
        if (stock > 0.001) {
          let fCad = (cadVal instanceof Date) ? cadVal : new Date(cadVal);
          let estaCaducado = (!isNaN(fCad.getTime()) && fCad < hoy);
          let prodInfo = mapProd[pId] || { nombre: pId, unidad: 'L' };

          if (prodInfo.unidad === 'L') {
              if (estaCaducado) stats.kpi.litrosCaducados += stock;
              else stats.kpi.litrosActivos += stock;
              
              if (!estaCaducado) {
                 let vNom = mapPres[prId] || 0;
                 if (Math.abs(vNom - 1.0) < 0.01 || Math.abs(vNom - 0.25) < 0.01) {
                    let parentName = obtenerNombrePadre(prodInfo.nombre);
                    if (!stats.topStock[parentName]) stats.topStock[parentName] = 0;
                    stats.topStock[parentName] += stock;
                 }
              }
          }
        }
      }
    }

    // 2. PEDIDOS (Muestras, Gastos y Geografía)
    const ESTADOS_MX = ["AGUASCALIENTES", "BAJA CALIFORNIA", "CAMPECHE", "CHIAPAS", "CHIHUAHUA", "COAHUILA", "COLIMA", "DISTRITO FEDERAL", "CIUDAD DE MEXICO", "CDMX", "DURANGO", "GUANAJUATO", "GUERRERO", "HIDALGO", "JALISCO", "ESTADO DE MEXICO", "MICHOACAN", "MORELOS", "NAYARIT", "NUEVO LEON", "OAXACA", "PUEBLA", "QUERETARO", "QUINTANA ROO", "SAN LUIS POTOSI", "SINALOA", "SONORA", "TABASCO", "TAMAULIPAS", "TLAXCALA", "VERACRUZ", "YUCATAN", "ZACATECAS"];
    
    let idPedAceptados = new Set();
    if (sPed && sPed.getLastRow() > 1) {
      const dP = sPed.getDataRange().getValues();
      for (let i = 1; i < dP.length; i++) {
        let idPed = String(dP[i][0]).trim(); if(!idPed) continue;
        let isExterno = idPed.startsWith("EXT-");
        let fVal = dP[i][1]; let f = (fVal instanceof Date) ? fVal : new Date(fVal);
        if(isNaN(f.getTime())) continue;

        let mesStr = `${f.getFullYear()}-${String(f.getMonth() + 1).padStart(2, '0')}`;
        let costo = Number(String(dP[i][9] || "0").replace(/[^0-9.-]+/g, "")) || 0;
        let estatus = String(dP[i][10] || "").toUpperCase();

        // 💰 GASTOS SUMAN SIEMPRE
        stats.kpi.gastosHistorico += costo;
        if(mesStr === mesActualStr) stats.kpi.gastosMesActual += costo;
        if(f.getFullYear() === anioActual) stats.tendenciaAnual.gastos[f.getMonth()] += costo;

        // 📦 PEDIDOS (Solo envíos de material reales)
        if (!isExterno) {
            idPedAceptados.add(idPed);
            if(mesStr === mesActualStr) {
                stats.kpi.pedidosMesActual++;
                if (estatus.includes("ENTREGADO")) stats.kpi.enviosCompletados++;
                else if (estatus.includes("CANCELADO")) stats.kpi.enviosCancelados++;
                else stats.kpi.enviosEnProceso++;
            }
            if(f.getFullYear() === anioActual) stats.tendenciaAnual.pedidos[f.getMonth()]++;

            // --- FILTRO: TOP EMPRESAS SOLO EXTERNAS ---
            let dataCli = mapEmpresas[String(dP[i][2]).trim()];
            let empresaNombre = dataCli ? dataCli.nombre : "Ventas Generales";
            let tipoCli = dataCli ? dataCli.tipo : "EXTERNO"; // Si no hay datos, asumimos externo
            
            // Si NO contiene la palabra INTERNO, entonces sí la contamos en la gráfica
            if (!tipoCli.includes("INTERNO")) {
                if(!stats.topEmpresas[empresaNombre]) stats.topEmpresas[empresaNombre] = 0;
                stats.topEmpresas[empresaNombre]++;
            }

            // GEOGRAFÍA: Tipo y Estado
            let tipoEnvio = String(dP[i][8] || "Nacional").trim();
            if(!stats.geoTipo[tipoEnvio]) stats.geoTipo[tipoEnvio] = 0;
            stats.geoTipo[tipoEnvio]++;

            let direccion = String(dP[i][4] || "").toUpperCase();
            direccion = direccion.replace(/Á/g, "A").replace(/É/g, "E").replace(/Í/g, "I").replace(/Ó/g, "O").replace(/Ú/g, "U");
            let estadoDetectado = "OTRO";
            for(let edo of ESTADOS_MX) {
                if(direccion.includes(edo)) {
                    estadoDetectado = (edo === "DISTRITO FEDERAL" || edo === "CIUDAD DE MEXICO") ? "CDMX" : edo;
                    break;
                }
            }
            if(!stats.geoEstados[estadoDetectado]) stats.geoEstados[estadoDetectado] = 0;
            stats.geoEstados[estadoDetectado]++;
        }
      }
    }

    // 3. DETALLE DE PEDIDOS (Top Muestras)
    if (sDet && sDet.getLastRow() > 1) {
      const dD = sDet.getDataRange().getDisplayValues();
      for (let i = 1; i < dD.length; i++) {
         let idPed = String(dD[i][0]).trim();
         if (!idPedAceptados.has(idPed)) continue; 

         let rawProdName = String(dD[i][1]).trim();
         if(rawProdName) {
            let parentName = obtenerNombrePadre(rawProdName);
            if(!stats.topProductosEnviados[parentName]) stats.topProductosEnviados[parentName] = 0;
            stats.topProductosEnviados[parentName] += 1;
         }
      }
    }

    // ORDENAMIENTO DE TOPS
    let sortObj = (obj) => Object.entries(obj).sort((a,b)=>b[1]-a[1]).slice(0,5).map(e => ({label: e[0], value: e[1]}));
    stats.topProductosArr = sortObj(stats.topProductosEnviados);
    stats.topStockArr = sortObj(stats.topStock);
    stats.topEmpresasArr = sortObj(stats.topEmpresas);
    stats.geoEstadosArr = Object.entries(stats.geoEstados).sort((a,b)=>b[1]-a[1]).slice(0,5).map(e => ({label: e[0], value: e[1]}));
    
    return { success: true, data: stats };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// REGISTRO DE ENVÍOS EXTERNOS (DOCUMENTOS/PLANTA)
// ==========================================
function registrarEnvioExterno(datos) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    verificarAccesoServidor();
    
    const idPedido = "EXT-" + Math.floor(Date.now() / 1000);
    const fechaHoy = new Date();

    // Etiquetamos visualmente la descripción para el historial
    let etiqueta = "📤 ENVÍO EXTERNO";
    let tipoLogistica = "Externo";
    
    if(datos.sentido === "Recepcion") { 
        etiqueta = "📥 RECEPCIÓN EXTERNA"; 
        tipoLogistica = "Recepción"; 
    } else if(datos.sentido === "Triangulacion") { 
        etiqueta = "🔄 TRIANGULACIÓN"; 
        tipoLogistica = "Triangulación"; 
    }

    const sPed = obtenerHojaOCrear("PEDIDOS", ["ID_PEDIDO", "FECHA", "ID_CLIENTE", "NOMBRE", "DIRECCION", "TELEFONO", "PAQUETERIA", "GUIA", "TIPO", "COSTO", "ESTATUS", "F_EST", "F_REAL", "LINK", "COMENTARIOS"]);
    sPed.appendRow([
      idPedido, fechaHoy, "GENERICO-EXT", datos.nombreContacto, datos.direccion, "", datos.paqueteria, datos.guia, tipoLogistica, datos.costo, "Pendiente", "", "", "", etiqueta + ": " + datos.descripcion
    ]);
    return { success: true, idPedido: idPedido };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// MÓDULO: AJUSTES FÍSICOS DE INVENTARIO
// ==========================================
function ajustarStockFisico(datos) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    verificarAccesoServidor();

    // 🔒 SEGURIDAD ESTRICTA BACKEND: Verificar que quien dispara esto sea realmente un Admin
    const emailUsuario = Session.getActiveUser().getEmail().toLowerCase();
    const rawSessions = PropertiesService.getScriptProperties().getProperty('ACTIVE_SESSIONS');
    if (rawSessions) {
        let sessions = JSON.parse(rawSessions);
        if (sessions[emailUsuario] && sessions[emailUsuario].rol !== "admin") {
            throw new Error("🔒 ALERTA DE SEGURIDAD: Acción denegada. Solo un Administrador tiene permisos para hacer ajustes directos de inventario.");
        }
    }

    const sInv = obtenerHojaSegura("INVENTARIO");
    const dataInv = sInv.getDataRange().getValues();
    let filaEncontrada = -1;
    let stockAnterior = 0;

    // Buscar la fila exacta del lote
    for (let i = 1; i < dataInv.length; i++) {
      if (
        String(dataInv[i][0]).trim() === String(datos.productoId).trim() &&
        String(dataInv[i][1]).trim() === String(datos.presentacionId).trim() &&
        String(dataInv[i][2]).trim() === String(datos.ubicacionId).trim() &&
        String(dataInv[i][6]).trim().toUpperCase() === String(datos.lote).trim().toUpperCase()
      ) {
        filaEncontrada = i + 1;
        stockAnterior = Number(dataInv[i][3]);
        break;
      }
    }

    if (filaEncontrada === -1) throw new Error("No se encontró el lote exacto en el inventario.");

    let nuevoStock = Number(datos.nuevoVolumen);
    if (nuevoStock < 0) nuevoStock = 0;

    // 1. Actualizar el Inventario
    sInv.getRange(filaEncontrada, 4).setValue(nuevoStock);

    // 2. Registrar en la Bitácora de Auditoría
    let diferencia = nuevoStock - stockAnterior;
    let tipoMovimiento = diferencia >= 0 ? "AJUSTE POSITIVO (+)" : "MERMA / AJUSTE NEGATIVO (-)";
    let detalle = `Producto: ${datos.productoNombre} | Lote: ${datos.lote} | Cambio: de ${stockAnterior} a ${nuevoStock} | Motivo: ${datos.motivo}`;
    
    registrarEnBitacora(emailUsuario, tipoMovimiento, detalle);

    return { success: true, diferencia: diferencia };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// MÓDULO: KARDEX Y TRAZABILIDAD (V3 - FAMILIAS Y VARIANTES)
// ==========================================
function obtenerKardexProducto(productoId, productoNombre) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    verificarAccesoServidor();
    
    let historial = [];
    const nombreBuscado = String(productoNombre).trim().toUpperCase();

    // Función auxiliar para extraer el "Nombre Padre" ignorando lo que está en paréntesis
    const obtenerNombrePadre = (rawName) => {
        let match = String(rawName).match(/(.*)\(([^)]+)\)$/);
        return match ? match[2].trim().toUpperCase() : String(rawName).trim().toUpperCase();
    };

    // 0. MAPEAR TODOS LOS IDs DE ESTA FAMILIA DE PRODUCTOS
    const sProd = obtenerHojaSegura("PRODUCTOS");
    let idsValidos = new Set();
    idsValidos.add(String(productoId).trim()); // Siempre incluir el ID principal
    
    if (sProd && sProd.getLastRow() > 1) {
        const dP = sProd.getDataRange().getValues();
        for (let i = 1; i < dP.length; i++) {
            let pId = String(dP[i][0]).trim();
            let pName = String(dP[i][1]).trim();
            // Si el nombre base coincide, agregamos este ID a la bolsa de búsqueda
            if (obtenerNombrePadre(pName) === nombreBuscado) {
                idsValidos.add(pId);
            }
        }
    }

    // 1. LEER ENTRADAS (De cualquier ID válido de la familia)
    const sEnt = obtenerHojaSegura("REGISTROS_ENTRADA");
    if (sEnt && sEnt.getLastRow() > 1) {
      const dataEnt = sEnt.getDataRange().getValues();
      for (let i = 1; i < dataEnt.length; i++) {
        let rowProdId = String(dataEnt[i][1]).trim();
        if (idsValidos.has(rowProdId)) { 
          historial.push({
            fecha: dataEnt[i][0], 
            tipo: "ENTRADA",
            cantidad: Number(dataEnt[i][4]) || 0, 
            lote: dataEnt[i][5] || "---",
            detalle: `Proveedor/Origen: ${dataEnt[i][6] || '---'}`
          });
        }
      }
    }

    // 2. LEER SALIDAS Y BAJAS
    const sSal = obtenerHojaSegura("REGISTROS_SALIDA");
    if (sSal && sSal.getLastRow() > 1) {
      const dataSal = sSal.getDataRange().getValues();
      for (let i = 1; i < dataSal.length; i++) {
        let rowProdId = String(dataSal[i][0]).trim();
        if (idsValidos.has(rowProdId)) { 
          let tipo = String(dataSal[i][8]).toUpperCase().includes("DESINCORPORACIÓN") ? "BAJA / MERMA" : "SALIDA";
          historial.push({
            fecha: dataSal[i][10] || dataSal[i][0], 
            tipo: tipo,
            cantidad: -(Number(dataSal[i][4]) || 0),
            lote: dataSal[i][7] || "---",
            detalle: `Destino: ${dataSal[i][8] || '---'} (Doc: ${dataSal[i][11] || '---'})`
          });
        }
      }
    }

    // 3. LEER BITÁCORA (Ajustes Físicos)
    const sBit = obtenerHojaSegura("BITACORA_ACTIVIDAD");
    if (sBit && sBit.getLastRow() > 1) {
      const dataBit = sBit.getDataRange().getValues();
      for (let i = 1; i < dataBit.length; i++) {
        let accion = String(dataBit[i][2] || "").toUpperCase();
        let detalleStr = String(dataBit[i][3] || "");

        if (accion.includes("AJUSTE")) {
          // Extraemos el nombre del producto directamente del texto del log
          let prodMatch = detalleStr.match(/Producto:\s([^|]+)/);
          let prodNameInBitacora = prodMatch ? prodMatch[1].trim() : "";
          
          // Si el "Nombre Padre" del log coincide con el producto que estamos buscando
          if (obtenerNombrePadre(prodNameInBitacora) === nombreBuscado) {
             let loteMatch = detalleStr.match(/Lote:\s([^|]+)/);
             let lote = loteMatch ? loteMatch[1].trim() : "---";

             let cambioMatch = detalleStr.match(/Cambio:\sde\s([\d.]+)\sa\s([\d.]+)/);
             let cantidadDif = 0;
             if (cambioMatch) cantidadDif = Number(cambioMatch[2]) - Number(cambioMatch[1]);

             let motivoMatch = detalleStr.split('| Motivo:');
             let motivoText = motivoMatch[1] ? motivoMatch[1].trim() : "Ajuste Manual";

             historial.push({
               fecha: dataBit[i][0], 
               tipo: accion.includes("POSITIVO") ? "AJUSTE (+)" : "AJUSTE (-)",
               cantidad: cantidadDif,
               lote: lote,
               detalle: `Autorizó: ${dataBit[i][1]} | Razón: ${motivoText}`
             });
          }
        }
      }
    }

    // 4. ORDENAR CRONOLÓGICAMENTE
    historial.sort((a, b) => {
      let dA = new Date(a.fecha).getTime();
      let dB = new Date(b.fecha).getTime();
      return dB - dA;
    });

    // 5. FORMATEAR FECHAS
    let historialLimpio = historial.map(h => {
        let f = new Date(h.fecha);
        let fechaTexto = isNaN(f.getTime()) ? String(h.fecha) : Utilities.formatDate(f, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
        return { fechaStr: fechaTexto, tipo: h.tipo, cantidad: h.cantidad, lote: h.lote, detalle: h.detalle };
    });

    return { success: true, data: historialLimpio };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// MÓDULO: ALERTAS AUTOMÁTICAS (CRON)
// ==========================================
function revisarCaducidadesYEnviarAlerta() {
  const sInv = obtenerHojaSegura("INVENTARIO");
  const sProd = obtenerHojaSegura("PRODUCTOS");
  const sPermisos = obtenerHojaSegura("PERMISOS");

  if (!sInv || sInv.getLastRow() < 2) return;

  const hoy = new Date();
  hoy.setHours(0, 0, 0, 0);
  
  const limite90Dias = new Date();
  limite90Dias.setDate(hoy.getDate() + 90);

  let mapProd = {};
  if (sProd && sProd.getLastRow() > 1) {
    sProd.getDataRange().getValues().slice(1).forEach(r => mapProd[String(r[0]).trim()] = r[1]);
  }

  // AHORA USAMOS OBJETOS PARA AGRUPAR LOTES IDÉNTICOS
  let mapCaducados = {};
  let mapProximos = {};

  const dataInv = sInv.getDataRange().getValues();
  
  for (let i = 1; i < dataInv.length; i++) {
    let pId = String(dataInv[i][0]).trim();
    let stock = Number(dataInv[i][3]) || 0;
    let cadVal = dataInv[i][4];
    let lote = String(dataInv[i][6]).trim();

    if (stock > 0.001 && cadVal && cadVal !== "---" && cadVal !== "SIN-FECHA") {
      let fCad = null;
      if (cadVal instanceof Date) {
        fCad = cadVal;
      } else if (typeof cadVal === "string") {
        if (cadVal.includes("-")) {
          let p = cadVal.split("T")[0].split("-");
          if (p.length === 3) fCad = new Date(p[0], p[1] - 1, p[2]);
        } else if (cadVal.includes("/")) {
          let p = cadVal.split("/");
          if (p.length === 3) fCad = new Date(p[2], p[1] - 1, p[0]);
        }
      }

      if (fCad && !isNaN(fCad.getTime())) {
        let nombreProd = mapProd[pId] || pId;
        let fechaStr = Utilities.formatDate(fCad, Session.getScriptTimeZone(), "dd/MM/yyyy");
        
        // Creamos una "llave única" combinando Nombre + Lote + Fecha
        let key = `${nombreProd}|${lote}|${fechaStr}`;

        if (fCad < hoy) {
          if (!mapCaducados[key]) mapCaducados[key] = { nombre: nombreProd, lote: lote, stock: 0, vence: fechaStr };
          mapCaducados[key].stock += stock; // Sumamos el stock si ya existía
        } else if (fCad <= limite90Dias) {
          if (!mapProximos[key]) mapProximos[key] = { nombre: nombreProd, lote: lote, stock: 0, vence: fechaStr };
          mapProximos[key].stock += stock; // Sumamos el stock si ya existía
        }
      }
    }
  }

  // CONVERTIMOS LOS GRUPOS EN LÍNEAS DE HTML PARA EL CORREO
  let caducados = Object.values(mapCaducados).map(item => 
    `<li><b>${item.nombre}</b> (Lote: <span style="font-family:monospace;">${item.lote}</span>) - Stock Total: ${item.stock.toFixed(2)} - Vence: ${item.vence}</li>`
  );

  let proximos = Object.values(mapProximos).map(item => 
    `<li><b>${item.nombre}</b> (Lote: <span style="font-family:monospace;">${item.lote}</span>) - Stock Total: ${item.stock.toFixed(2)} - Vence: ${item.vence}</li>`
  );

  // SI NO HAY ALERTAS, NO MANDAMOS CORREO
  if (caducados.length === 0 && proximos.length === 0) return;

  // OBTENER CORREOS
  let correosDestino = [];
  if (sPermisos && sPermisos.getLastRow() > 1) {
    const dataPerm = sPermisos.getDataRange().getValues();
    for(let i = 1; i < dataPerm.length; i++) {
       let esAdmin = dataPerm[i][8] === true || String(dataPerm[i][8]).toUpperCase() === 'TRUE';
       let correo = String(dataPerm[i][0]).trim();
       if (esAdmin && correo.includes("@")) correosDestino.push(correo);
    }
  }
  
  if (correosDestino.length === 0) correosDestino.push(Session.getActiveUser().getEmail());

  // ARMAR EL CORREO HTML
  let htmlBody = `
    <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; border: 1px solid #e2e8f0; border-radius: 10px; overflow: hidden;">
      <div style="background-color: #1e293b; padding: 20px; text-align: center; color: white;">
        <h2 style="margin: 0;">🚨 Alerta de Inventario WMS</h2>
        <p style="margin: 5px 0 0 0; font-size: 0.9rem; opacity: 0.8;">Reporte automático del estado de caducidades</p>
      </div>
      <div style="padding: 20px; background-color: #f8fafc; color: #334155;">
        <p>Hola Administrador,</p>
        <p>El sistema ha detectado lotes que requieren tu atención inmediata para evitar mermas financieras.</p>
  `;

  if (caducados.length > 0) {
    htmlBody += `
        <h3 style="color: #dc2626; border-bottom: 2px solid #fca5a5; padding-bottom: 5px;">🛑 Productos Caducados (${caducados.length})</h3>
        <ul style="padding-left: 20px; line-height: 1.6;">
          ${caducados.join("")}
        </ul>
        <p style="font-size: 0.85rem; color: #64748b;"><em>* Sugerencia: Entra al sistema, ve a la pestaña "Bajas" y desincorpora estos lotes.</em></p>
    `;
  }

  if (proximos.length > 0) {
    htmlBody += `
        <h3 style="color: #f59e0b; border-bottom: 2px solid #fcd34d; padding-bottom: 5px; margin-top: 25px;">⚠️ Próximos a Caducar (< 90 días) (${proximos.length})</h3>
        <ul style="padding-left: 20px; line-height: 1.6;">
          ${proximos.join("")}
        </ul>
        <p style="font-size: 0.85rem; color: #64748b;"><em>* Sugerencia: Prioriza la salida de estos lotes en tus próximos envíos FIFO.</em></p>
    `;
  }

  htmlBody += `
      </div>
      <div style="background-color: #e2e8f0; padding: 15px; text-align: center; font-size: 0.8rem; color: #64748b;">
        Este es un mensaje automático generado por tu Sistema de Gestión de Inventario.<br>
        No es necesario responder a este correo.
      </div>
    </div>
  `;

  // ENVIAR EL CORREO
  MailApp.sendEmail({
    to: correosDestino.join(","),
    subject: "🚨 Alerta WMS: Tienes productos caducados o por caducar",
    htmlBody: htmlBody
  });
}


// ==========================================
// MÓDULO: DIRECTORIO DE CLIENTES Y DIRECCIONES
// ==========================================

function obtenerDirectorioClientes() {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    verificarAccesoServidor();
    const sCli = obtenerHojaSegura("CLIENTES");
    if (!sCli || sCli.getLastRow() < 2) return { success: true, data: [] };

    const data = sCli.getDataRange().getDisplayValues();
    let clientes = [];
    
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue; // Si no hay ID, saltar
      
      let dirs = [];
      try {
         // Intentamos leer el JSON de la columna G (índice 6)
         dirs = JSON.parse(data[i][6] || "[]");
      } catch(e) {
         // Si era texto normal viejo, lo metemos como un elemento de lista
         if (data[i][6]) dirs = [data[i][6]]; 
      }

      clientes.push({
        id: String(data[i][0]).trim(),
        nombre: String(data[i][1]).trim(),
        empresa: String(data[i][2]).trim(),
        tipo: String(data[i][3]).trim() || "Externo",
        telefono: String(data[i][4]).trim(),
        correo: String(data[i][5]).trim(),
        direcciones: Array.isArray(dirs) ? dirs : []
      });
    }
    
    // Ordenar alfabéticamente por empresa
    clientes.sort((a, b) => a.empresa.localeCompare(b.empresa));
    
    return { success: true, data: clientes };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}


// 2. PEGA ESTAS 3 FUNCIONES NUEVAS AL FINAL DEL ARCHIVO Controller.gs:
function guardarClienteCRM(datos) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    verificarAccesoServidor();
    const sCli = obtenerHojaSegura("CLIENTES");
    const data = sCli.getDataRange().getValues();

    let fila = -1;
    if (datos.id && datos.id.trim() !== "") {
        for (let i = 1; i < data.length; i++) {
            if (String(data[i][0]).trim() === String(datos.id).trim()) {
                fila = i + 1;
                break;
            }
        }
    }

    if (fila > 0) {
        sCli.getRange(fila, 2).setValue(datos.nombre);
        sCli.getRange(fila, 3).setValue(datos.empresa);
        sCli.getRange(fila, 4).setValue(datos.tipo);
        sCli.getRange(fila, 5).setValue(datos.telefono);
        sCli.getRange(fila, 6).setValue(datos.email);
        return { success: true, id: datos.id };
    } else {
        let nuevoId = "CLI-" + Math.floor(Date.now() / 1000);
        sCli.appendRow([nuevoId, datos.nombre, datos.empresa, datos.tipo, datos.telefono, datos.email, "[]"]);
        return { success: true, id: nuevoId };
    }
  } catch(e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

function agregarDireccionCliente(idCliente, nuevaDireccion) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    verificarAccesoServidor();
    const sCli = obtenerHojaSegura("CLIENTES");
    const data = sCli.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(idCliente).trim()) {
        let dirs = [];
        try { dirs = JSON.parse(data[i][6] || "[]"); } catch(e) {}
        
        if (!dirs.includes(nuevaDireccion)) {
            dirs.push(nuevaDireccion);
            sCli.getRange(i + 1, 7).setValue(JSON.stringify(dirs)); 
        }
        return { success: true, direcciones: dirs };
      }
    }
    throw new Error("Cliente no encontrado.");
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

function actualizarDireccionesCliente(idCliente, arrayDirecciones) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    verificarAccesoServidor();
    const sCli = obtenerHojaSegura("CLIENTES");
    const data = sCli.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(idCliente).trim()) {
        sCli.getRange(i + 1, 7).setValue(JSON.stringify(arrayDirecciones));
        return { success: true, direcciones: arrayDirecciones };
      }
    }
    throw new Error("Cliente no encontrado.");
  } catch(e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}