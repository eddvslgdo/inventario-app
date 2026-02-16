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
const DB_SPREADSHEET_ID = '1zCxn5Cvuvfs29Hbpp58W6VCvV6AczGMG1o7CkhS8d2E';

function obtenerSpreadsheet() {
  try {
    // Intentamos abrir por ID explícito
    return SpreadsheetApp.openById(DB_SPREADSHEET_ID);
  } catch (e) {
    // Fallback de emergencia
    return SpreadsheetApp.getActiveSpreadsheet();
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
    return Utilities.formatDate(valor, Session.getScriptTimeZone(), "dd/MM/yyyy");
  }
  return String(valor).trim();
}

// ==========================================
// 1. CATÁLOGOS
// ==========================================
function obtenerCatalogos() {
  const sProd = obtenerHojaOCrear('PRODUCTOS', ['ID', 'NOMBRE', 'DESCRIPCION', 'UNIDAD']);
  const sPres = obtenerHojaOCrear('PRESENTACIONES', ['ID', 'NOMBRE', 'VOLUMEN']);
  const sUbic = obtenerHojaOCrear('UBICACIONES', ['ID', 'NOMBRE']);

  const leer = (s, esPres) => {
    if (!s || s.getLastRow() < 2) return [];
    return s.getDataRange().getValues().slice(1).map(r => ({
      id: String(r[0]).trim(), 
      nombre: r[1], 
      volumen: esPres ? (Number(r[2]) || 0) : 0
    })).filter(i => i.id);
  };

  return { productos: leer(sProd), presentaciones: leer(sPres, true), ubicaciones: leer(sUbic) };
}

function obtenerListaClientes() {
  const s = obtenerHojaOCrear('CLIENTES', ['ID', 'NOMBRE', 'EMPRESA', 'DIRECCION', 'TELEFONO', 'EMAIL']);
  if (!s || s.getLastRow() < 2) return [];
  return s.getDataRange().getValues().slice(1).map(r => ({
    id: r[0], nombre: r[1], empresa: r[2], direccion: r[3], telefono: r[4], email: r[5]
  })).filter(i => i.id);
}

// ==========================================
// 2. GESTIÓN DE INVENTARIO (CONSOLIDACIÓN)
// ==========================================
function obtenerDatosUbicaciones() {
  const sInv = obtenerHojaOCrear('INVENTARIO', ['ID_PROD', 'ID_PRES', 'ID_UBIC', 'STOCK', 'CADUCIDAD', 'ELABORACION', 'LOTE', 'F_ENTRADA', 'PROVEEDOR']);
  const sUbic = obtenerHojaOCrear('UBICACIONES', []);
  const sProd = obtenerHojaOCrear('PRODUCTOS', []);
  const sPres = obtenerHojaOCrear('PRESENTACIONES', []);

  const mapProd = {}, mapPres = {}, mapPresVol = {};
  if (sProd.getLastRow() > 1) sProd.getDataRange().getValues().slice(1).forEach(r => mapProd[String(r[0]).trim()] = r[1]);
  if (sPres.getLastRow() > 1) sPres.getDataRange().getValues().slice(1).forEach(r => {
    const id = String(r[0]).trim(); mapPres[id] = r[1]; mapPresVol[id] = Number(r[2]) || 0;
  });

  let ubicaciones = [];
  if (sUbic.getLastRow() > 1) {
    sUbic.getDataRange().getValues().slice(1).forEach(r => {
      if (r[0]) ubicaciones.push({ id: String(r[0]).trim(), nombre: r[1] || 'S/N', items: [], totalVolumen: 0 });
    });
  }

  if (sInv.getLastRow() > 1) {
    const dInv = sInv.getDataRange().getValues();
    for (let i = 1; i < dInv.length; i++) {
      const stock = Number(dInv[i][3]);
      if (stock > 0.001) {
        const uId = String(dInv[i][2]).trim();
        let ubic = ubicaciones.find(u => u.id === uId);
        if (!ubic) { ubic = { id: uId, nombre: "Ubic: " + uId, items: [], totalVolumen: 0 }; ubicaciones.push(ubic); }

        const pId = String(dInv[i][0]).trim();
        const prId = String(dInv[i][1]).trim();
        const lote = String(dInv[i][6]).trim();
        const caducidadStr = _fmtFechaDisplay(dInv[i][4]);

        const nProd = mapProd[pId] || pId;
        const nPres = mapPres[prId] || prId;

        let itemExistente = ubic.items.find(it => 
            it.raw_producto_id === pId && 
            it.raw_presentacion_id === prId && 
            it.lote === lote && 
            it.caducidad === caducidadStr
        );

        if (itemExistente) {
          itemExistente.volumen += stock;
        } else {
          ubic.items.push({
            raw_producto_id: pId, producto: nProd, raw_presentacion_id: prId, presentacion: nPres,
            lote: lote, volumen: stock, volumen_nominal: mapPresVol[prId] || 0,
            caducidad: caducidadStr, nombre_completo: `${nProd} (${nPres})`
          });
        }
        ubic.totalVolumen += stock;
      }
    }
  }
  return ubicaciones;
}

function obtenerDatosProductos() {
  const sInv = obtenerHojaOCrear('INVENTARIO', []);
  if (sInv.getLastRow() < 2) return [];

  const mapProd = {}, mapUbic = {}, mapPres = {};
  const sP = obtenerHojaOCrear('PRODUCTOS', []), sU = obtenerHojaOCrear('UBICACIONES', []), sPr = obtenerHojaOCrear('PRESENTACIONES', []);
  if (sP.getLastRow() > 1) sP.getDataRange().getValues().slice(1).forEach(r => mapProd[String(r[0]).trim()] = r[1]);
  if (sU.getLastRow() > 1) sU.getDataRange().getValues().slice(1).forEach(r => mapUbic[String(r[0]).trim()] = r[1]);
  if (sPr.getLastRow() > 1) sPr.getDataRange().getValues().slice(1).forEach(r => mapPres[String(r[0]).trim()] = r[1]);

  const dataInv = sInv.getDataRange().getValues();
  let productosMap = {};

  for (let i = 1; i < dataInv.length; i++) {
    const stock = Number(dataInv[i][3]);
    if (stock > 0.001) {
      const pId = String(dataInv[i][0]).trim();
      if (!productosMap[pId]) productosMap[pId] = { id: pId, nombre: mapProd[pId] || pId, totalVolumen: 0, lotes: [] };
      
      productosMap[pId].totalVolumen += stock;
      
      const uId = String(dataInv[i][2]).trim();
      const prId = String(dataInv[i][1]).trim();
      const lote = String(dataInv[i][6]).trim();
      const uName = mapUbic[uId] || uId;
      const presName = mapPres[prId] || prId;
      const cadStr = _fmtFechaDisplay(dataInv[i][4]);
      
      let loteExistente = productosMap[pId].lotes.find(l => 
          l.lote === lote && 
          l.ubicacion === uName && 
          l.presentacion === presName && 
          l.caducidad === cadStr
      );
      
      if(loteExistente) { loteExistente.volumen += stock; }
      else {
        productosMap[pId].lotes.push({
          lote: lote, volumen: stock, ubicacion: uName,
          presentacion: presName, caducidad: cadStr
        });
      }
    }
  }
  return Object.values(productosMap);
}

// ==========================================
// 3. REGISTRO DE MOVIMIENTOS
// ==========================================
function registrarNuevoProducto(d) { obtenerHojaOCrear('PRODUCTOS').appendRow([Utilities.getUuid(), d.nombre, d.descripcion, d.unidad]); return true; }
function registrarNuevaPresentacion(n) { 
  let v=0, m=String(n).match(/[\d\.]+/); if(m){ v=parseFloat(m[0]); if(String(n).toLowerCase().includes('ml')) v/=1000; }
  obtenerHojaOCrear('PRESENTACIONES').appendRow([Utilities.getUuid(), n, v]); return true; 
}
function registrarNuevaUbicacion(n) { obtenerHojaOCrear('UBICACIONES').appendRow([Utilities.getUuid(), n]); return true; }

function registrarEntradaUnica(datos) { return registrarEntradaMasiva([datos]); }

function registrarEntradaMasiva(lista) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const sInv = obtenerHojaOCrear('INVENTARIO', ['ID_PROD', 'ID_PRES', 'ID_UBIC', 'STOCK', 'CADUCIDAD', 'ELABORACION', 'LOTE', 'F_ENTRADA', 'PROVEEDOR']);
    const sLog = obtenerHojaOCrear('REGISTROS_ENTRADA', ['FECHA', 'PROD', 'PRES', 'UBIC', 'VOL', 'LOTE', 'PROVEEDOR', 'TIPO']);
    
    lista.forEach(d => {
      if(d.producto_id.includes("Selecciona") || d.producto_id === "undefined") throw new Error("Producto inválido.");

      let fElab = d.fecha_elaboracion;
      let fCad = d.fecha_caducidad;

      if (!fElab) {
         let now = new Date();
         fElab = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd");
      }

      if (!fCad || fCad === "") {
         let partes = String(fElab).split('-');
         if (partes.length === 3) {
            let anio = parseInt(partes[0]) + 2; 
            fCad = `${anio}-${partes[1]}-${partes[2]}`;
         } else {
            fCad = "SIN-FECHA"; 
         }
      }

      sInv.appendRow([
          d.producto_id, d.presentacion_id, d.ubicacion_id, Number(d.volumen_L), 
          fCad, fElab, d.lote, new Date(), d.proveedor
      ]);
      
      sLog.appendRow([new Date(), d.producto_id, d.presentacion_id, d.ubicacion_id, d.volumen_L, d.lote, d.proveedor, 'Entrada']);
    });
    return true;
  } catch(e){ throw e; } finally { lock.releaseLock(); }
}

// ==========================================
// 4. MOVER Y TRANSFORMAR
// ==========================================
function transferirProducto(origenId, destinoId, productoId, presentacionId, lote, cantidadMover) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sInv = obtenerHojaOCrear('INVENTARIO', []);
    const data = sInv.getDataRange().getValues();
    
    let filasOrigen = [], stockTotalDisp = 0;
    for(let i=1; i<data.length; i++) {
      if(String(data[i][0])==String(productoId) && String(data[i][1])==String(presentacionId) && 
         String(data[i][2])==String(origenId) && String(data[i][6])==String(lote)) {
         filasOrigen.push({index: i+1, stock: Number(data[i][3]), rowData: data[i]});
         stockTotalDisp += Number(data[i][3]);
      }
    }

    if(filasOrigen.length === 0) throw new Error("Origen no encontrado");
    if(stockTotalDisp < cantidadMover - 0.001) throw new Error("Stock insuficiente");

    let restante = cantidadMover;
    for(let item of filasOrigen) {
        if(restante <= 0) break;
        let restar = Math.min(item.stock, restante);
        let ns = item.stock - restar;
        sInv.getRange(item.index, 4).setValue(ns <= 0.001 ? 0 : ns);
        restante -= restar;
    }

    sInv.appendRow([productoId, presentacionId, destinoId, Number(cantidadMover), filasOrigen[0].rowData[4], filasOrigen[0].rowData[5], lote, new Date(), "Transferencia"]);
    return {success:true};
  } catch(e) { return {success:false, error:e.message}; } finally { lock.releaseLock(); }
}

function realizarTransformacion(datos) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sInv = obtenerHojaOCrear('INVENTARIO', []);
    const data = sInv.getDataRange().getValues();
    
    let filasOrigen = [], stockTotalDisp = 0;
    for(let i=1; i<data.length; i++) {
      if(String(data[i][2])==String(datos.origenId) && String(data[i][0])==String(datos.productoIdOrigen) && 
         String(data[i][6]).toUpperCase()==String(datos.loteOrigen).toUpperCase()) {
         filasOrigen.push({index: i+1, stock: Number(data[i][3]), rowData: data[i]});
         stockTotalDisp += Number(data[i][3]);
      }
    }

    if(filasOrigen.length === 0) throw new Error("Lote origen no encontrado");
    if(stockTotalDisp < Number(datos.cantidadLitros)) throw new Error("Stock insuficiente");

    let restante = Number(datos.cantidadLitros);
    for(let item of filasOrigen) {
        if(restante <= 0) break;
        let restar = Math.min(item.stock, restante);
        let ns = item.stock - restar;
        sInv.getRange(item.index, 4).setValue(ns <= 0.001 ? 0 : ns);
        restante -= restar;
    }

    const lFin = datos.nuevoLote || datos.loteOrigen;
    sInv.appendRow([datos.nuevoProductoId, datos.nuevaPresentacionId, datos.origenId, Number(datos.cantidadLitros), 
                    filasOrigen[0].rowData[4], filasOrigen[0].rowData[5], lFin, new Date(), "Transformación"]);
    return {success:true};
  } catch(e) { return {success:false, error:e.message}; } finally { lock.releaseLock(); }
}

// ==========================================
// 5. PROCESAMIENTO DE SALIDAS (PEDIDOS)
// ==========================================
function registrarSalidas(lista) {
  return procesarPedidoCompleto({
      idCliente: 'GENERICO', nombreCliente: 'Salida Rápida', direccion:'-', telefono:'-', email:'-', tipoEnvio:'Nacional', paqueteria:'Mostrador', guia:'-', costoEnvio:0
  }, lista, false);
}

function procesarPedidoCompleto(datosPedido, itemsCarrito, guardarCliente) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    itemsCarrito.forEach(item => {
      const p = String(item.producto_id);
      if (p.includes("Selecciona") || p === "undefined" || !p) throw new Error("⚠️ Producto inválido.");
      if (!item.lote || item.lote === "undefined") throw new Error("⚠️ Lote inválido.");
    });

    const unicos = {};
    itemsCarrito.forEach(i => {
      const k = `${i.producto_id}|${i.presentacion_id}|${i.lote}|${i.ubicacion_id}`;
      if(unicos[k]) { unicos[k].volumen_L += Number(i.volumen_L); unicos[k].piezas += Number(i.piezas); }
      else { unicos[k] = {...i, volumen_L: Number(i.volumen_L), piezas: Number(i.piezas)}; }
    });
    const itemsProcesar = Object.values(unicos);

    const idPedido = "PED-" + Math.floor(Date.now() / 1000);
    const fechaHoy = new Date();

    if (guardarCliente) {
      const sCli = obtenerHojaOCrear('CLIENTES', ['ID', 'NOMBRE', 'EMPRESA', 'DIRECCION', 'TELEFONO', 'EMAIL']);
      let idCli = datosPedido.idCliente;
      if (!idCli || idCli === 'nuevo') idCli = "CLI-" + Math.floor(Math.random()*10000);
      sCli.appendRow([idCli, datosPedido.nombreCliente, datosPedido.empresa, datosPedido.direccion, datosPedido.telefono, datosPedido.email]);
    }

    const sPed = obtenerHojaOCrear('PEDIDOS', ['ID_PEDIDO', 'FECHA', 'ID_CLIENTE', 'NOMBRE', 'DIRECCION', 'TELEFONO', 'PAQUETERIA', 'GUIA', 'TIPO', 'COSTO', 'ESTATUS', 'F_EST', 'F_REAL', 'LINK']);
    sPed.appendRow([idPedido, fechaHoy, datosPedido.idCliente, datosPedido.nombreCliente, datosPedido.direccion, datosPedido.telefono, datosPedido.paqueteria, datosPedido.guia, datosPedido.tipoEnvio, datosPedido.costoEnvio, 'Pendiente', '', '', '']);

    // <--- CAMBIO 1: Agregamos 'PIEZAS' a los encabezados de la hoja DETALLE_PEDIDOS
    const sInv = obtenerHojaOCrear('INVENTARIO', []), sSal = obtenerHojaOCrear('REGISTROS_SALIDA', []), sDet = obtenerHojaOCrear('DETALLE_PEDIDOS', ['ID_PEDIDO', 'PRODUCTO', 'PRESENTACION', 'LOTE', 'VOLUMEN', 'PIEZAS']);
    const dataInv = sInv.getDataRange().getValues();

    itemsProcesar.forEach(item => {
      // <--- CAMBIO 2: Agregamos Number(item.piezas || 0) a la fila que se va a guardar
      sDet.appendRow([idPedido, item.nombre_producto, item.nombre_presentacion || '---', item.lote, Number(item.volumen_L), Number(item.piezas || 0)]);

      let restante = item.volumen_L;
      for (let i = 1; i < dataInv.length; i++) {
        if(restante <= 0.001) break;
        const invProd = String(dataInv[i][0]).trim();
        const invLote = String(dataInv[i][6]).trim().toUpperCase();
        const invUbic = String(dataInv[i][2]).trim(); 
        
        if (invProd === String(item.producto_id).trim() && invLote === String(item.lote).trim().toUpperCase() && invUbic === String(item.ubicacion_id).trim()) {
          const stock = Number(dataInv[i][3]);
          if (stock > 0) {
             const restar = Math.min(stock, restante);
             sInv.getRange(i + 1, 4).setValue(stock - restar);
             restante -= restar;
          }
        }
      }
      sSal.appendRow([item.producto_id, item.nombre_producto, item.presentacion_id, item.nombre_presentacion, item.volumen_L, item.piezas, item.ubicacion_id, item.lote, datosPedido.nombreCliente, Session.getActiveUser().getEmail(), fechaHoy, idPedido]);
    });

    return { success: true, idPedido: idPedido };
  } catch (e) { return { success: false, error: e.message }; } finally { lock.releaseLock(); }
}

// ==========================================
// 6. HISTORIAL DE ENVÍOS
// ==========================================
function obtenerHistorialPedidos() {
  const sPed = obtenerHojaSegura('PEDIDOS');
  if (!sPed) return [];
  const data = sPed.getDataRange().getDisplayValues();
  if (data.length < 2) return [];

  const sDet = obtenerHojaSegura('DETALLE_PEDIDOS'), sCli = obtenerHojaSegura('CLIENTES');
  const mapProd = {}, mapEmp = {};
  
  if(sDet) {
    const dDet = sDet.getDataRange().getDisplayValues();
    for(let i=1; i<dDet.length; i++) {
      let id = String(dDet[i][0]).trim();
      if(id) {
        if(!mapProd[id]) mapProd[id] = [];
        mapProd[id].push(`${dDet[i][1]} (${dDet[i][4]}L)`);
      }
    }
  }
  if(sCli) sCli.getDataRange().getDisplayValues().forEach((r,i) => { if(i>0) mapEmp[String(r[0]).trim()] = r[2]; });

  let pedidos = [];
  for(let i = data.length - 1; i >= 1; i--) {
    let r = data[i], id = String(r[0]).trim();
    if(id) {
      let prods = mapProd[id] || [], resumen = prods.length > 0 ? prods.join(", ") : "---";
      if(resumen.length > 55) resumen = resumen.substring(0, 55) + "...";
      pedidos.push({
        id: id, fecha: r[1], cliente: r[3], empresa: mapEmp[String(r[2]).trim()] || "",
        resumenProductos: resumen, logistica: r[6], guia: r[7], estatus: r[10] || "Pendiente",
        costo: r[9] ? String(r[9]).replace(/[^0-9.]/g, '') : 0
      });
    }
  }
  return pedidos;
}

function obtenerDetallePedidoCompleto(idPedido) {
  const sPed = obtenerHojaSegura('PEDIDOS'), sDet = obtenerHojaSegura('DETALLE_PEDIDOS'), sCli = obtenerHojaSegura('CLIENTES');
  const id = String(idPedido).trim();
  const dP = sPed.getDataRange().getDisplayValues();
  let cab = null, items = [];

  for(let i=1; i<dP.length; i++) {
    if(String(dP[i][0]).trim() === id) {
      let r = dP[i], email = "---";
      if(sCli) {
        const dC = sCli.getDataRange().getDisplayValues();
        for(let k=1; k<dC.length; k++) if(String(dC[k][0]).trim() === String(r[2]).trim()) { email = dC[k][5]; break; }
      }
      cab = {
        id: r[0], cliente: r[3], direccion: r[4], telefono: r[5], email: email,
        paqueteria: r[6], guia: r[7], costoEnvio: r[9] ? r[9].replace(/[^0-9.]/g, '') : 0,
        estatus: r[10], fechaEst: _fmtF(r[12]), fechaReal: _fmtF(r[13])
      };
      break;
    }
  }
  if(!cab) throw new Error("Pedido no encontrado");

  if(sDet) {
    const dD = sDet.getDataRange().getDisplayValues();
    for(let i=1; i<dD.length; i++) {
      if(String(dD[i][0]).trim() === id) {
        let pName = dD[i][1];
        if(pName.includes("Selecciona") || pName === "undefined") pName = "⚠️ Error Datos";
        items.push({ producto: pName, presentacion: dD[i][2], lote: dD[i][3], volumen: dD[i][4], piezas: dD[i][5] || 0 });
      }
    }
  }
  return { cabecera: cab, items: items };
}

function actualizarPedido(id, fe, fr, st, guia) { 
  const s = obtenerHojaSegura('PEDIDOS');
  const d = s.getDataRange().getValues();
  for(let i=1; i<d.length; i++) {
    if(String(d[i][0]).trim() === String(id).trim()) {
      
      // 1. Actualizar Estatus y Fechas (Código original)
      s.getRange(i+1, 11).setValue(st);
      s.getRange(i+1, 13).setValue(fe); 
      s.getRange(i+1, 14).setValue(fr); 
      
      // 2. --- NUEVO: Actualizar Guía si se envió ---
      if (guia && guia.trim() !== "") {
          s.getRange(i+1, 8).setValue(guia); // Columna H es la 8
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
    const sheetInv = obtenerHojaSegura('INVENTARIO');
    if(!sheetInv || sheetInv.getLastRow() < 2) return JSON.stringify({ success: false, error: "Sin stock" });
    
    const data = sheetInv.getDataRange().getValues();
    
    // 1. Obtener Nombres
    const mapUbic = {}, mapPres = {};
    const sU = obtenerHojaSegura('UBICACIONES'), sP = obtenerHojaSegura('PRESENTACIONES');
    if(sU) sU.getDataRange().getValues().slice(1).forEach(r => mapUbic[String(r[0]).trim()] = r[1]);
    if(sP) sP.getDataRange().getValues().slice(1).forEach(r => mapPres[String(r[0]).trim()] = r[1]);

    let lotes = [];
    
    // 2. Leer y Agrupar
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === prodIdBuscado) {
        const uId = String(data[i][2]).trim(); 
        const presId = String(data[i][1]).trim();
        const lote = String(data[i][6]).trim();
        const caducidadStr = _fmtFechaDisplay(data[i][4]);
        const stock = Number(data[i][3]);

        // Buscamos si ya existe esta combinación
        let existente = lotes.find(l => 
            l.presentacion_id === presId && 
            l.ubicacion_id === uId && 
            l.lote === lote && 
            _fmtFechaDisplay(l.caducidad) === caducidadStr
        );

        if (existente) {
            existente.stock_real += stock; 
        } else {
            lotes.push({
              producto_id: prodIdBuscado, 
              presentacion_id: presId, 
              ubicacion_id: uId,
              stock_real: stock, 
              caducidad: data[i][4], 
              elaboracion: data[i][5],
              lote: lote, 
              fecha_entrada: data[i][7],
              nombre_ubicacion: mapUbic[uId] || uId, 
              nombre_presentacion: mapPres[presId] || presId
            });
        }
      }
    }

    // 3. Descontar Carrito (Si aplica)
    if (carrito && Array.isArray(carrito)) {
      carrito.forEach(item => {
        const l = lotes.find(x => x.lote === String(item.lote).trim() && x.ubicacion_id === String(item.ubicacion_id).trim());
        if (l) l.stock_real -= Number(item.volumen_L);
      });
    }

    const validos = lotes.filter(l => l.stock_real > 0.001);
    if (validos.length === 0) return JSON.stringify({ success: false, error: "Sin stock disponible" });
    
    const getMs = (d) => (d instanceof Date) ? d.getTime() : 0;
    validos.sort((a, b) => getMs(a.elaboracion||a.fecha_entrada) - getMs(b.elaboracion||b.fecha_entrada));

    const mejor = validos[0];
    
    return JSON.stringify({
      success: true,
      mejor_candidato: { 
          presentacion_id: mejor.presentacion_id, 
          ubicacion_id: mejor.ubicacion_id, 
          lote: mejor.lote, 
          stock_real: mejor.stock_real 
      },
      lista_completa: validos.map(l => ({
        presentacion_id: l.presentacion_id, ubicacion_id: l.ubicacion_id, lote: l.lote,
        stock: l.stock_real.toFixed(2), caducidad: _fmtFechaDisplay(l.caducidad),
        nombre_ubicacion: l.nombre_ubicacion, nombre_presentacion: l.nombre_presentacion,
        es_sugerido: (l.lote === mejor.lote && l.ubicacion_id === mejor.ubicacion_id)
      }))
    });
  } catch (e) { return JSON.stringify({ success: false, error: e.message }); }
}

function _fmtF(f) {
  try {
    if(!f) return "";
    let s = String(f).trim();
    
    if(s.includes('-')) {
      let p = s.split('-');
      // Solo intenta procesar si realmente tiene 3 partes (día, mes, año)
      if(p.length >= 3) {
          return p[0].length === 4 ? s.substring(0,10) : `${p[2].split(' ')[0]}-${p[1]}-${p[0]}`;
      }
    }
    
    if(s.includes('/')) {
      let p = s.split('/');
      // Solo intenta procesar si realmente tiene 3 partes (día, mes, año)
      if(p.length >= 3) {
          return `${p[2].split(' ')[0]}-${p[1]}-${p[0]}`;
      }
    }
    
    return ""; // Si es un texto raro que no parece fecha, devuelve vacío
  } catch(e) {
    return ""; // Si ocurre cualquier error, no rompe el programa
  }
}
function formatearFechaInput(f) { return _fmtF(f); }

function EJECUTAR_DIAGNOSTICO_DEBUG() {
  const ss = obtenerSpreadsheet();
  let sD = ss.getSheetByName('DEBUG'); if (!sD) sD = ss.insertSheet('DEBUG'); sD.clear();
  let rep = [["HOJA", "EXISTE?", "FILAS"]];
  ['PEDIDOS', 'DETALLE_PEDIDOS', 'CLIENTES', 'INVENTARIO'].forEach(n => {
    let s = ss.getSheetByName(n); rep.push([n, s ? "SÍ" : "NO", s ? s.getLastRow() : "-"]);
  });
  sD.getRange(1,1,rep.length,3).setValues(rep);
}

// ==========================================
// 7. MÓDULO DE DESINCORPORACIÓN (BAJAS)
// ==========================================

function obtenerMaterialesCaducados() {
  const sInv = obtenerHojaSegura('INVENTARIO');
  if (!sInv || sInv.getLastRow() < 2) return [];

  const mapProd = {}, mapPres = {}, mapUbic = {};
  const sP = obtenerHojaSegura('PRODUCTOS');
  const sPr = obtenerHojaSegura('PRESENTACIONES');
  const sU = obtenerHojaSegura('UBICACIONES');

  if (sP) sP.getDataRange().getValues().slice(1).forEach(r => mapProd[String(r[0]).trim()] = r[1]);
  if (sPr) sPr.getDataRange().getValues().slice(1).forEach(r => mapPres[String(r[0]).trim()] = r[1]);
  if (sU) sU.getDataRange().getValues().slice(1).forEach(r => mapUbic[String(r[0]).trim()] = r[1]);

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
      } else if (typeof caducidad === 'string') {
        if (caducidad.includes('-')) {
          const partes = caducidad.split('T')[0].split('-'); // YYYY-MM-DD
          if (partes.length === 3) fechaCad = new Date(partes[0], partes[1] - 1, partes[2]);
        } else if (caducidad.includes('/')) {
          const partes = caducidad.split('/'); // DD/MM/YYYY
          if (partes.length === 3) fechaCad = new Date(partes[2], partes[1] - 1, partes[0]);
        }
      }

      // Si tiene fecha válida y ya pasó (es menor a hoy)
      if (fechaCad && fechaCad < hoy) {
        const pId = String(dataInv[i][0]).trim();
        const prId = String(dataInv[i][1]).trim();
        const uId = String(dataInv[i][2]).trim();

        caducados.push({
          producto_id: pId,
          producto: mapProd[pId] || pId,
          presentacion_id: prId,
          presentacion: mapPres[prId] || prId,
          ubicacion_id: uId,
          ubicacion: mapUbic[uId] || uId,
          lote: String(dataInv[i][6]).trim(),
          volumen: stock,
          caducidadStr: _fmtFechaDisplay(caducidad)
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

    // --- PASO 1: VERIFICAR ID DE DRIVE ---
    if(!ID_CARPETA_BAJAS_DRIVE || ID_CARPETA_BAJAS_DRIVE.includes("drive.google.com") || ID_CARPETA_BAJAS_DRIVE.includes("/")) {
       throw new Error("Paso 1: El ID de la carpeta de Drive es inválido. Asegúrate de poner solo las letras y números del final del link.");
    }
    
    // --- PASO 2: BUSCAR LA PLANTILLA ---
    const sTemplate = obtenerHojaSegura('TEMPLATE_BAJAS');
    if (!sTemplate) throw new Error("Paso 2: No se encontró la pestaña 'TEMPLATE_BAJAS'. Revisa que esté escrita exactamente así en tu Excel.");

    const sInv = obtenerHojaSegura('INVENTARIO');
    const sSal = obtenerHojaSegura('REGISTROS_SALIDA');
    const dataInv = sInv.getDataRange().getValues();
    const fechaHoy = new Date();
    const timestampFile = Utilities.formatDate(fechaHoy, Session.getScriptTimeZone(), "yyyyMMdd_HHmm");

    // --- PASO 3: DESCONTAR DEL INVENTARIO ---
    itemsBaja.forEach(item => {
      let restante = Number(item.volumen);
      for (let i = 1; i < dataInv.length; i++) {
        if (restante <= 0.001) break;
        const invProd = String(dataInv[i][0]).trim();
        const invLote = String(dataInv[i][6]).trim().toUpperCase();
        const invUbic = String(dataInv[i][2]).trim();

        if (invProd === String(item.producto_id).trim() && invLote === String(item.lote).trim().toUpperCase() && invUbic === String(item.ubicacion_id).trim()) {
          const stock = Number(dataInv[i][3]);
          if (stock > 0) {
             const restar = Math.min(stock, restante);
             sInv.getRange(i + 1, 4).setValue(stock - restar);
             restante -= restar;
          }
        }
      }
      sSal.appendRow([
        item.producto_id, item.producto, item.presentacion_id, item.presentacion,
        item.volumen, 0, item.ubicacion_id, item.lote,
        "DESINCORPORACIÓN", Session.getActiveUser().getEmail(), fechaHoy, `BAJA-${timestampFile}`
      ]);
    });

    // --- PASO 4: CONECTAR CON CARPETA DE DRIVE ---
    let folderDestino;
    try {
       folderDestino = DriveApp.getFolderById(ID_CARPETA_BAJAS_DRIVE);
    } catch(e) {
       throw new Error("Paso 4: No se encontró la carpeta en Drive. Revisa el ID y los permisos. (" + e.message + ")");
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
    } catch(e) {
        throw new Error("Paso 5: Falló al crear el Excel o copiar la plantilla. (" + e.message + ")");
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
            targetSheet.getRange(currentRow, 4).setValue("ENVASE PET DE " + item.presentacion);
            targetSheet.getRange(currentRow, 5).setValue(Number(item.volumen));
            
            totalVolumen += Number(item.volumen);
        });

        const endRow = startRow + itemsBaja.length - 1;

        // COPIAR FORMATO: Si hay más de 1 item, copiamos el formato de la fila inicial hacia abajo
        if (itemsBaja.length > 1) {
            const maxCols = targetSheet.getMaxColumns();
            const firstRowRange = targetSheet.getRange(startRow, 1, 1, maxCols);
            const targetRange = targetSheet.getRange(startRow + 1, 1, itemsBaja.length - 1, maxCols);
            firstRowRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
        }

        // LIMPIEZA Y TOTALES: Dejamos 1 fila en blanco, quitamos formato y ponemos totales
        rowTotal = endRow + 2; 
        
        // Quita bordes y fondos de las 2 filas debajo de la tabla para un aspecto limpio
        targetSheet.getRange(rowTotal - 1, 1, 2, targetSheet.getMaxColumns()).clearFormat();
        
        // Inserta "TOTAL" y la suma en negritas
        targetSheet.getRange(rowTotal, 4).setValue("TOTAL").setFontWeight("bold").setHorizontalAlignment("right");
        targetSheet.getRange(rowTotal, 5).setValue(totalVolumen).setFontWeight("bold");

        SpreadsheetApp.flush(); 
    } catch(e) {
        throw new Error("Paso 6: Falló al escribir los datos en el nuevo archivo. (" + e.message + ")");
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
        if (totalRows > rowTotal) pdfSheet.hideRows(rowTotal + 1, totalRows - rowTotal);
        
        SpreadsheetApp.flush();
        
        // Generamos el PDF basado en la hoja temporal
        const pdfBlob = excelFile.getAs(MimeType.PDF).setName(`${docName}.pdf`);
        pdfFile = folderDestino.createFile(pdfBlob);
        pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        
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
        excelFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch(e) {
        throw new Error("Paso 8: Falló al organizar el archivo Excel final. (" + e.message + ")");
    }
    
    // --- ÉXITO ---
    return { 
      success: true, 
      urlExcel: excelFile.getUrl(), 
      urlPDF: pdfFile.getUrl() 
    };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// 8. GESTIÓN DE DOCUMENTOS EN DRIVE (PEDIDOS)
// ==========================================

// REEMPLAZA ESTO CON EL ID DE TU CARPETA MAESTRA EN DRIVE
const ID_CARPETA_PADRE_PEDIDOS = "119ZLT4_yRFpkhndI5sPQF6lZCXA0FFUm"; 

function obtenerOCrearCarpetaPedido(idPedido) {
  const sPed = obtenerHojaSegura('PEDIDOS');
  const d = sPed.getDataRange().getValues();
  let rowIndex = -1;

  // Buscar la fila del pedido
  for(let i=1; i<d.length; i++) {
    if(String(d[i][0]).trim() === String(idPedido).trim()) {
      rowIndex = i + 1;
      break;
    }
  }

  if (rowIndex === -1) throw new Error("Pedido no encontrado en la base de datos.");

  let folderPadre;
  try {
     folderPadre = DriveApp.getFolderById(ID_CARPETA_PADRE_PEDIDOS);
  } catch(e) {
     throw new Error("No se encontró la carpeta principal en Drive. Verifica el ID.");
  }

  let folderDestino;
  let carpetas = folderPadre.getFoldersByName(idPedido);
  
  if (carpetas.hasNext()) {
    folderDestino = carpetas.next();
  } else {
    // Si no existe, la crea
    folderDestino = folderPadre.createFolder(idPedido);
    folderDestino.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    // CORRECCIÓN: Lo guardamos en la columna 15 (O) para no chocar con las fechas
    sPed.getRange(rowIndex, 15).setValue(folderDestino.getUrl());
  }

  return folderDestino;
}

function subirArchivoPedido(idPedido, nombreArchivo, base64Data, mimeType) {
  try {
    const folder = obtenerOCrearCarpetaPedido(idPedido);
    
    // Decodificar el archivo y crearlo en Drive
    const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, nombreArchivo);
    const file = folder.createFile(blob);
    
    return { success: true, url: file.getUrl(), nombre: file.getName(), urlCarpeta: folder.getUrl() };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function obtenerArchivosPedido(idPedido) {
  try {
    const folderPadre = DriveApp.getFolderById(ID_CARPETA_PADRE_PEDIDOS);
    let carpetas = folderPadre.getFoldersByName(idPedido);
    
    // Si no hay carpeta aún, regresamos vacío
    if (!carpetas.hasNext()) return { success: true, archivos: [], urlCarpeta: null };

    const folder = carpetas.next();
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