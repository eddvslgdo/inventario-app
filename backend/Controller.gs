/**
 * Controller.gs - VERSIÓN V21 (INTEGRAL FINAL)
 * - Mantiene el arreglo matemático de Envíos (V19).
 * - RESTAURA las funciones perdidas: transferirProducto y realizarTransformacion.
 * - Elimina la edición de nombres de ubicaciones (Solicitud usuario).
 * - Incluye todas las funciones de Entradas, Salidas y Ubicaciones.
 */

// ==========================================
// 0. HERRAMIENTAS BASE
// ==========================================
function obtenerHojaSegura(nombre) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(nombre);
}

function obtenerHojaOCrear(nombre, encabezados) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(nombre);
  if (!sheet) {
    sheet = ss.insertSheet(nombre);
    if (encabezados && encabezados.length > 0) sheet.appendRow(encabezados);
  }
  return sheet;
}

// ==========================================
// 1. CATÁLOGOS E INICIALIZACIÓN
// ==========================================
function obtenerCatalogos() {
  const sProd = obtenerHojaOCrear('PRODUCTOS', ['ID', 'NOMBRE', 'DESCRIPCION', 'UNIDAD']);
  const sPres = obtenerHojaOCrear('PRESENTACIONES', ['ID', 'NOMBRE', 'VOLUMEN']);
  const sUbic = obtenerHojaOCrear('UBICACIONES', ['ID', 'NOMBRE']);

  const leer = (s, esPres) => {
    if (!s || s.getLastRow() < 2) return [];
    const data = s.getDataRange().getValues();
    let res = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) {
        let item = { id: String(data[i][0]).trim(), nombre: data[i][1] };
        if (esPres) item.volumen = Number(data[i][2]) || 0;
        res.push(item);
      }
    }
    return res;
  };

  return {
    productos: leer(sProd, false),
    presentaciones: leer(sPres, true),
    ubicaciones: leer(sUbic, false)
  };
}

function obtenerListaClientes() {
  const s = obtenerHojaOCrear('CLIENTES', ['ID', 'NOMBRE', 'EMPRESA', 'DIRECCION', 'TELEFONO', 'EMAIL']);
  if (!s || s.getLastRow() < 2) return [];
  const data = s.getDataRange().getValues();
  let clientes = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) clientes.push({
      id: data[i][0], nombre: data[i][1], empresa: data[i][2],
      direccion: data[i][3], telefono: data[i][4], email: data[i][5]
    });
  }
  return clientes;
}

// ==========================================
// 2. GESTIÓN DE INVENTARIO (DASHBOARDS)
// ==========================================
function obtenerDatosUbicaciones() {
  const sInv = obtenerHojaOCrear('INVENTARIO', ['ID_PROD', 'ID_PRES', 'ID_UBIC', 'STOCK', 'CADUCIDAD', 'ELABORACION', 'LOTE', 'F_ENTRADA', 'PROVEEDOR']);
  const sUbic = obtenerHojaOCrear('UBICACIONES', []);
  const sProd = obtenerHojaOCrear('PRODUCTOS', []);
  const sPres = obtenerHojaOCrear('PRESENTACIONES', []);

  const mapProd = {}, mapPres = {}, mapPresVol = {};
  if (sProd.getLastRow() > 1) sProd.getDataRange().getValues().forEach((r, i) => { if (i > 0) mapProd[String(r[0]).trim()] = r[1]; });
  if (sPres.getLastRow() > 1) sPres.getDataRange().getValues().forEach((r, i) => {
    if (i > 0) {
      const id = String(r[0]).trim();
      mapPres[id] = r[1];
      mapPresVol[id] = Number(r[2]) || 0;
    }
  });

  let ubicaciones = [];
  if (sUbic.getLastRow() > 1) {
    const dU = sUbic.getDataRange().getValues();
    for (let i = 1; i < dU.length; i++) {
      if (dU[i][0]) ubicaciones.push({ id: String(dU[i][0]).trim(), nombre: dU[i][1] || 'Sin Nombre', items: [], totalVolumen: 0 });
    }
  }

  if (sInv.getLastRow() > 1) {
    const dInv = sInv.getDataRange().getValues();
    for (let i = 1; i < dInv.length; i++) {
      const stock = Number(dInv[i][3]);
      if (stock > 0.001) {
        const uId = String(dInv[i][2]).trim();
        let ubic = ubicaciones.find(u => u.id === uId) || { id: uId, nombre: "Ubic: " + uId, items: [], totalVolumen: 0 };
        if (!ubicaciones.includes(ubic)) ubicaciones.push(ubic);

        const pId = String(dInv[i][0]).trim();
        const presId = String(dInv[i][1]).trim();
        const nProd = mapProd[pId] || pId;
        const nPres = mapPres[presId] || presId;

        ubic.items.push({
          raw_producto_id: pId, producto: nProd, raw_presentacion_id: presId, presentacion: nPres,
          lote: String(dInv[i][6]), volumen: stock, volumen_nominal: mapPresVol[presId] || 0,
          caducidad: dInv[i][4] instanceof Date ? dInv[i][4].toLocaleDateString() : dInv[i][4],
          nombre_completo: `${nProd} (${nPres})`
        });
        ubic.totalVolumen += stock;
      }
    }
  }
  return ubicaciones;
}

function obtenerDatosProductos() {
  const sInv = obtenerHojaOCrear('INVENTARIO', []);
  if(sInv.getLastRow() < 2) return [];

  const mapProd = {}, mapUbic = {}, mapPres = {};
  const sP = obtenerHojaOCrear('PRODUCTOS', []), sU = obtenerHojaOCrear('UBICACIONES', []), sPr = obtenerHojaOCrear('PRESENTACIONES', []);
  if(sP.getLastRow()>1) sP.getDataRange().getValues().forEach((r,i)=> { if(i>0) mapProd[String(r[0]).trim()] = r[1]; });
  if(sU.getLastRow()>1) sU.getDataRange().getValues().forEach((r,i)=> { if(i>0) mapUbic[String(r[0]).trim()] = r[1]; });
  if(sPr.getLastRow()>1) sPr.getDataRange().getValues().forEach((r,i)=> { if(i>0) mapPres[String(r[0]).trim()] = r[1]; });

  const dataInv = sInv.getDataRange().getValues();
  let productosMap = {};

  for (let i = 1; i < dataInv.length; i++) {
    const stock = Number(dataInv[i][3]);
    if (stock > 0.001) {
       const pId = String(dataInv[i][0]).trim();
       if (!productosMap[pId]) productosMap[pId] = { id: pId, nombre: mapProd[pId] || pId, totalVolumen: 0, lotes: [] };
       productosMap[pId].totalVolumen += stock;
       const uId = String(dataInv[i][2]).trim();
       const presId = String(dataInv[i][1]).trim();
       productosMap[pId].lotes.push({
         lote: dataInv[i][6], volumen: stock, ubicacion: mapUbic[uId] || uId,
         presentacion: mapPres[presId] || presId, caducidad: dataInv[i][4] instanceof Date ? dataInv[i][4].toLocaleDateString() : ''
       });
    }
  }
  return Object.values(productosMap);
}

// ==========================================
// 3. REGISTRO DE MOVIMIENTOS (ENTRADAS)
// ==========================================
function registrarNuevoProducto(d) {
  const s = obtenerHojaOCrear('PRODUCTOS', ['ID', 'NOMBRE', 'DESCRIPCION', 'UNIDAD']);
  s.appendRow([Utilities.getUuid(), d.nombre, d.descripcion, d.unidad]); return true;
}

function registrarNuevaPresentacion(nombre) {
  const s = obtenerHojaOCrear('PRESENTACIONES', ['ID', 'NOMBRE', 'VOLUMEN']);
  let volumen = 0;
  const match = String(nombre).match(/[\d\.]+/);
  if (match) {
    volumen = parseFloat(match[0]);
    if (String(nombre).toLowerCase().includes('ml')) volumen /= 1000;
  }
  s.appendRow([Utilities.getUuid(), nombre, volumen]); return true;
}

function registrarNuevaUbicacion(nombre) {
  const s = obtenerHojaOCrear('UBICACIONES', ['ID', 'NOMBRE']);
  s.appendRow([Utilities.getUuid(), nombre]); return true;
}

// Se elimina actualizarNombreUbicacion por petición del usuario

function borrarUbicacion(id) {
  const s = obtenerHojaOCrear('UBICACIONES', []);
  const d = s.getDataRange().getValues();
  for(let i=1; i<d.length; i++) if(d[i][0]==id) { s.deleteRow(i+1); return true; }
}

function registrarEntradaUnica(datos) { return registrarEntradaMasiva([datos]); }

function registrarEntradaMasiva(lista) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const sInv = obtenerHojaOCrear('INVENTARIO', ['ID_PROD', 'ID_PRES', 'ID_UBIC', 'STOCK', 'CADUCIDAD', 'ELABORACION', 'LOTE', 'F_ENTRADA', 'PROVEEDOR']);
    const sLog = obtenerHojaOCrear('REGISTROS_ENTRADA', ['FECHA', 'PROD', 'PRES', 'UBIC', 'VOL', 'LOTE', 'PROVEEDOR', 'TIPO']);
    lista.forEach(d => {
      if(d.producto_id.includes("Selecciona") || d.producto_id === "undefined") throw new Error("Error: Producto no válido.");
      sInv.appendRow([d.producto_id, d.presentacion_id, d.ubicacion_id, Number(d.volumen_L), d.fecha_caducidad, d.fecha_elaboracion, d.lote, new Date(), d.proveedor]);
      sLog.appendRow([new Date(), d.producto_id, d.presentacion_id, d.ubicacion_id, d.volumen_L, d.lote, d.proveedor, 'Entrada']);
    });
    return true;
  } catch(e){ throw e; } finally { lock.releaseLock(); }
}

// ==========================================
// 4. MOVER Y TRANSFORMAR (¡RESTAURADAS!)
// ==========================================
function transferirProducto(origenId, destinoId, productoId, presentacionId, lote, cantidadMover) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sInv = obtenerHojaOCrear('INVENTARIO', []);
    const data = sInv.getDataRange().getValues();
    
    // Buscar origen
    let idxOrigen = -1, stockOrigen = 0;
    for(let i=1; i<data.length; i++) {
      if(String(data[i][0])==String(productoId) && String(data[i][1])==String(presentacionId) && 
         String(data[i][2])==String(origenId) && String(data[i][6])==String(lote)) {
         idxOrigen = i+1; stockOrigen = Number(data[i][3]); break;
      }
    }
    if(idxOrigen == -1) throw new Error("Origen no encontrado");
    if(stockOrigen < cantidadMover - 0.001) throw new Error("Stock insuficiente");

    // Restar
    const nStock = stockOrigen - cantidadMover;
    if(nStock <= 0.001) sInv.deleteRow(idxOrigen); else sInv.getRange(idxOrigen, 4).setValue(nStock);

    // Sumar a Destino (Búsqueda nueva por si cambiaron índices)
    const d2 = sInv.getDataRange().getValues();
    let idxDest = -1;
    for(let i=1; i<d2.length; i++) {
      if(String(d2[i][0])==String(productoId) && String(d2[i][1])==String(presentacionId) && 
         String(d2[i][2])==String(destinoId) && String(d2[i][6])==String(lote)) {
         idxDest = i+1; break;
      }
    }

    if(idxDest > -1) {
      sInv.getRange(idxDest, 4).setValue(Number(sInv.getRange(idxDest, 4).getValue()) + Number(cantidadMover));
    } else {
      const meta = data[idxOrigen-1]; // Usar datos originales
      sInv.appendRow([productoId, presentacionId, destinoId, Number(cantidadMover), meta[4], meta[5], lote, new Date(), "Transferencia"]);
    }
    return {success:true};
  } catch(e) { return {success:false, error:e.message}; } finally { lock.releaseLock(); }
}

function realizarTransformacion(datos) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sInv = obtenerHojaOCrear('INVENTARIO', []);
    const data = sInv.getDataRange().getValues();
    
    // Buscar Origen
    let idxOrg = -1;
    for(let i=1; i<data.length; i++) {
      if(String(data[i][2])==String(datos.origenId) && String(data[i][0])==String(datos.productoIdOrigen) && 
         String(data[i][6]).toUpperCase()==String(datos.loteOrigen).toUpperCase()) {
         idxOrg = i+1; break;
      }
    }
    if(idxOrg == -1) throw new Error("Lote origen no encontrado");
    
    const stock = Number(sInv.getRange(idxOrg, 4).getValue());
    if(stock < Number(datos.cantidadLitros)) throw new Error("Stock insuficiente");

    // Restar
    const nS = stock - Number(datos.cantidadLitros);
    if(nS <= 0.001) sInv.deleteRow(idxOrg); else sInv.getRange(idxOrg, 4).setValue(nS);

    // Agregar nuevo
    const lFin = datos.nuevoLote || datos.loteOrigen;
    sInv.appendRow([datos.nuevoProductoId, datos.nuevaPresentacionId, datos.origenId, Number(datos.cantidadLitros), 
                    data[idxOrg-1][4], data[idxOrg-1][5], lFin, new Date(), "Transformación"]);
    
    return {success:true};
  } catch(e) { return {success:false, error:e.message}; } finally { lock.releaseLock(); }
}

// ==========================================
// 5. PROCESAMIENTO DE SALIDAS (PEDIDOS V19)
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
    
    // 1. VALIDACIÓN
    itemsCarrito.forEach(item => {
      const p = String(item.producto_id);
      if (p.includes("Selecciona") || p === "undefined" || !p) throw new Error("⚠️ Producto inválido. Recarga.");
      if (!item.lote || item.lote === "undefined") throw new Error("⚠️ Lote inválido.");
    });

    // 2. CONSOLIDACIÓN (AGRUPAR ITEMS IGUALES)
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

    const sInv = obtenerHojaOCrear('INVENTARIO', []), sSal = obtenerHojaOCrear('REGISTROS_SALIDA', []), sDet = obtenerHojaOCrear('DETALLE_PEDIDOS', ['ID_PEDIDO', 'PRODUCTO', 'PRESENTACION', 'LOTE', 'VOLUMEN']);
    const dataInv = sInv.getDataRange().getValues();

    itemsProcesar.forEach(item => {
      sDet.appendRow([idPedido, item.nombre_producto, item.nombre_presentacion || '---', item.lote, Number(item.volumen_L)]);

      let updated = false;
      for (let i = 1; i < dataInv.length; i++) {
        if (String(dataInv[i][0]).trim() == String(item.producto_id).trim() && 
            String(dataInv[i][6]).trim().toUpperCase() == String(item.lote).trim().toUpperCase() &&
            String(dataInv[i][2]).trim() == String(item.ubicacion_id).trim()) {
          const actual = Number(dataInv[i][3]);
          if(actual < item.volumen_L - 0.01) throw new Error(`Stock insuficiente para ${item.nombre_producto} (${actual}L disponibles)`);
          sInv.getRange(i + 1, 4).setValue(actual - item.volumen_L);
          updated = true; break;
        }
      }
      if(!updated) console.warn("Lote no encontrado para restar: " + item.lote);
      
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

  const sDet = obtenerHojaSegura('DETALLE_PEDIDOS');
  const sCli = obtenerHojaSegura('CLIENTES');

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
        items.push({ producto: pName, presentacion: dD[i][2], lote: dD[i][3], volumen: dD[i][4] });
      }
    }
  }
  return { cabecera: cab, items: items };
}

function actualizarPedido(id, fe, fr, st) {
  const s = obtenerHojaSegura('PEDIDOS');
  const d = s.getDataRange().getValues();
  for(let i=1; i<d.length; i++) {
    if(String(d[i][0]).trim() === String(id).trim()) {
      s.getRange(i+1, 11).setValue(st); s.getRange(i+1, 13).setValue(fe); s.getRange(i+1, 14).setValue(fr); return "OK";
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
    const mapUbic = {}, mapPres = {};
    const sU = obtenerHojaSegura('UBICACIONES'), sP = obtenerHojaSegura('PRESENTACIONES');
    if(sU) sU.getDataRange().getValues().forEach((r,i)=> {if(i>0) mapUbic[String(r[0]).trim()] = r[1]});
    if(sP) sP.getDataRange().getValues().forEach((r,i)=> {if(i>0) mapPres[String(r[0]).trim()] = r[1]});

    let lotes = [];
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === prodIdBuscado) {
        const uId = String(data[i][2]).trim(); const presId = String(data[i][1]).trim();
        lotes.push({
          producto_id: prodIdBuscado, presentacion_id: presId, ubicacion_id: uId,
          stock_real: Number(data[i][3]), caducidad: data[i][4], elaboracion: data[i][5],
          lote: String(data[i][6]).trim(), fecha_entrada: data[i][7],
          nombre_ubicacion: mapUbic[uId] || uId, nombre_presentacion: mapPres[presId] || presId
        });
      }
    }

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
      mejor_candidato: { presentacion_id: mejor.presentacion_id, ubicacion_id: mejor.ubicacion_id, lote: mejor.lote, stock_real: mejor.stock_real },
      lista_completa: lotes.map(l => ({
        presentacion_id: l.presentacion_id, ubicacion_id: l.ubicacion_id, lote: l.lote,
        stock: l.stock_real.toFixed(2), caducidad: l.caducidad instanceof Date ? l.caducidad.toLocaleDateString() : String(l.caducidad),
        nombre_ubicacion: l.nombre_ubicacion, nombre_presentacion: l.nombre_presentacion,
        es_sugerido: (l.lote === mejor.lote && l.ubicacion_id === mejor.ubicacion_id)
      }))
    });
  } catch (e) { return JSON.stringify({ success: false, error: e.message }); }
}

function _fmtF(f) {
  if(!f) return ""; if(f.match(/^\d{4}-\d{2}-\d{2}$/)) return f;
  if(f.includes('-')) { let p = f.split('-'); return p[0].length===4 ? f : `${p[2].split(' ')[0]}-${p[1]}-${p[0]}`; }
  if(f.includes('/')) { let p = f.split('/'); return `${p[2].split(' ')[0]}-${p[1]}-${p[0]}`; }
  return "";
}
function formatearFechaInput(f) { return _fmtF(f); }

function EJECUTAR_DIAGNOSTICO_DEBUG() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sD = ss.getSheetByName('DEBUG'); if (!sD) sD = ss.insertSheet('DEBUG'); sD.clear();
  let rep = [["HOJA", "EXISTE?", "FILAS"]];
  ['PEDIDOS', 'DETALLE_PEDIDOS', 'CLIENTES', 'INVENTARIO'].forEach(n => {
    let s = ss.getSheetByName(n); rep.push([n, s ? "SÍ" : "NO", s ? s.getLastRow() : "-"]);
  });
  sD.getRange(1,1,rep.length,3).setValues(rep);
}