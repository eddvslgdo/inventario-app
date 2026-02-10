/**
 * Controller.gs - VERSIÓN FINAL CON TODAS LAS FUNCIONES Y CORRECCIÓN DE SALIDAS
 */

// ==========================================
// 0. HERRAMIENTA DE REPARACIÓN
// ==========================================
function obtenerHojaOCrear(nombre, encabezados) {
  const ss = SpreadsheetApp.openById(ID_HOJA);
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
  const headersProd = ['ID', 'NOMBRE', 'DESCRIPCION', 'UNIDAD'];
  const headersPres = ['ID', 'NOMBRE', 'VOLUMEN'];
  const headersUbic = ['ID', 'NOMBRE'];

  const sProd = obtenerHojaOCrear('PRODUCTOS', headersProd);
  const sPres = obtenerHojaOCrear('PRESENTACIONES', headersPres);
  const sUbic = obtenerHojaOCrear('UBICACIONES', headersUbic);

  const leerData = (sheet, esPresentacion = false) => {
    if (sheet.getLastRow() < 2) return [];
    const data = sheet.getDataRange().getValues();
    let res = [];
    for(let i=1; i<data.length; i++) {
      if(data[i][0]) {
        let item = { id: String(data[i][0]).trim(), nombre: data[i][1] };
        // Si es presentación, leemos el volumen. Si no existe, ponemos 0.
        if(esPresentacion) item.volumen = Number(data[i][2]) || 0; 
        res.push(item);
      }
    }
    return res;
  };

  return {
    productos: leerData(sProd),
    presentaciones: leerData(sPres, true),
    ubicaciones: leerData(sUbic)
  };
}

function obtenerListaClientes() {
  const sheet = obtenerHojaOCrear('CLIENTES', ['ID', 'NOMBRE', 'EMPRESA', 'DIRECCION', 'TELEFONO', 'EMAIL']);
  if (sheet.getLastRow() < 2) return [];
  const data = sheet.getDataRange().getValues();
  let clientes = [];
  for(let i=1; i<data.length; i++) {
    if(data[i][0]) clientes.push({
        id: data[i][0], nombre: data[i][1], empresa: data[i][2],
        direccion: data[i][3], telefono: data[i][4], email: data[i][5]
    });
  }
  return clientes;
}

// ==========================================
// 2. DASHBOARDS (CORRECCIÓN CRÍTICA DE VOLUMEN)
// ==========================================
function obtenerDatosUbicaciones() {
  const sheetInv = obtenerHojaOCrear('INVENTARIO', ['ID_PROD', 'ID_PRES', 'ID_UBIC', 'STOCK', 'CADUCIDAD', 'ELABORACION', 'LOTE', 'F_ENTRADA', 'PROVEEDOR']);
  const sheetUbic = obtenerHojaOCrear('UBICACIONES', ['ID', 'NOMBRE']);
  
  // 1. Mapas de Nombres
  const mapProd = {}; 
  const sProd = obtenerHojaOCrear('PRODUCTOS', []);
  if(sProd.getLastRow() > 1) sProd.getDataRange().getValues().forEach((r,i) => { if(i>0) mapProd[String(r[0]).trim()] = r[1]; });

  // 2. Mapas de Presentaciones (NOMBRE Y VOLUMEN)
  const mapPres = {}; 
  const mapPresVol = {}; // <--- AQUÍ GUARDAMOS EL VOLUMEN UNITARIO
  const sPres = obtenerHojaOCrear('PRESENTACIONES', []);
  if(sPres.getLastRow() > 1) {
    sPres.getDataRange().getValues().forEach((r,i) => { 
      if(i>0) {
        const idPres = String(r[0]).trim();
        mapPres[idPres] = r[1];
        mapPresVol[idPres] = Number(r[2]) || 0; // Columna C es el volumen
      }
    });
  }

  let ubicaciones = [];
  if(sheetUbic.getLastRow() > 1) {
    const dataU = sheetUbic.getDataRange().getValues();
    for(let i=1; i<dataU.length; i++) {
      if(dataU[i][0]) ubicaciones.push({ id: String(dataU[i][0]).trim(), nombre: dataU[i][1] || 'Sin Nombre', items: [], totalVolumen: 0 });
    }
  }

  const dataInv = sheetInv.getDataRange().getValues();
  for (let i = 1; i < dataInv.length; i++) {
    const stock = Number(dataInv[i][3]);
    if (stock > 0.001) {
      const uId = String(dataInv[i][2]).trim();
      let ubic = ubicaciones.find(u => u.id === uId);
      if (!ubic) {
        ubic = { id: uId, nombre: "Ubic: " + uId, items: [], totalVolumen: 0 };
        ubicaciones.push(ubic);
      }
      const pId = String(dataInv[i][0]).trim();
      const presId = String(dataInv[i][1]).trim();
      const nProd = mapProd[pId] || (pId.length>8 ? "ID:"+pId.substring(0,5) : pId);
      const nPres = mapPres[presId] || presId;
      
      // RECUPERAMOS EL VOLUMEN UNITARIO PARA EL FRONTEND
      const volUnitario = mapPresVol[presId] || 0;

      ubic.items.push({
        raw_producto_id: pId, 
        producto: nProd,
        raw_presentacion_id: presId, 
        presentacion: nPres,
        lote: String(dataInv[i][6]), 
        volumen: stock, // Stock total en litros
        volumen_nominal: volUnitario, // <--- ESTO ARREGLA EL ERROR DE "SIN VOLUMEN"
        caducidad: dataInv[i][4] instanceof Date ? dataInv[i][4].toLocaleDateString() : dataInv[i][4],
        nombre_completo: `${nProd} (${nPres})`
      });
      ubic.totalVolumen += stock;
    }
  }
  return ubicaciones;
}

function obtenerDatosProductos() {
  const sheetInv = obtenerHojaOCrear('INVENTARIO', []);
  if(sheetInv.getLastRow() < 2) return [];

  const mapProd = {}; const sP = obtenerHojaOCrear('PRODUCTOS', []);
  if(sP.getLastRow()>1) sP.getDataRange().getValues().forEach((r,i)=> {if(i>0) mapProd[String(r[0]).trim()] = r[1]});
  const mapUbic = {}; const sU = obtenerHojaOCrear('UBICACIONES', []);
  if(sU.getLastRow()>1) sU.getDataRange().getValues().forEach((r,i)=> {if(i>0) mapUbic[String(r[0]).trim()] = r[1]});
  const mapPres = {}; const sPr = obtenerHojaOCrear('PRESENTACIONES', []);
  if(sPr.getLastRow()>1) sPr.getDataRange().getValues().forEach((r,i)=> {if(i>0) mapPres[String(r[0]).trim()] = r[1]});

  const dataInv = sheetInv.getDataRange().getValues();
  let productosMap = {};

  for (let i = 1; i < dataInv.length; i++) {
    const stock = Number(dataInv[i][3]);
    if (stock > 0.001) {
       const pId = String(dataInv[i][0]).trim();
       if (!productosMap[pId]) {
         productosMap[pId] = { 
           id: pId, nombre: mapProd[pId] || "ID: " + pId.substring(0,8), 
           totalVolumen: 0, lotes: [] 
         };
       }
       productosMap[pId].totalVolumen += stock;
       const uId = String(dataInv[i][2]).trim();
       const presId = String(dataInv[i][1]).trim();

       productosMap[pId].lotes.push({
         lote: dataInv[i][6], volumen: stock,
         ubicacion: mapUbic[uId] || uId, 
         presentacion: mapPres[presId] || presId, 
         caducidad: dataInv[i][4] instanceof Date ? dataInv[i][4].toLocaleDateString() : ''
       });
    }
  }
  return Object.values(productosMap);
}

// ==========================================
// 3. FIFO
// ==========================================
function obtenerSugerenciaFIFO(productoId, carrito) {
  try {
    if (!productoId) throw new Error("ID vacío");
    const prodIdBuscado = String(productoId).trim();
    const sheetInv = obtenerHojaOCrear('INVENTARIO', []);
    if(sheetInv.getLastRow() < 2) return JSON.stringify({ success: false, error: "Inventario vacío" });
    const data = sheetInv.getDataRange().getValues();
    
    const mapUbic = {}; const sU = obtenerHojaOCrear('UBICACIONES', []);
    if(sU.getLastRow()>1) sU.getDataRange().getValues().forEach((r,i)=> {if(i>0) mapUbic[String(r[0]).trim()] = r[1]});
    const mapPres = {}; const sP = obtenerHojaOCrear('PRESENTACIONES', []);
    if(sP.getLastRow()>1) sP.getDataRange().getValues().forEach((r,i)=> {if(i>0) mapPres[String(r[0]).trim()] = r[1]});

    let lotes = [];
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === prodIdBuscado) {
        const uId = String(data[i][2]).trim();
        const presId = String(data[i][1]).trim();
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
    if (validos.length === 0) return JSON.stringify({ success: false, error: "Sin stock." });

    const getMs = (d) => (d instanceof Date) ? d.getTime() : 0;
    validos.sort((a, b) => {
      let fA = getMs(a.elaboracion) || getMs(a.fecha_entrada);
      let fB = getMs(b.elaboracion) || getMs(b.fecha_entrada);
      return fA - fB;
    });

    const mejor = validos[0];
    return JSON.stringify({
      success: true,
      mejor_candidato: {
         presentacion_id: mejor.presentacion_id, ubicacion_id: mejor.ubicacion_id,
         lote: mejor.lote, stock_real: mejor.stock_real
      },
      lista_completa: lotes.map(l => ({
        presentacion_id: l.presentacion_id, ubicacion_id: l.ubicacion_id, lote: l.lote,
        stock: l.stock_real.toFixed(2),
        caducidad: l.caducidad instanceof Date ? l.caducidad.toLocaleDateString() : String(l.caducidad),
        nombre_ubicacion: l.nombre_ubicacion, nombre_presentacion: l.nombre_presentacion,
        es_sugerido: (l.lote === mejor.lote && l.ubicacion_id === mejor.ubicacion_id)
      }))
    });
  } catch (e) { return JSON.stringify({ success: false, error: e.message }); }
}

// ==========================================
// 4. TRANSACCIONES
// ==========================================
function registrarEntradaUnica(datos) { return registrarEntradaMasiva([datos]); }

function registrarEntradaMasiva(lista) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheetInv = obtenerHojaOCrear('INVENTARIO', ['ID_PROD', 'ID_PRES', 'ID_UBIC', 'STOCK', 'CADUCIDAD', 'ELABORACION', 'LOTE', 'F_ENTRADA', 'PROVEEDOR']);
    const sheetLog = obtenerHojaOCrear('REGISTROS_ENTRADA', ['FECHA', 'PROD', 'PRES', 'UBIC', 'VOL', 'LOTE', 'PROVEEDOR', 'TIPO']);
    
    lista.forEach(d => {
       sheetInv.appendRow([d.producto_id, d.presentacion_id, d.ubicacion_id, Number(d.volumen_L), d.fecha_caducidad, d.fecha_elaboracion, d.lote, new Date(), d.proveedor]);
       sheetLog.appendRow([new Date(), d.producto_id, d.presentacion_id, d.ubicacion_id, d.volumen_L, d.lote, d.proveedor, 'Entrada']);
    });
    return true;
  } catch(e) { throw e; } finally { lock.releaseLock(); }
}

function registrarSalidaMasiva(lista) {
  return procesarPedidoCompleto({
      idCliente: 'GENERICO', nombreCliente: 'Salida Rápida', direccion:'-', telefono:'-', email:'-', tipoEnvio:'Nacional', paqueteria:'Mostrador', guia:'-', costoEnvio:0
  }, lista, false);
}

// =========================================================================
// FUNCIÓN CORREGIDA Y ROBUSTA PARA PROCESAR PEDIDOS Y SALIDAS
// =========================================================================
function procesarPedidoCompleto(datosPedido, itemsCarrito, guardarCliente) {
  const lock = LockService.getScriptLock();
  try {
    // Esperamos hasta 30 segundos para evitar choques de escritura
    lock.waitLock(30000);
    
    const idPedido = "PED-" + Math.floor(Date.now() / 1000);
    const fechaHoy = new Date();

    // 1. GUARDAR CLIENTE (Solo si se solicitó)
    if (guardarCliente) {
        const sheetCli = obtenerHojaOCrear('CLIENTES', ['ID', 'NOMBRE', 'EMPRESA', 'DIRECCION', 'TELEFONO', 'EMAIL']);
        let idCli = datosPedido.idCliente;
        // Si es nuevo o no tiene ID, generamos uno
        if (!idCli || idCli === 'nuevo' || idCli === '') {
            idCli = "CLI-" + Math.floor(Math.random()*10000);
        }
        sheetCli.appendRow([
            idCli, 
            datosPedido.nombreCliente, 
            datosPedido.empresa, 
            datosPedido.direccion, 
            datosPedido.telefono, 
            datosPedido.email
        ]);
    }

    // 2. REGISTRAR EL PEDIDO (Cabecera)
    const sheetPed = obtenerHojaOCrear('PEDIDOS', ['ID_PEDIDO', 'FECHA', 'ID_CLIENTE', 'NOMBRE', 'DIRECCION', 'TELEFONO', 'PAQUETERIA', 'GUIA', 'TIPO', 'COSTO', 'ESTATUS', 'F_EST', 'F_REAL', 'LINK']);
    sheetPed.appendRow([
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
        'Pendiente', // Estatus inicial
        '', '', ''   // Fechas y link vacíos
    ]);

    // 3. PROCESAR INVENTARIO Y DETALLE (Salidas)
    const sheetInv = obtenerHojaOCrear('INVENTARIO', []);
    const sheetSal = obtenerHojaOCrear('REGISTROS_SALIDA', ['PROD_ID', 'PROD_NOM', 'PRES_ID', 'PRES_NOM', 'VOL', 'PZAS', 'UBIC', 'LOTE', 'CLIENTE', 'USUARIO', 'FECHA', 'ID_PEDIDO']);
    
    // Obtenemos todos los datos del inventario de una vez para ser rápidos
    const dataInv = sheetInv.getDataRange().getValues();
    
    // Recorremos cada producto del carrito
    itemsCarrito.forEach(item => {
      let inventarioActualizado = false;
      
      // Limpieza de datos recibidos (Trim y UpperCase para evitar errores tontos)
      const reqProd = String(item.producto_id).trim();
      const reqPres = String(item.presentacion_id).trim();
      const reqUbic = String(item.ubicacion_id).trim();
      const reqLote = String(item.lote).trim().toUpperCase(); // CLAVE: Lote siempre mayúsculas
      
      // Buscamos coincidencia en el inventario
      for (let i = 1; i < dataInv.length; i++) {
        const invProd = String(dataInv[i][0]).trim();
        const invPres = String(dataInv[i][1]).trim();
        const invUbic = String(dataInv[i][2]).trim();
        const invLote = String(dataInv[i][6]).trim().toUpperCase();

        // Si todo coincide, descontamos
        if (invProd === reqProd && invPres === reqPres && invUbic === reqUbic && invLote === reqLote) {
          
          const stockActual = Number(dataInv[i][3]);
          const cantidadRestar = Number(item.volumen_L);
          
          const nuevoStock = stockActual - cantidadRestar;
          
          // Guardamos el nuevo stock en la celda correspondiente (Columna D = índice 4)
          sheetInv.getRange(i + 1, 4).setValue(nuevoStock);
          
          inventarioActualizado = true;
          break; // Dejamos de buscar este item, pasamos al siguiente
        }
      }

      // 4. GUARDAR EN REGISTROS_SALIDA (Independientemente de si se descontó o no)
      // Esto asegura que la venta quede registrada aunque el inventario tenga un error
      sheetSal.appendRow([
        item.producto_id, 
        item.nombre_producto || '---', 
        item.presentacion_id, 
        item.nombre_presentacion || '---', 
        Number(item.volumen_L), 
        Number(item.piezas), 
        item.ubicacion_id, 
        item.lote, 
        datosPedido.nombreCliente, 
        Session.getActiveUser().getEmail(), 
        fechaHoy, 
        idPedido
      ]);
      
      if (!inventarioActualizado) {
         console.warn("⚠️ Advertencia: No se encontró stock exacto para descontar el lote " + item.lote + ", pero se registró la salida.");
      }
    });

    return { success: true, idPedido: idPedido };

  } catch (e) {
    // Si algo falla, devolvemos el error al frontend para mostrarlo
    return { success: false, error: "Error en servidor: " + e.message };
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// 5. HISTORIAL / CREACIÓN
// ==========================================
function obtenerHistorialPedidos() {
  const sheet = obtenerHojaOCrear('PEDIDOS', []);
  if(sheet.getLastRow() < 2) return [];
  
  const data = sheet.getDataRange().getValues();
  let pedidos = [];
  
  // Recorremos de abajo hacia arriba (más recientes primero)
  for(let i=data.length-1; i>=1; i--) {
    // IMPORTANTE: Verificamos que data[i][0] (el ID) exista
    if(data[i][0] && String(data[i][0]).trim() !== "") {
      pedidos.push({
        id: data[i][0],
        fecha: data[i][1] instanceof Date ? data[i][1].toLocaleDateString() : data[i][1],
        cliente: data[i][3], 
        destino: data[i][4], 
        paqueteria: data[i][6], 
        guia: data[i][7], 
        estatus: data[i][10] || 'Pendiente'
      });
    }
  }
  return pedidos;
}

function obtenerDetallePedidoCompleto(idPedido) {
  const sheetItems = obtenerHojaOCrear('REGISTROS_SALIDA', []);
  if(sheetItems.getLastRow() < 2) return { cabecera: {}, items: [] };
  const sheetPed = obtenerHojaOCrear('PEDIDOS', []);
  let cabecera = { id: idPedido, cliente: "Desconocido" };

  if(sheetPed.getLastRow() > 1) {
     const d = sheetPed.getDataRange().getValues();
     for(let i=1; i<d.length; i++) {
        if(String(d[i][0]) === String(idPedido)) {
           cabecera = {
             id: d[i][0], cliente: d[i][3], direccion: d[i][4], 
             paqueteria: d[i][6], guia: d[i][7],
             fechaEst: d[i][11] instanceof Date ? d[i][11].toISOString().split('T')[0] : '',
             fechaReal: d[i][12] instanceof Date ? d[i][12].toISOString().split('T')[0] : '',
             estatus: d[i][10]
           };
           break;
        }
     }
  }
  const dataItems = sheetItems.getDataRange().getValues();
  let items = [];
  const colIdIndex = dataItems[0].length - 1; 
  for(let i=1; i<dataItems.length; i++){
    if(String(dataItems[i][colIdIndex]) === String(idPedido)){
      items.push({
        producto: dataItems[i][1], presentacion: dataItems[i][3],
        lote: dataItems[i][7], volumen: dataItems[i][4], cantidad: dataItems[i][5]
      });
    }
  }
  return { cabecera: cabecera, items: items };
}

function actualizarPedido(idPedido, fEst, fReal, estatus) {
  const sheet = obtenerHojaOCrear('PEDIDOS', []);
  if(sheet.getLastRow() < 2) return "Error";
  const data = sheet.getDataRange().getValues();
  for(let i=1; i<data.length; i++){
    if(String(data[i][0]) === String(idPedido)){
      sheet.getRange(i+1, 12).setValue(fEst); 
      sheet.getRange(i+1, 13).setValue(fReal); 
      sheet.getRange(i+1, 11).setValue(estatus); 
      return "OK";
    }
  }
  return "No encontrado";
}

// ----------------------------------------------------
// OPERACIONES DE MOVER Y TRANSFORMAR
// ----------------------------------------------------

function transferirProducto(origenId, destinoId, productoId, presentacionId, lote, cantidadMover) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.openById(ID_HOJA);
    const sheet = ss.getSheetByName('INVENTARIO');
    const data = sheet.getDataRange().getValues();
    
    // 1. Origen
    let filaOrigen = -1; let datosOrigen = null;
    for (let i = 1; i < data.length; i++) {
      if (data[i][2] == origenId && String(data[i][6]) == String(lote) && data[i][0] == productoId) {
        filaOrigen = i + 1; datosOrigen = data[i]; break;
      }
    }
    if (filaOrigen === -1) throw new Error("Origen no encontrado");
    
    // 2. Restar
    const stockActual = Number(sheet.getRange(filaOrigen, 4).getValue());
    if (cantidadMover > stockActual + 0.001) throw new Error("Stock insuficiente");

    const nuevoStock = stockActual - cantidadMover;
    if (nuevoStock <= 0.001) sheet.deleteRow(filaOrigen);
    else sheet.getRange(filaOrigen, 4).setValue(nuevoStock);

    // 3. Destino
    const dataNew = sheet.getDataRange().getValues();
    let filaDestino = -1;
    for (let i = 1; i < dataNew.length; i++) {
      if (dataNew[i][2] == destinoId && String(dataNew[i][6]) == String(lote) && dataNew[i][0] == productoId) {
        filaDestino = i + 1; break;
      }
    }

    if (filaDestino > 0) {
      const cell = sheet.getRange(filaDestino, 4);
      cell.setValue(Number(cell.getValue()) + Number(cantidadMover));
    } else {
      sheet.appendRow([
        datosOrigen[0], datosOrigen[1], destinoId, Number(cantidadMover),
        datosOrigen[4], datosOrigen[5], datosOrigen[6], new Date(), "Transferencia"
      ]);
    }
    return true;
  } catch (e) { throw new Error(e.message); } finally { lock.releaseLock(); }
}

function realizarTransformacion(datos) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.openById(ID_HOJA);
    const sheet = ss.getSheetByName('INVENTARIO');
    const data = sheet.getDataRange().getValues();
    
    let filaOrigen = -1; let infoOrigen = null;
    for (let i = 1; i < data.length; i++) {
      if (data[i][2] == datos.origenId && String(data[i][6]) == String(datos.loteOrigen) && data[i][0] == datos.productoIdOrigen) {
        filaOrigen = i + 1; infoOrigen = { cad: data[i][4], elab: data[i][5] }; break;
      }
    }
    if(filaOrigen===-1) throw new Error("Origen no encontrado");
    
    const cellVol = sheet.getRange(filaOrigen, 4);
    const nuevoVol = cellVol.getValue() - datos.cantidadLitros;
    if(nuevoVol <= 0.001) sheet.deleteRow(filaOrigen); else cellVol.setValue(nuevoVol);

    const dataNew = sheet.getDataRange().getValues();
    let filaDestino = -1;
    const loteFinal = datos.nuevoLote || datos.loteOrigen;
    
    for (let i = 1; i < dataNew.length; i++) {
      if (dataNew[i][0] == datos.nuevoProductoId && dataNew[i][1] == datos.nuevaPresentacionId && 
          dataNew[i][2] == datos.origenId && String(dataNew[i][6]) == String(loteFinal)) {
        filaDestino = i + 1; break;
      }
    }

    if (filaDestino > 0) {
      const c = sheet.getRange(filaDestino, 4); c.setValue(Number(c.getValue()) + Number(datos.cantidadLitros));
    } else {
      sheet.appendRow([
        datos.nuevoProductoId, datos.nuevaPresentacionId, datos.origenId, Number(datos.cantidadLitros),
        infoOrigen.cad, infoOrigen.elab, loteFinal, new Date(), "Transformación"
      ]);
    }
    return true;
  } catch (e) { throw new Error(e.message); } finally { lock.releaseLock(); }
}

// ----------------------------------------------------
// CREACIÓN INTELIGENTE (ESTO ARREGLA LOS FUTUROS 0)
// ----------------------------------------------------
function registrarNuevaPresentacion(nombre) {
  const s = obtenerHojaOCrear('PRESENTACIONES', ['ID', 'NOMBRE', 'VOLUMEN']);
  
  // LOGICA INTELIGENTE DE EXTRACCIÓN
  let volumen = 0;
  // Buscamos cualquier número (entero o decimal) en el texto
  const match = String(nombre).match(/[\d\.]+/);
  if (match) {
    volumen = parseFloat(match[0]);
    // Si dice "ml", convertimos a litros
    if (String(nombre).toLowerCase().includes('ml')) {
      volumen = volumen / 1000;
    }
  }
  
  s.appendRow([Utilities.getUuid(), nombre, volumen]); 
  return true;
}

function registrarNuevoProducto(d) { 
  const s = obtenerHojaOCrear('PRODUCTOS', ['ID', 'NOMBRE', 'DESCRIPCION', 'UNIDAD']);
  s.appendRow([Utilities.getUuid(), d.nombre, d.descripcion, d.unidad]); return true;
}
function registrarNuevaUbicacion(d) {
  const s = obtenerHojaOCrear('UBICACIONES', ['ID', 'NOMBRE']);
  s.appendRow([Utilities.getUuid(), d]); return true;
}
function actualizarNombreUbicacion(id, n) { 
  const s = obtenerHojaOCrear('UBICACIONES', []);
  const d = s.getDataRange().getValues();
  for(let i=1; i<d.length; i++) if(d[i][0]==id) { s.getRange(i+1, 2).setValue(n); return true; }
}
function borrarUbicacion(id) {
  const s = obtenerHojaOCrear('UBICACIONES', []);
  const d = s.getDataRange().getValues();
  for(let i=1; i<d.length; i++) if(d[i][0]==id) { s.deleteRow(i+1); return true; }
}

// --- PUENTE PARA EL FRONTEND ---
function registrarSalidas(lista) {
  return registrarSalidaMasiva(lista);
}