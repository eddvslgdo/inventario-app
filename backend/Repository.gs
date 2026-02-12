// --- CONFIGURACIÓN ---
var ID_HOJA = '1zCxn5Cvuvfs29Hbpp58W6VCvV6AczGMG1o7CkhS8d2E'; // <--- TU ID

const SHEETS = {
  PRODUCTOS: 'CAT_PRODUCTOS',
  PRESENTACIONES: 'CAT_PRESENTACIONES',
  UBICACIONES: 'CAT_UBICACIONES',
  INVENTARIO: 'INVENTARIO',
  ENTRADAS: 'REGISTROS_ENTRADA',
  SALIDAS: 'REGISTROS_SALIDA'
};

function getDb() {
  return SpreadsheetApp.openById(ID_HOJA);
}

// --- LECTURA GENÉRICA ---
function getDataFromSheet_(sheetName) {
  const ss = getDb();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  
  const range = sheet.getDataRange();
  if (range.isBlank()) return []; 
  
  const data = range.getValues();
  const headers = data.shift();
  
  if (data.length === 0) return [];
  
  return data.map(row => {
    let obj = {};
    headers.forEach((header, index) => {
      const key = header.toString().trim(); 
      if(key) obj[key] = row[index];
    });
    return obj;
  });
}

// --- GETTERS (CATÁLOGOS) ---
function getListaProductos() {
  return getDataFromSheet_(SHEETS.PRODUCTOS)
    .filter(i => i.activo == true) 
    .map(i => ({ id: i.producto_id, nombre: i.nombre }));
}

function getListaPresentaciones() {
  // AHORA ENVIAMOS EL VOLUMEN NOMINAL
  return getDataFromSheet_(SHEETS.PRESENTACIONES)
    .filter(i => i.activo == true)
    .map(i => ({ 
      id: i.presentacion_id, 
      nombre: i.descripcion,
      volumen: Number(i.volumen_nominal_L) // <--- ESTO ES NUEVO
    }));
}

function getListaUbicaciones() {
  return getDataFromSheet_(SHEETS.UBICACIONES)
    .filter(i => i.activo == true)
    .map(i => ({ id: i.ubicacion_id, nombre: i.nombre }));
}

// --- CREACIÓN ---
function crearProductoNuevo(datos) {
  const sheet = getDb().getSheetByName(SHEETS.PRODUCTOS);
  const nuevoId = Utilities.getUuid();
  sheet.appendRow([nuevoId, datos.nombre, datos.descripcion || '', datos.unidad, true]);
  return { id: nuevoId, nombre: datos.nombre };
}

function crearPresentacionNueva(datos) {
  const sheet = getDb().getSheetByName(SHEETS.PRESENTACIONES);
  const nuevoId = Utilities.getUuid();
  // Asumimos 0 si no se especifica, luego podrás editarlo en Sheets si quieres
  sheet.appendRow([nuevoId, datos.descripcion, 0, true]);
  return { id: nuevoId, nombre: datos.descripcion, volumen: 0 };
}

function crearUbicacionNueva(nombre) {
  const sheet = getDb().getSheetByName(SHEETS.UBICACIONES);
  const nuevoId = Utilities.getUuid();
  sheet.appendRow([nuevoId, nombre, true]);
  return { id: nuevoId, nombre: nombre };
}

// --- ESCRITURA ---
function guardarLogEntrada(record) {
  getDb().getSheetByName(SHEETS.ENTRADAS).appendRow([
    new Date(), record.producto_id, record.presentacion_id, 
    record.ubicacion_destino, record.volumen_L, record.lote, 
    record.proveedor, record.comentario
  ]);
}

// ==========================================
// HERRAMIENTA DE DEBUG (USANDO getDb)
// ==========================================
function logToDebug(paso, mensaje, datos) {
  try {
    // CORRECCIÓN 1: Usamos getDb() en lugar de getActiveSpreadsheet()
    const ss = getDb(); 
    let sheet = ss.getSheetByName("DEBUG");
    
    if (!sheet) {
      sheet = ss.insertSheet("DEBUG");
      sheet.appendRow(["TIMESTAMP", "PASO", "MENSAJE", "DATOS (JSON)"]);
    }
    
    sheet.appendRow([
      new Date(),
      paso,
      mensaje,
      JSON.stringify(datos)
    ]);
  } catch (e) {
    // Si falla el log, lo mostramos en la consola de Apps Script
    console.error("Error en logToDebug: " + e.message);
  }
}

// ==========================================
// FUNCIÓN PRINCIPAL DE INVENTARIO (CORREGIDA)
// ==========================================
function actualizarInventarioEntrada(datos) {
  const ss = getDb();
  const sheet = ss.getSheetByName("INVENTARIO");

  // --- 1. LÓGICA DE CÁLCULO DE CADUCIDAD ---
  // Si el frontend no mandó caducidad, la calculamos basada en la elaboración (+2 AÑOS)
  let fechaCalculada = datos.fecha_caducidad;

  if (!fechaCalculada && datos.fecha_elaboracion) {
    try {
      // Convertimos la fecha de texto (yyyy-mm-dd) a objeto fecha
      // Nota: Es mejor manipular el string directamente para evitar problemas de zona horaria
      var partes = String(datos.fecha_elaboracion).split('-'); // [2026, 02, 05]
      
      if (partes.length === 3) {
        var anio = parseInt(partes[0]) + 2; // <--- AQUÍ SUMAMOS 2 AÑOS
        var mes = partes[1];
        var dia = partes[2];
        fechaCalculada = anio + "-" + mes + "-" + dia; // Resultado: 2028-02-05
      }
    } catch (e) {
      console.error("Error calculando fecha caducidad: " + e.message);
    }
  }
  // -----------------------------------------

  // Preparamos los valores asegurando que sean texto (con apóstrofe para que Excel no los cambie)
  const valCad = fechaCalculada ? "'" + fechaCalculada : "'SIN-FECHA";
  const valElab = datos.fecha_elaboracion ? "'" + datos.fecha_elaboracion : "'SIN-FECHA";

  const data = sheet.getDataRange().getValues();
  let fila = -1;
  const loteBusqueda = String(datos.lote).trim().toUpperCase();

  // Buscar si ya existe el lote en esa ubicación
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    // Columna A(0)=Prod, C(2)=Ubic, G(6)=Lote
    if (String(row[0]) == datos.producto_id && 
        String(row[2]) == datos.ubicacion_id && 
        String(row[6]).trim().toUpperCase() == loteBusqueda) {
       fila = i + 1;
       break;
    }
  }

  if (fila > 0) {
     // CASO: EXISTE -> ACTUALIZAR STOCK Y FECHAS
     const celdaVol = sheet.getRange(fila, 4);
     const nuevoVol = Number(celdaVol.getValue()) + Number(datos.volumen_L);
     celdaVol.setValue(nuevoVol);
     
     // Actualizamos las fechas también por si estaban vacías antes
     sheet.getRange(fila, 5).setValue(valCad);  // Columna E (Caducidad)
     sheet.getRange(fila, 6).setValue(valElab); // Columna F (Elaboración)
     sheet.getRange(fila, 8).setValue(new Date()); // Columna H (Ultima act)
     
     return fila;
  } else {
     // CASO: NUEVO -> CREAR FILA
     sheet.appendRow([
        datos.producto_id,
        datos.presentacion_id,
        datos.ubicacion_id,
        Number(datos.volumen_L),
        valCad,   // Columna E (Caducidad Calculada)
        valElab,  // Columna F (Elaboración)
        loteBusqueda,
        new Date(),
        Session.getActiveUser().getEmail()
     ]);
     return sheet.getLastRow();
  }
}

// --- SALIDAS ---
function getInventarioDisponible(prod, pres, ubic) {
  const data = getDb().getSheetByName(SHEETS.INVENTARIO).getDataRange().getValues();
  let lotes = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == prod && data[i][1] == pres && data[i][2] == ubic) {
      lotes.push({
        fila: i + 1, volumen: Number(data[i][3]),
        caducidad: data[i][4], fecha_ingreso: data[i][7], lote: data[i][6]
      });
    }
  }
  return lotes;
}

function actualizarVolumenFila(fila, vol) {
  getDb().getSheetByName(SHEETS.INVENTARIO).getRange(fila, 4).setValue(vol);
}

function borrarFilaInventario(fila) {
  getDb().getSheetByName(SHEETS.INVENTARIO).deleteRow(fila);
}

function guardarLogSalida(record) {
  getDb().getSheetByName(SHEETS.SALIDAS).appendRow([
    new Date(), record.producto_id, record.presentacion_id, 
    record.ubicacion_origen, record.volumen_L, record.lote, 
    record.destino, record.comentario
  ]);
}

// --- DASHBOARD (MAPA) ---
// --- EN backend/Repository.gs ---

function getReporteUbicaciones() {
  const ubicaciones = getDataFromSheet_(SHEETS.UBICACIONES).filter(u => u.activo == true);
  if (ubicaciones.length === 0) return [];

  const inventario = getDataFromSheet_(SHEETS.INVENTARIO);
  
  // MAPAS
  const mapProd = {};
  getDataFromSheet_(SHEETS.PRODUCTOS).forEach(p => mapProd[p.producto_id] = p.nombre);

  // OJO: Aquí guardamos el objeto completo de presentación para sacar el volumen nominal
  const mapPres = {};
  getDataFromSheet_(SHEETS.PRESENTACIONES).forEach(p => {
    mapPres[p.presentacion_id] = {
      nombre: p.descripcion,
      volumen: Number(p.volumen_nominal_L || 0) // <--- IMPORTANTE
    };
  });

  return ubicaciones.map(u => {
    const items = inventario
      .filter(i => i.ubicacion_id == u.ubicacion_id)
      .map(i => {
        let caducidad = i.fecha_caducidad;
        if (caducidad instanceof Date) caducidad = caducidad.toLocaleDateString();
        
        const nombreProd = mapProd[i.producto_id] || 'Desc. Borrado';
        const infoPres = mapPres[i.presentacion_id] || { nombre: '', volumen: 0 };
        const nombreCompleto = infoPres.nombre ? `${nombreProd} (${infoPres.nombre})` : nombreProd;

        return {
          producto: nombreProd,
          presentacion: infoPres.nombre,
          nombre_completo: nombreCompleto,
          
          // DATOS OCULTOS PARA LÓGICA
          raw_producto_id: i.producto_id,
          raw_presentacion_id: i.presentacion_id,
          volumen_nominal: infoPres.volumen, // <--- NECESARIO PARA CALCULAR PIEZAS
          
          volumen: Number(i.volumen_actual_L || 0),
          lote: i.lote_referencia || 'S/L',
          caducidad: caducidad,
          proveedor: i.proveedor || ''
        };
      });
      
    return {
      id: u.ubicacion_id, 
      nombre: u.nombre, 
      items: items,
      totalVolumen: items.reduce((s, i) => s + i.volumen, 0)
    };
  });
}

function editarUbicacion(id, nombre) {
  const sheet = getDb().getSheetByName(SHEETS.UBICACIONES);
  const data = sheet.getDataRange().getValues();
  for(let i=1; i<data.length; i++) {
    if(data[i][0] == id) { sheet.getRange(i+1, 2).setValue(nombre); return true; }
  }
}

function eliminarUbicacion(id) {
  const inv = getDb().getSheetByName(SHEETS.INVENTARIO).getDataRange().getValues();
  for(let i=1; i<inv.length; i++) {
    if(inv[i][2] == id && inv[i][3] > 0) throw new Error("La ubicación tiene stock");
  }
  const sheet = getDb().getSheetByName(SHEETS.UBICACIONES);
  const data = sheet.getDataRange().getValues();
  for(let i=1; i<data.length; i++) {
    if(data[i][0] == id) { sheet.getRange(i+1, 3).setValue(false); return true; }
  }
}

// --- FUNCIÓN DE REUBICACIÓN Y BORRADO ---
function reubicarYEliminar(idOrigen, idDestino) {
  const ss = getDb();
  
  // 1. MOVER EL INVENTARIO
  const sheetInv = ss.getSheetByName(SHEETS.INVENTARIO);
  const dataInv = sheetInv.getDataRange().getValues();
  
  // Recorremos buscando filas que tengan la ubicación de origen (columna C -> índice 2)
  for (let i = 1; i < dataInv.length; i++) {
    if (String(dataInv[i][2]) === String(idOrigen)) {
      // Escribimos el ID del destino en la columna 3 (índice 2+1 en notación R1C1)
      sheetInv.getRange(i + 1, 3).setValue(idDestino);
    }
  }

  // 2. DESACTIVAR LA UBICACIÓN ORIGINAL
  const sheetUbi = ss.getSheetByName(SHEETS.UBICACIONES);
  const dataUbi = sheetUbi.getDataRange().getValues();
  
  for (let i = 1; i < dataUbi.length; i++) {
    if (String(dataUbi[i][0]) === String(idOrigen)) {
      // Columna C es "activo" (índice 2) -> Ponemos false
      sheetUbi.getRange(i + 1, 3).setValue(false);
      return true;
    }
  }
}


// --- REPORTE POR PRODUCTOS (VISTA DE COMPRAS) ---
function getReportePorProductos() {
  const inventario = getDataFromSheet_(SHEETS.INVENTARIO);
  if (inventario.length === 0) return [];

  // Mapas de ayuda
  const mapProd = {};
  getDataFromSheet_(SHEETS.PRODUCTOS).forEach(p => mapProd[p.producto_id] = { nombre: p.nombre, unidad: p.unidad });

  const mapPres = {};
  getDataFromSheet_(SHEETS.PRESENTACIONES).forEach(p => mapPres[p.presentacion_id] = p.descripcion);

  const mapUbic = {};
  getDataFromSheet_(SHEETS.UBICACIONES).forEach(u => mapUbic[u.ubicacion_id] = u.nombre);

  // Agrupamiento
  const reporte = {};

  inventario.forEach(item => {
    const pId = item.producto_id;
    
    // Si es la primera vez que vemos este producto, lo inicializamos
    if (!reporte[pId]) {
      const infoProd = mapProd[pId] || { nombre: 'Desconocido', unidad: '?' };
      reporte[pId] = {
        id: pId,
        nombre: infoProd.nombre,
        unidad: infoProd.unidad,
        totalVolumen: 0,
        lotes: [] // Aquí guardaremos el detalle
      };
    }

    // Sumamos al total global
    const vol = Number(item.volumen_actual_L || 0);
    reporte[pId].totalVolumen += vol;

    // Agregamos el detalle del lote
    reporte[pId].lotes.push({
      lote: item.lote_referencia,
      presentacion: mapPres[item.presentacion_id] || 'Granel',
      ubicacion: mapUbic[item.ubicacion_id] || 'Perdido',
      caducidad: item.fecha_caducidad instanceof Date ? item.fecha_caducidad.toLocaleDateString() : item.fecha_caducidad,
      volumen: vol
    });
  });

  // Convertimos el objeto a array y ordenamos por nombre
  return Object.values(reporte).sort((a, b) => a.nombre.localeCompare(b.nombre));
}
