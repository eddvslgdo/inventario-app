// ==========================================
// 5. LÓGICA DE SALIDAS (FIFO + CHECKOUT)
// ==========================================
function calcularTotalSalida() {
  const total = calcTotalGenerico("sal_presentacion", "sal_cantidad");
  document.getElementById("sal_infoTotal").innerText = total + " L";
  return total;
}

function aplicarFIFO() {
  const selectProd = document.getElementById("sal_producto");
  const prodId = selectProd.value;
  if (
    !prodId ||
    prodId === "Cargando..." ||
    prodId === "" ||
    prodId === "Selecciona..."
  )
    return;

  const card = document.getElementById("vista-salidas").querySelector(".card");
  card.style.opacity = "0.6";
  mostrarToast("🔍 Analizando stock...", "info");
  limpiarCamposSalidaParcial();

  google.script.run
    .withSuccessHandler((respuestaTexto) => {
      card.style.opacity = "1";
      let resultado;
      try {
        resultado = JSON.parse(respuestaTexto);
      } catch (e) {
        return;
      }
      if (!resultado.success) {
        mostrarToast(resultado.error, "error");
        document.getElementById("sal_cantidad").placeholder = "Sin stock";
        return;
      }
      renderizarTablaDisponibilidad(resultado.lista_completa);
      const mejor = resultado.mejor_candidato;
      const indexVisual = resultado.lista_completa.findIndex(
        (l) => l.es_sugerido,
      );
      seleccionarLote(
        mejor.lote,
        mejor.presentacion_id,
        mejor.ubicacion_id,
        mejor.stock_real,
        indexVisual,
      );
      mostrarToast(`Sugerido: ${mejor.lote}`, "success");
    })
    .withFailureHandler((err) => {
      card.style.opacity = "1";
      mostrarError(err);
    })
    .obtenerSugerenciaFIFO(prodId, COLA_SALIDAS);
}

function seleccionarLote(lote, presId, ubicId, stockMax, indexVisual) {
  document.getElementById("sal_lote").value = lote;
  setVal("sal_presentacion", presId);
  setVal("sal_ubicacion", ubicId);
  ["sal_lote", "sal_presentacion", "sal_ubicacion"].forEach((id) => {
    const el = document.getElementById(id);
    el.disabled = true;
    el.style.backgroundColor = "#e9ecef";
  });
  const inputCant = document.getElementById("sal_cantidad");
  inputCant.value = "";
  inputCant.placeholder = `Máx: ${Number(stockMax).toFixed(2)} L`;
  inputCant.setAttribute("data-max-stock", stockMax);
  document.getElementById("sal_infoTotal").innerText = "0 L";

  const filas = document.querySelectorAll(".fila-lote");
  filas.forEach((f) => (f.style.background = "transparent"));
  const filaActiva = document.getElementById("fila-" + indexVisual);
  if (filaActiva) filaActiva.style.background = "#e3f2fd";
}

function setVal(id, val) {
  const el = document.getElementById(id);
  if (!el) return;
  el.value = val;
  if (el.selectedIndex <= 0 && val) {
    const opt = document.createElement("option");
    opt.value = val;
    opt.text = "ID: " + val.substring(0, 8);
    opt.style.color = "red";
    opt.selected = true;
    el.appendChild(opt);
  }
}

function getNombreFromSelect(selectId, valId) {
  const sel = document.getElementById(selectId);
  if (!sel) return valId;
  for (let i = 0; i < sel.options.length; i++) {
    if (sel.options[i].value === valId) return sel.options[i].text;
  }
  return "ID: " + valId.substring(0, 6) + "...";
}

function renderizarTablaDisponibilidad(lista) {
  const div = document.getElementById("div-disponibilidad");
  if (!div) return;
  let html = `
      <div style="background:white; border:1px solid #ddd; border-radius:8px; overflow:hidden;">
        <table style="width:100%; font-size:0.85rem; border-collapse:collapse;">
          <tr style="background:#f8f9fa; text-align:left; color:#666; font-size:0.75rem;">
            <th style="padding:8px;">Lote</th><th style="padding:8px;">Ubic</th><th style="padding:8px;">Pres</th><th style="padding:8px; text-align:right;">Disp.</th>
          </tr>`;
  lista.forEach((l, idx) => {
    const stockNum = Number(l.stock);
    const stockStyle =
      stockNum <= 0
        ? "color:#ccc; text-decoration:line-through;"
        : "color:#2ecc71; font-weight:bold;";
    const cursor = stockNum > 0 ? "cursor:pointer;" : "cursor:not-allowed;";
    const clickAction =
      stockNum > 0
        ? `onclick="seleccionarLote('${l.lote}', '${l.presentacion_id}', '${l.ubicacion_id}', ${l.stock}, ${idx})"`
        : "";
    const nombreUbic = getNombreFromSelect("sal_ubicacion", l.ubicacion_id);
    const nombrePres = getNombreFromSelect(
      "sal_presentacion",
      l.presentacion_id,
    );
    html += `
        <tr id="fila-${idx}" class="fila-lote" style="border-bottom:1px solid #eee; ${cursor}" ${clickAction}>
          <td style="padding:10px 8px; color:#333; font-weight:500;">${l.lote}</td>
          <td style="padding:10px 8px; color:#666;">${nombreUbic}</td>
          <td style="padding:10px 8px; color:#666;">${nombrePres}</td>
          <td style="padding:10px 8px; text-align:right; ${stockStyle}">${l.stock} L</td>
        </tr>`;
  });
  html += `</table></div>`;
  div.innerHTML = html;
  div.style.display = "block";
}

function agregarColaSalida() {
  const vol = calcularTotalSalida();
  if (vol <= 0) return mostrarToast("Cantidad incorrecta", "error");
  const lote = document.getElementById("sal_lote").value.trim();
  if (!lote) return mostrarToast("Falta Lote", "error");

  const item = {
    producto_id: document.getElementById("sal_producto").value,
    nombre_producto: getTextoSelect("sal_producto"),
    presentacion_id: document.getElementById("sal_presentacion").value,
    nombre_presentacion: getTextoSelect("sal_presentacion"),
    ubicacion_id: document.getElementById("sal_ubicacion").value,
    nombre_ubicacion: getTextoSelect("sal_ubicacion"),
    volumen_L: vol,
    piezas: document.getElementById("sal_cantidad").value,
    lote: lote,
  };
  COLA_SALIDAS.push(item);
  renderizarColaSalida();
  limpiarCamposSalidaParcial();
  mostrarToast("Agregado al pedido", "success");
}

function guardarSalidaDirecta() {
  const vol = calcularTotalSalida();
  if (vol <= 0) return mostrarToast("Cantidad incorrecta", "error");
  const lote = document.getElementById("sal_lote").value;
  if (!lote) return mostrarToast("Falta Lote", "error");

  const item = {
    producto_id: document.getElementById("sal_producto").value,
    nombre_producto: getTextoSelect("sal_producto"),
    presentacion_id: document.getElementById("sal_presentacion").value,
    nombre_presentacion: getTextoSelect("sal_presentacion"),
    ubicacion_id: document.getElementById("sal_ubicacion").value,
    nombre_ubicacion: getTextoSelect("sal_ubicacion"),
    volumen_L: vol,
    piezas: document.getElementById("sal_cantidad").value,
    lote: lote,
  };
  // Forzamos "pedido" de 1 solo item, pero abriendo el modal de checkout
  COLA_SALIDAS = [item];
  abrirModalGenerico("modalLogistica");
}

function procesarSalidas() {
  if (COLA_SALIDAS.length === 0) return mostrarToast("Carrito vacío", "error");
  abrirModalGenerico("modalLogistica");
}

function limpiarFormularioSalidaTotal() {
  document.getElementById("sal_producto").value = "";
  limpiarCamposSalidaParcial();
}

function limpiarCamposSalidaParcial() {
  document.getElementById("sal_cantidad").value = "";
  document.getElementById("sal_infoTotal").innerText = "0 L";
  document.getElementById("sal_lote").value = "";
  document.getElementById("div-disponibilidad").style.display = "none";
  ["sal_lote", "sal_presentacion", "sal_ubicacion"].forEach((id) => {
    const el = document.getElementById(id);
    el.value = "";
    el.disabled = false;
    el.style.backgroundColor = "white";
  });
}

// --- LOGÍSTICA / CHECKOUT ---
function alSeleccionarCliente() {
  const id = document.getElementById("log_cliente_select").value;
  const cliente = LISTA_CLIENTES.find((c) => c.id === id);
  if (cliente) {
    document.getElementById("log_nombre").value = cliente.nombre;
    document.getElementById("log_empresa").value = cliente.empresa;
    document.getElementById("log_direccion").value = cliente.direccion;
    document.getElementById("log_telefono").value = cliente.telefono;
    document.getElementById("log_email").value = cliente.email;
    document.getElementById("log_guardar_cliente").checked = false;
  } else {
    document.getElementById("log_nombre").value = "";
    document.getElementById("log_empresa").value = "";
    document.getElementById("log_direccion").value = "";
    document.getElementById("log_telefono").value = "";
    document.getElementById("log_email").value = "";
    document.getElementById("log_guardar_cliente").checked = true;
  }
}

function ejecutarGuardadoMasivo() {
  const btn = document.querySelector("#modalConfirmarMasivo button:last-child");
  btn.innerText = "Procesando...";
  btn.disabled = true;

  if (TIPO_ACCION_MASIVA === "ENTRADA") {
    google.script.run
      .withSuccessHandler(() => {
        mostrarToast("Entradas Guardadas", "success");
        COLA_ENTRADAS = [];
        renderizarColaEntrada();
        limpiarFormularioEntrada();
        recargarSelectores();
        cerrarModalGenerico("modalConfirmarMasivo");
        btn.innerText = "Sí, Guardar";
        btn.disabled = false;
      })
      .registrarEntradaMasiva(COLA_ENTRADAS);
  }
}

function renderizarColaSalida() {
  renderColaGenerica(
    COLA_SALIDAS,
    "area-cola-salidas",
    "tablaColaSalidaBody",
    "contadorColaSal",
    "sal",
  );
}

// --- UTILIDADES ---
function calcTotalGenerico(idPres, idCant) {
  const sel = document.getElementById(idPres);
  const inp = document.getElementById(idCant);
  if (!sel || !inp) return 0;
  const opt = sel.options[sel.selectedIndex];
  const vol = opt ? Number(opt.getAttribute("data-volumen")) : 0;
  const cant = Number(inp.value);
  return vol > 0 && cant > 0 ? Number((cant * vol).toFixed(2)) : 0;
}

function renderColaGenerica(lista, idArea, idBody, idContador, tipo) {
  const area = document.getElementById(idArea);
  const tbody = document.getElementById(idBody);
  document.getElementById(idContador).innerText = lista.length;
  if (lista.length === 0) {
    area.style.display = "none";
    return;
  }
  area.style.display = "block";
  tbody.innerHTML = "";
  lista
    .slice()
    .reverse()
    .forEach((item, idxReal) => {
      const idx = lista.length - 1 - idxReal;
      const tr = document.createElement("tr");
      tr.style.borderBottom = "1px solid #eee";
      const colorLote = tipo === "sal" ? "#d63031" : "#007aff";
      tr.innerHTML = `<td style="padding:8px;"><div style="font-weight:600;">${item.nombre_producto}</div><div style="font-size:0.8rem;color:#666;">${item.nombre_presentacion}</div></td><td style="padding:8px;font-family:monospace;color:${colorLote};">${item.lote}</td><td style="padding:8px;text-align:right;"><strong>${item.volumen_L} L</strong></td><td style="padding:8px;text-align:center;"><button onclick="${tipo === "ent" ? "eliminarEnt" : "eliminarSal"}(${idx})" style="color:red;border:none;background:none;cursor:pointer;">🗑️</button></td>`;
      tbody.appendChild(tr);
    });
}
function eliminarEnt(i) {
  COLA_ENTRADAS.splice(i, 1);
  renderizarColaEntrada();
}
function eliminarSal(i) {
  COLA_SALIDAS.splice(i, 1);
  renderizarColaSalida();
}

// --- MODALES (COMPLETOS CON LÓGICA) ---

function guardarProducto() {
  const d = {
    nombre: document.getElementById("nuevoNombre").value,
    descripcion: document.getElementById("nuevaDescripcion").value,
    unidad: document.getElementById("nuevaUnidad").value,
  };
  if (!d.nombre) return mostrarToast("Falta nombre", "error");
  google.script.run
    .withSuccessHandler(() => {
      recargarSelectores();
      cerrarModal();
      mostrarToast("Guardado", "success");
    })
    .registrarNuevoProducto(d);
}
function guardarPresentacion() {
  const v = document.getElementById("nuevaPresDesc").value;
  if (!v) return mostrarToast("Falta dato", "error");
  google.script.run
    .withSuccessHandler(() => {
      recargarSelectores();
      cerrarModalGenerico("modalPresentacion");
      mostrarToast("Guardado", "success");
    })
    .registrarNuevaPresentacion(v);
}

// TRANSFERENCIA COMPLETA
function abrirTransferencia(ori, lot, stL, volNominal, nom) {
  if (!volNominal || volNominal <= 0)
    return mostrarToast("Error: Producto sin volumen.", "error");
  const stockPiezas = Math.round(stL / volNominal);
  const u = DATOS_UBICACIONES.find((x) => x.id === ori);
  const it = u.items.find((x) => x.lote === lot && x.nombre_completo === nom);

  TRANSF_DATA = {
    origenId: ori,
    productoId: it.raw_producto_id,
    presentacionId: it.raw_presentacion_id,
    lote: lot,
    volumenUnitario: volNominal,
    maxPiezas: stockPiezas,
  };

  document.getElementById("txtProductoTransf").innerText = nom;
  document.getElementById("txtLoteTransf").innerText =
    `Lote: ${lot} | Max: ${stockPiezas} pzas`;
  document.getElementById("cantTransferir").max = stockPiezas;
  document.getElementById("cantTransferir").value = "";
  document.getElementById("lblEquivalencia").innerText = "0 L";
  document.getElementById("lblRestante").innerText = "---";

  const sel = document.getElementById("selDestinoTransf");
  sel.innerHTML = "";
  DATOS_UBICACIONES.forEach((x) => {
    if (x.id !== ori) {
      const o = document.createElement("option");
      o.value = x.id;
      o.text = x.nombre;
      sel.appendChild(o);
    }
  });
  abrirModalGenerico("modalTransferencia");
}
function calcularSimulacionTransferencia() {
  const p = Number(document.getElementById("cantTransferir").value);
  document.getElementById("lblEquivalencia").innerText =
    (p * TRANSF_DATA.volumenUnitario).toFixed(2) + " L";
  const rest = TRANSF_DATA.maxPiezas - p;
  const lblRest = document.getElementById("lblRestante");
  lblRest.innerText = rest + " pzas";
  lblRest.style.color = rest < 0 ? "red" : "#666";
}
function confirmarTransferencia() {
  const p = Number(document.getElementById("cantTransferir").value);
  if (p <= 0 || p > TRANSF_DATA.maxPiezas)
    return mostrarToast("Cantidad inválida", "error");
  const des = document.getElementById("selDestinoTransf").value;
  google.script.run
    .withSuccessHandler(() => {
      mostrarToast("Movido", "success");
      cerrarModalGenerico("modalTransferencia");
      cerrarModalGenerico("modalDetalleUbicacion");
      cargarDashboardUbicaciones();
      recargarSelectores();
    })
    .transferirProducto(
      TRANSF_DATA.origenId,
      des,
      TRANSF_DATA.productoId,
      TRANSF_DATA.presentacionId,
      TRANSF_DATA.lote,
      p * TRANSF_DATA.volumenUnitario,
    );
}

// TRANSFORMACIÓN COMPLETA
function abrirTransformacion(ori, lot, stL, volNominal, nom) {
  if (!volNominal)
    return mostrarToast("Error: Producto origen sin volumen.", "error");
  const u = DATOS_UBICACIONES.find((x) => x.id === ori);
  const it = u.items.find((x) => x.lote === lot && x.nombre_completo === nom);

  DATA_TRANSF = {
    ubicId: ori,
    prodIdOrigen: it.raw_producto_id,
    presIdOrigen: it.raw_presentacion_id,
    loteOrigen: lot,
    volNominalOrigen: volNominal,
    stockMaxPiezas: Math.round(stL / volNominal),
  };

  document.getElementById("lblTransfOrigen").innerText = nom;
  document.getElementById("lblTransfLote").innerText = lot;
  document.getElementById("cantTransfPiezas").max = DATA_TRANSF.stockMaxPiezas;
  document.getElementById("cantTransfPiezas").value = "";
  document.getElementById("selTransfNuevoProd").innerHTML =
    document.getElementById("producto").innerHTML;
  document.getElementById("selTransfNuevaPres").innerHTML =
    document.getElementById("presentacion").innerHTML;

  document.getElementById("lblTransfResultado").innerText = "---";
  document.querySelector(
    "#modalTransformacion .btn-field:last-child",
  ).disabled = false;
  abrirModalGenerico("modalTransformacion");
}
function abrirModalDesdeTransf() {
  cerrarModalGenerico("modalTransformacion");
  abrirModal();
}

function calcTransf() {
  const p = Number(document.getElementById("cantTransfPiezas").value);
  const resEl = document.getElementById("lblTransfResultado");
  const btnTransf = document.querySelector(
    "#modalTransformacion .btn-field:last-child",
  );

  if (p > DATA_TRANSF.stockMaxPiezas) {
    resEl.innerText = "⚠️ Excede existencia";
    resEl.style.color = "red";
    btnTransf.disabled = true;
    return;
  }
  btnTransf.disabled = false;
  const l = p * DATA_TRANSF.volNominalOrigen;
  document.getElementById("lblTransfLitros").innerText = l.toFixed(2) + " L";
  const sel = document.getElementById("selTransfNuevaPres");
  const opt = sel.options[sel.selectedIndex];
  const volD = opt ? Number(opt.getAttribute("data-volumen")) : 0;
  resEl.innerText = volD > 0 ? Math.floor(l / volD) + " Pzas nuevas" : "---";
  resEl.style.color = "#2ecc71";
}

function confirmarTransformacion() {
  const p = Number(document.getElementById("cantTransfPiezas").value);
  if (p <= 0 || p > DATA_TRANSF.stockMaxPiezas)
    return mostrarToast("Error en cantidad", "error");
  const l = p * DATA_TRANSF.volNominalOrigen;
  const d = {
    origenId: DATA_TRANSF.ubicId,
    productoIdOrigen: DATA_TRANSF.prodIdOrigen,
    presIdOrigen: DATA_TRANSF.presIdOrigen,
    loteOrigen: DATA_TRANSF.loteOrigen,
    cantidadLitros: l,
    nuevoProductoId: document.getElementById("selTransfNuevoProd").value,
    nuevaPresentacionId: document.getElementById("selTransfNuevaPres").value,
    nuevoLote:
      document.getElementById("inputTransfNuevoLote").value.trim() ||
      DATA_TRANSF.loteOrigen,
  };
  google.script.run
    .withSuccessHandler(() => {
      mostrarToast("Transformado", "success");
      cerrarModalGenerico("modalTransformacion");
      cerrarModalGenerico("modalDetalleUbicacion");
      cargarDashboardUbicaciones();
      recargarSelectores();
    })
    .realizarTransformacion(d);
}
