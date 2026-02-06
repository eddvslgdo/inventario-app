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

function confirmarPedidoCompleto() {
  const nombre = document.getElementById("log_nombre").value.trim();
  if (!nombre) return mostrarToast("Falta nombre", "error");

  const btn = event.target;
  btn.disabled = true;
  btn.innerText = "Procesando...";

  const datosLogistica = {
    idCliente: document.getElementById("log_cliente_select").value || "nuevo",
    nombreCliente: nombre,
    empresa: document.getElementById("log_empresa").value,
    direccion: document.getElementById("log_direccion").value,
    telefono: document.getElementById("log_telefono").value,
    email: document.getElementById("log_email").value,
    tipoEnvio: document.getElementById("log_tipo_envio").value,
    paqueteria: document.getElementById("log_paqueteria").value,
    guia: document.getElementById("log_guia").value,
    costoEnvio: document.getElementById("log_costo").value || 0,
  };
  const guardarCli = document.getElementById("log_guardar_cliente").checked;

  google.script.run
    .withSuccessHandler((res) => {
      mostrarToast(`Pedido Generado: ${res.idPedido} 🚀`, "success");
      COLA_SALIDAS = [];
      renderizarColaSalida();
      limpiarFormularioSalidaTotal();
      cerrarModalGenerico("modalLogistica");
      recargarSelectores(); // Recarga y limpia ceros
      btn.disabled = false;
      btn.innerText = "✅ CONFIRMAR ENVÍO";
    })
    .withFailureHandler((err) => {
      mostrarError(err);
      btn.disabled = false;
      btn.innerText = "✅ CONFIRMAR ENVÍO";
    })
    .procesarPedidoCompleto(datosLogistica, COLA_SALIDAS, guardarCli);
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

// --- DASHBOARD UBICACIONES (FILTRADO 0 Y RESUMEN) ---
function cargarDashboardUbicaciones() {
  const grid = document.getElementById("grid-ubicaciones");
  grid.innerHTML = "";
  google.script.run
    .withSuccessHandler((u) => {
      DATOS_UBICACIONES = u;
      renderizarUbicaciones(u);
    })
    .obtenerDatosUbicaciones();
}
function renderizarUbicaciones(lista) {
  const grid = document.getElementById("grid-ubicaciones");
  const activas = lista.filter((u) => u.totalVolumen > 0.001);
  if (activas.length === 0) {
    grid.innerHTML = "<p>Almacén vacío.</p>";
    return;
  }
  activas.forEach((u) => {
    const card = document.createElement("div");
    card.className = "ubicacion-card";
    card.onclick = function (e) {
      if (e.target.tagName !== "BUTTON") verDetalle(u.id);
    };
    const total = u.totalVolumen.toFixed(2);
    const resumen = {};
    u.items.forEach((i) => {
      if (!resumen[i.producto]) resumen[i.producto] = 0;
      resumen[i.producto] += i.volumen;
    });
    let htmlResumen =
      '<ul style="margin:10px 0; padding-left:20px; font-size:0.85rem; color:#555;">';
    Object.keys(resumen)
      .slice(0, 3)
      .forEach((p) => {
        htmlResumen += `<li><strong>${p}:</strong> ${resumen[p].toFixed(2)} L</li>`;
      });
    htmlResumen += "</ul>";
    card.innerHTML = `<div class="ubicacion-header"><strong>${u.nombre}</strong><span class="badge">${total} L</span></div>${htmlResumen}<div class="actions-row"><button class="btn-icon" onclick="editarUbicacion('${u.id}','${u.nombre}')">✏️</button><button class="btn-icon" style="background:#ff3b30;" onclick="borrarUbicacion('${u.id}')">🗑️</button></div>`;
    grid.appendChild(card);
  });
}

// --- VISTA PRODUCTOS (FILTRADO 0) ---
function cargarDashboardProductos() {
  const tbody = document.getElementById("tablaProductosBody");
  tbody.innerHTML = "";
  google.script.run
    .withSuccessHandler((p) => {
      renderizarTablaProductos(p);
    })
    .obtenerDatosProductos();
}
function renderizarTablaProductos(productos) {
  const tbody = document.getElementById("tablaProductosBody");
  const activos = productos.filter((p) => p.totalVolumen > 0.001);
  if (activos.length === 0) {
    tbody.innerHTML = "<tr><td>Vacío</td></tr>";
    return;
  }
  activos.forEach((p) => {
    const tr = document.createElement("tr");
    tr.style.borderBottom = "1px solid #eee";
    tr.innerHTML = `<td style="padding:15px; width:50%;"><strong>${p.nombre}</strong><br><span style="font-size:0.8rem;color:#888">${p.lotes.length} lotes</span></td><td style="padding:15px; width:30%; text-align:right; font-weight:bold; color:#007aff;">${p.totalVolumen.toFixed(2)} L</td><td style="padding:15px; width:20%; text-align:center;"><button onclick="toggleDetalleProducto('${p.id}')" class="btn-field" style="height:35px;font-size:0.8rem;background:#f0f0f5;color:#333;width:auto;padding:0 15px;border:none;">Ver lotes</button></td>`;
    tbody.appendChild(tr);
    const trD = document.createElement("tr");
    trD.id = "detalle-" + p.id;
    trD.style.display = "none";
    trD.style.background = "#f9f9f9";
    let html =
      '<div style="padding:10px;"><table style="width:100%;font-size:0.8rem;"><thead><tr><th>Ubic</th><th>Pres</th><th>Lote</th><th>Cad</th><th>Cant</th></tr></thead><tbody>';
    p.lotes.forEach((l) => {
      html += `<tr><td>${l.ubicacion}</td><td>${l.presentacion}</td><td style="color:#d63031;font-family:monospace;">${l.lote}</td><td>${l.caducidad || "-"}</td><td><strong>${l.volumen}</strong></td></tr>`;
    });
    html += "</tbody></table></div>";
    trD.innerHTML = `<td colspan="3" style="padding:0">${html}</td>`;
    tbody.appendChild(trD);
  });
}
function toggleDetalleProducto(id) {
  const el = document.getElementById("detalle-" + id);
  el.style.display = el.style.display === "none" ? "table-row" : "none";
}
function filtrarProductos() {
  const txt = document.getElementById("buscadorProductos").value.toLowerCase();
  document
    .querySelectorAll('#tablaProductosBody > tr:not([id^="detalle-"])')
    .forEach((tr) => {
      tr.style.display = tr.innerText.toLowerCase().includes(txt) ? "" : "none";
    });
}

// --- VISTA PEDIDOS (ACTUALIZADA) ---
function cargarDashboardPedidos() {
  const tbody = document.getElementById("tablaPedidosBody");
  tbody.innerHTML =
    '<tr><td colspan="5" style="text-align:center; padding:20px;">Cargando...</td></tr>';

  google.script.run
    .withSuccessHandler((lista) => {
      tbody.innerHTML = "";
      if (lista.length === 0) {
        tbody.innerHTML =
          '<tr><td colspan="5" style="text-align:center; padding:20px;">Sin envíos registrados.</td></tr>';
        return;
      }

      lista.forEach((p) => {
        const tr = document.createElement("tr");
        tr.style.borderBottom = "1px solid #eee";

        let statusColor = "#95a5a6"; // Gris por defecto
        if (p.estatus === "Pendiente") statusColor = "#f1c40f"; // Amarillo
        if (p.estatus === "En Camino") statusColor = "#3498db"; // Azul
        if (p.estatus === "Entregado") statusColor = "#2ecc71"; // Verde
        if (p.estatus === "Cancelado") statusColor = "#e74c3c"; // Rojo

        tr.innerHTML = `
           <td style="padding:12px;">
             <div style="font-weight:bold; color:#007aff; font-size:0.9rem;">${p.id}</div>
             <div style="font-size:0.75rem; color:#888;">${p.fecha}</div>
           </td>
           <td style="padding:12px;">
             <div style="font-weight:500;">${p.cliente}</div>
             <div style="font-size:0.75rem; color:#888;">${p.destino ? p.destino.substring(0, 20) + "..." : "-"}</div>
           </td>
           <td style="padding:12px; text-align:center;">
             <span style="background:${statusColor}; color:white; padding:4px 8px; border-radius:12px; font-size:0.75rem; font-weight:bold;">${p.estatus}</span>
           </td>
           <td style="padding:12px; text-align:right;">
             <button onclick="verDetallePedido('${p.id}')" class="btn-field" style="width:auto; padding:0 12px; height:32px; font-size:0.8rem; background:#f1f3f5; color:#333; border:none;">👁️ Ver</button>
           </td>
         `;
        tbody.appendChild(tr);
      });
    })
    .obtenerHistorialPedidos();
}

// --- DETALLE Y ACTUALIZACIÓN DE PEDIDOS ---
let PEDIDO_ACTUAL_ID = null;

function verDetallePedido(idPedido) {
  PEDIDO_ACTUAL_ID = idPedido;
  document.getElementById("tituloDetallePedido").innerText = "Cargando...";
  abrirModalGenerico("modalDetallePedido");

  google.script.run
    .withSuccessHandler((data) => {
      const cab = data.cabecera;
      document.getElementById("tituloDetallePedido").innerText =
        "Pedido: " + cab.id;
      document.getElementById("det_cliente").innerText = cab.cliente;
      document.getElementById("det_direccion").innerText = cab.direccion;
      document.getElementById("det_guia").innerText =
        `${cab.paqueteria} - ${cab.guia}`;

      // Fechas
      document.getElementById("det_fecha_est").value = cab.fechaEst;
      document.getElementById("det_fecha_real").value = cab.fechaReal;
      document.getElementById("det_estatus").value = cab.estatus || "Pendiente";

      // Tabla Items
      const tbody = document.getElementById("tablaItemsPedido");
      tbody.innerHTML = "";
      data.items.forEach((item) => {
        tbody.innerHTML += `
            <tr style="border-bottom:1px solid #eee;">
              <td style="padding:8px;">${item.producto}<br><small style="color:#888">${item.presentacion}</small></td>
              <td style="padding:8px; font-family:monospace; color:#d63031;">${item.lote}</td>
              <td style="padding:8px; text-align:right;"><strong>${item.volumen} L</strong><br><small>(${item.cantidad} pzs)</small></td>
            </tr>`;
      });
    })
    .obtenerDetallePedidoCompleto(idPedido);
}

function guardarActualizacionPedido() {
  if (!PEDIDO_ACTUAL_ID) return;

  const fEst = document.getElementById("det_fecha_est").value;
  const fReal = document.getElementById("det_fecha_real").value;
  const st = document.getElementById("det_estatus").value;

  const btn = event.target;
  btn.disabled = true;
  btn.innerText = "Guardando...";

  google.script.run
    .withSuccessHandler(() => {
      mostrarToast("Pedido actualizado ✅", "success");
      cerrarModalGenerico("modalDetallePedido");
      cargarDashboardPedidos(); // Refrescar lista
      btn.disabled = false;
      btn.innerText = "💾 Actualizar Seguimiento";
    })
    .actualizarPedido(PEDIDO_ACTUAL_ID, fEst, fReal, st);
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
function guardarUbicacion() {
  const v = document.getElementById("nuevaUbicNombre").value;
  if (!v) return mostrarToast("Falta dato", "error");
  google.script.run
    .withSuccessHandler(() => {
      recargarSelectores();
      cargarDashboardUbicaciones();
      cerrarModalGenerico("modalUbicacion");
      mostrarToast("Guardado", "success");
    })
    .registrarNuevaUbicacion(v);
}

function verDetalle(id) {
  const u = DATOS_UBICACIONES.find((x) => x.id === id);
  if (!u) return;
  document.getElementById("tituloDetalle").innerText = "Contenido: " + u.nombre;
  const tb = document.getElementById("tablaDetalleBody");
  tb.innerHTML = "";
  if (u.items.length === 0)
    tb.innerHTML =
      '<tr><td colspan="5" style="text-align:center">Vacío</td></tr>';
  else {
    u.items.forEach((i) => {
      const tr = document.createElement("tr");
      tr.style.borderBottom = "1px solid #eee";
      tr.innerHTML = `<td>${i.producto}<br><small>${i.presentacion}</small></td><td style="color:#d63031">${i.lote}</td><td>${i.caducidad || "-"}</td><td style="text-align:right"><strong>${i.volumen} L</strong></td><td style="text-align:center"><button onclick="abrirTransferencia('${u.id}','${i.lote}',${i.volumen},${i.volumen_nominal},'${i.nombre_completo}')" class="btn-action btn-move">🔄</button> <button onclick="abrirTransformacion('${u.id}','${i.lote}',${i.volumen},${i.volumen_nominal},'${i.nombre_completo}')" class="btn-action btn-transf">🛠️</button></td>`;
      tb.appendChild(tr);
    });
  }
  abrirModalGenerico("modalDetalleUbicacion");
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

function borrarUbicacion(id) {
  const u = DATOS_UBICACIONES.find((x) => x.id === id);
  if (u.totalVolumen > 0) {
    ID_UBICACION_A_BORRAR = id;
    const s = document.getElementById("selectDestinoReubicacion");
    s.innerHTML = "";
    DATOS_UBICACIONES.forEach((x) => {
      if (x.id !== id) {
        const o = document.createElement("option");
        o.value = x.id;
        o.text = x.nombre;
        s.appendChild(o);
      }
    });
    abrirModalGenerico("modalReubicacion");
  } else {
    if (confirm("¿Eliminar?"))
      google.script.run
        .withSuccessHandler(() => {
          mostrarToast("Eliminado", "info");
          cargarDashboardUbicaciones();
        })
        .borrarUbicacion(id);
  }
}
function confirmarReubicacion() {
  google.script.run
    .withSuccessHandler(() => {
      mostrarToast("Hecho", "success");
      cerrarModalGenerico("modalReubicacion");
      cargarDashboardUbicaciones();
    })
    .procesarReubicacion(
      ID_UBICACION_A_BORRAR,
      document.getElementById("selectDestinoReubicacion").value,
    );
}
function editarUbicacion(id, nom) {
  ID_UBICACION_A_EDITAR = id;
  document.getElementById("inputEditarNombre").value = nom;
  abrirModalGenerico("modalEditarUbicacion");
}
function guardarEdicionUbicacion() {
  google.script.run
    .withSuccessHandler(() => {
      mostrarToast("Editado", "success");
      cerrarModalGenerico("modalEditarUbicacion");
      cargarDashboardUbicaciones();
    })
    .actualizarNombreUbicacion(
      ID_UBICACION_A_EDITAR,
      document.getElementById("inputEditarNombre").value,
    );
}
