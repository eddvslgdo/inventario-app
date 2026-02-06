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
