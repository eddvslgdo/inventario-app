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
