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

function editarUbicacion(id, nom) {
  ID_UBICACION_A_EDITAR = id;
  document.getElementById("inputEditarNombre").value = nom;
  abrirModalGenerico("modalEditarUbicacion");
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
