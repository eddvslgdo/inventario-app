// ==========================================
// 4. LÓGICA DE ENTRADAS
// ==========================================
function calcularTotalEntrada() {
  const total = calcTotalGenerico("presentacion", "cantidad");
  document.getElementById("infoTotal").innerText = total + " L";
  return total;
}

function guardarEntradaDirecta() {
  const item = construirObjetoEntrada();
  if (!item) return;
  const btn = event.target;
  const txt = btn.innerText;
  btn.disabled = true;
  btn.innerText = "Guardando...";
  google.script.run
    .withSuccessHandler(() => {
      mostrarToast("Entrada Directa Guardada ⚡", "success");
      limpiarFormularioEntrada();
      recargarSelectores();
      btn.disabled = false;
      btn.innerText = txt;
    })
    .withFailureHandler((err) => {
      mostrarError(err);
      btn.disabled = false;
      btn.innerText = txt;
    })
    .registrarEntradaUnica(item);
}

function agregarColaEntrada() {
  const item = construirObjetoEntrada();
  if (!item) return;
  COLA_ENTRADAS.push(item);
  renderizarColaEntrada();
  document.getElementById("cantidad").value = "";
  document.getElementById("lote").value = "";
  document.getElementById("infoTotal").innerText = "0 L";
  document.getElementById("lote").focus();
  mostrarToast("Agregado a la lista", "success");
}

function construirObjetoEntrada() {
  const vol = calcularTotalEntrada();
  if (vol <= 0) {
    mostrarToast("Cantidad inválida", "error");
    return null;
  }
  const lote = document.getElementById("lote").value.trim();
  if (!lote) {
    mostrarToast("Falta Lote", "error");
    document.getElementById("lote").focus();
    return null;
  }
  let fc = new Date();
  fc.setFullYear(fc.getFullYear() + 2);
  let fab =
    document.getElementById("elaboracion").value ||
    new Date().toISOString().split("T")[0];
  return {
    producto_id: document.getElementById("producto").value,
    nombre_producto: getTextoSelect("producto"),
    presentacion_id: document.getElementById("presentacion").value,
    nombre_presentacion: getTextoSelect("presentacion"),
    ubicacion_id: document.getElementById("ubicacion").value,
    volumen_L: vol,
    piezas: document.getElementById("cantidad").value,
    lote: lote,
    proveedor: document.getElementById("proveedor").value.trim(),
    fecha_elaboracion: fab,
    fecha_caducidad: fc.toISOString().split("T")[0],
  };
}

function limpiarFormularioEntrada() {
  document.getElementById("cantidad").value = "";
  document.getElementById("lote").value = "";
  document.getElementById("infoTotal").innerText = "0 L";
  document.getElementById("proveedor").value = "";
  document.getElementById("elaboracion").value = "";
  document.getElementById("producto").value = "";
  document.getElementById("presentacion").value = "";
  document.getElementById("ubicacion").value = "";
}

function procesarEntradas() {
  if (COLA_ENTRADAS.length === 0) return;
  TIPO_ACCION_MASIVA = "ENTRADA";
  document.getElementById("lblCantMasiva").innerText =
    COLA_ENTRADAS.length + " partidas (Entrada)";
  abrirModalGenerico("modalConfirmarMasivo");
}

function renderizarColaEntrada() {
  renderColaGenerica(
    COLA_ENTRADAS,
    "area-cola-entradas",
    "tablaColaEntradaBody",
    "contadorColaEnt",
    "ent",
  );
}
