function mostrarToast(mensaje, tipo = "info") {
  const toast = document.getElementById("toast-container");
  if (!toast) return alert(mensaje);
  toast.className = "show " + tipo;
  toast.innerText = mensaje;
  setTimeout(() => {
    toast.className = toast.className.replace("show", "");
    setTimeout(() => {
      toast.className = "";
    }, 500);
  }, 3000);
}

function mostrarError(err) {
  mostrarToast("Error: " + err.message, "error");
}

function abrirModalGenerico(id) {
  const el = document.getElementById(id);
  if (el) el.style.display = "flex";
}
function cerrarModalGenerico(id) {
  const el = document.getElementById(id);
  if (el) el.style.display = "none";
}
function abrirModal() {
  abrirModalGenerico("modalProducto");
}
function cerrarModal() {
  cerrarModalGenerico("modalProducto");
}
