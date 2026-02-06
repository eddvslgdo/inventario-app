// En tu VS Code: Inventario-app/Code.gs

function doGet() {
  return HtmlService.createTemplateFromFile('frontend/index')
      .evaluate()
      .setTitle('Sistema de Inventario')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// backend/Code.gs

function include(filename) {
  try {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
  } catch (e) {
    // CORRECCIÓN: Devolvemos solo código JS, sin etiquetas <script>
    // Esto funciona seguro dentro de tu js.loader
    return 'console.error("❌ NO SE ENCONTRÓ EL ARCHIVO: ' + filename + ' (Verifica el nombre en Apps Script)");';
  }
}