function doGet() {
  // Renderiza la plantilla principal
return HtmlService.createTemplateFromFile('frontend/index')
    .evaluate()
    .setTitle('Sistema de Inventario')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Función auxiliar para incluir archivos HTML dentro de otros (CSS/JS)
 * Se usa en el HTML así: <?!= include('frontend/css'); ?>
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}