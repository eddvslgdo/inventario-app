function doGet(e) {
  const template = HtmlService.createTemplateFromFile('frontend/index');
  
  // Atrapamos si alguien entró escaneando un QR
  template.qrUbicacionId = (e && e.parameter && e.parameter.ubicacion) ? e.parameter.ubicacion : null;
  
  // EXTRAEMOS LA URL OFICIAL DEL SISTEMA DESDE EL SERVIDOR
  template.appUrl = ScriptApp.getService().getUrl(); 

  return template.evaluate()
      .setTitle('Sistema de Inventario')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// backend/Code.gs

function include(filename) {
  try {
    // CAMBIO IMPORTANTE: Usamos createTemplateFromFile y .evaluate()
    // Esto permite que los <?!= include ?> dentro de tus archivos se ejecuten.
    return HtmlService.createTemplateFromFile(filename)
      .evaluate()
      .getContent();
      
  } catch (e) {
    return 'console.error("❌ Error al cargar ' + filename + ': ' + e.message + '");';
  }
}

function getLogoBase64() {
  const fileId = "1jwrdSOB9rA62BegzV5sMv_4KTxO4gtjt";
  const file = DriveApp.getFileById(fileId);
  const blob = file.getBlob();
  const base64 = Utilities.base64Encode(blob.getBytes());
  
  return {
    data: base64,
    type: blob.getContentType()
  };
}