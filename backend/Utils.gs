function ss() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function hoja(nombre) {
  return ss().getSheetByName(nombre);
}

function ahora() {
  return new Date();
}

// ==========================================
// CONTROL DE ENTORNOS (PROD / TEST)
// ==========================================
const DB_PROD_ID = "1zCxn5Cvuvfs29Hbpp58W6VCvV6AczGMG1o7CkhS8d2E";
const DB_TEST_ID = "1N7ofFjp98-B3-QvTBNEUzeA7G0V6rMVQVPlEyZ_ZJT4";

function getActiveDbId() {
  // Por defecto apunta a PROD si no hay nada configurado
  const env = PropertiesService.getUserProperties().getProperty('ENV_MODE') || 'PROD';
  return env === 'TEST' ? DB_TEST_ID : DB_PROD_ID;
}

function switchEnvironment(mode) {
  if(mode !== 'PROD' && mode !== 'TEST') return;
  PropertiesService.getUserProperties().setProperty('ENV_MODE', mode);
  return mode;
}

function getCurrentEnvironment() {
  return PropertiesService.getUserProperties().getProperty('ENV_MODE') || 'PROD';
}