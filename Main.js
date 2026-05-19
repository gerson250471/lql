/**
 * ARQUIVO: Main.js
 * Ponto de entrada da aplicação Web (Web App)
 */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('CRM LQL Consig')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}