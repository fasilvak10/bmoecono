function include(file) {
  return HtmlService.createHtmlOutputFromFile(file).getContent();
}

function loadMainModalUser() {
  const htmlServ = HtmlService.createTemplateFromFile("main");
  const html = htmlServ.evaluate();
  html.setWidth(1300).setHeight(750);
  const ui = SpreadsheetApp.getUi();
  ui.showModalDialog(html, "🗃️");
}

function onOpen() {
  createMenu();
}

function doGet() {
  const htmlServer = HtmlService.createTemplateFromFile('main');
  const html = htmlServer.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  return html;

}

function obtenerContenidoHTML(page) {
  const contenidoHTML = HtmlService.createHtmlOutputFromFile(page).getContent();
  return contenidoHTML;
}

