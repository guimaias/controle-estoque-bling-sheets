function aplicarDesignProfissional() { ... }

function protegerColunasSistema() { ... }

function limparMensagens() { ... }

function mostrarColunasTecnicas() {
  const sheet = SpreadsheetApp.getActiveSheet();
  sheet.showColumns(1, 2);
  sheet.showColumns(6);
}

function ocultarColunasTecnicas() {
  const sheet = SpreadsheetApp.getActiveSheet();
  sheet.hideColumns(1);
  sheet.hideColumns(2);
  sheet.hideColumns(6);
}
