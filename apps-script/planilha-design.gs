function aplicarDesignProfissional() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  // Fonte padrão
  sheet.getRange(1, 1, lastRow, lastCol)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setFontFamily("Arial")
    .setFontSize(10);

  // Cabeçalho
  sheet.getRange(1, 1, 1, lastCol)
    .setBackground("#1F2937") // cinza escuro profissional
    .setFontColor("#FFFFFF")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

  // Congelar cabeçalho
  sheet.setFrozenRows(1);

  // Ajustar largura automática
  sheet.autoResizeColumns(1, lastCol);

  // Alinhamento padrão
  sheet.getRange(2, 1, lastRow, lastCol)
    .setVerticalAlignment("middle");

  // Status (coluna 7)
  const statusRange = sheet.getRange(2, 7, lastRow);
  statusRange.setHorizontalAlignment("center");

  // Formatação condicional - STATUS
  const rules = sheet.getConditionalFormatRules();

  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("OK")
      .setBackground("#DCFCE7") // verde claro
      .setFontColor("#166534")
      .setRanges([statusRange])
      .build()
  );

  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("ERRO")
      .setBackground("#FEE2E2") // vermelho claro
      .setFontColor("#991B1B")
      .setRanges([statusRange])
      .build()
  );

  sheet.setConditionalFormatRules(rules);

  // Mensagem API (coluna 8)
  sheet.getRange(2, 8, lastRow)
    .setWrap(true);

  // Data atualização (coluna 9)
  sheet.getRange(2, 9, lastRow)
    .setNumberFormat("dd/MM/yyyy HH:mm:ss");

  // Ocultar colunas técnicas
  sheet.hideColumns(1); // Produto ID
  sheet.hideColumns(2); // Depósito ID
  sheet.hideColumns(6); // Nome do depósito
