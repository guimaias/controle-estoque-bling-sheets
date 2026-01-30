function criarOuAtualizarDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetDados = ss.getActiveSheet();
  const nomeDashboard = "📊 Dashboard";

  let dash = ss.getSheetByName(nomeDashboard);
  if (!dash) {
    dash = ss.insertSheet(nomeDashboard);
  }

  dash.clear();
  dash.setHiddenGridlines(true);

  // Ocultar todas as colunas primeiro
  dash.showColumns(1, dash.getMaxColumns());
  dash.hideColumns(1, 1); // A
  dash.hideColumns(11, dash.getMaxColumns() - 10); // após J

  // Layout base
  dash.setColumnWidths(2, 9, 140); // B até J
  dash.setRowHeights(1, 20, 42);

  const lastRow = sheetDados.getLastRow();

  let total = 0;
  let ok = 0;
  let erro = 0;
  let ultimaAtualizacao = "-";

  if (lastRow >= 2) {
    const status = sheetDados.getRange(2, 7, lastRow - 1).getValues().flat();
    const datas = sheetDados.getRange(2, 9, lastRow - 1).getValues().flat();

    status.forEach((s, i) => {
      if (s) {
        total++;
        if (s === "OK") ok++;
        if (s === "ERRO") erro++;
        if (datas[i]) ultimaAtualizacao = datas[i];
      }
    });
  }

  // Título
  const titulo = dash.getRange("B2:J3").merge();
  titulo
    .setValue("📦 - Painel de Controle de Estoque")
    .setFontSize(18)
    .setFontWeight("bold")
    .setFontFamily("Arial")
    .setHorizontalAlignment("left")
    .setBorder(true, true, true, true, false, false, "#D1D5DB", SpreadsheetApp.BorderStyle.SOLID);

  // Subtítulo
  dash.getRange("B4:J4").merge()
    .setValue("Visão geral das últimas atualizações realizadas")
    .setFontSize(10)
    .setFontColor("#6B7280");

  // Cards
  criarCard(dash, "B6:D9", "TOTAL PROCESSADOS", total, "#F9FAFB", "#111827");
  criarCard(dash, "E6:G9", "SUCESSO (OK)", ok, "#ECFDF5", "#065F46");
  criarCard(dash, "H6:J9", "ERROS", erro, "#FEF2F2", "#991B1B");

  criarCard(dash, "B11:J13", "ÚLTIMA ATUALIZAÇÃO", ultimaAtualizacao, "#EFF6FF", "#1E3A8A");

    // 🔒 Ocultar linhas excedentes (limpar visual do dashboard)
  const ultimaLinhaPainel = 13;
  const totalLinhas = dash.getMaxRows();

  if (totalLinhas > ultimaLinhaPainel) {
    dash.showRows(1, totalLinhas);
    dash.hideRows(ultimaLinhaPainel + 1, totalLinhas - ultimaLinhaPainel);
  }

  if (total === 0) return;


  // 📌 Torna o Dashboard a aba ativa (ESSENCIAL para PDF não sair em branco)
  ss.setActiveSheet(dash);

  // 🔒 Força o Sheets a renderizar tudo
  SpreadsheetApp.flush();

  // ⏳ Aguarda o layout estabilizar
  Utilities.sleep(2000);

// 📄 Gera o PDF do dashboard
gerarPdfDashboard();

}

function gerarPdfDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("📊 Dashboard");
  if (!sheet) return;

  const folderName = "Dashboards Estoque";
  let folder;

  // 🔎 Verifica se a pasta já existe
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) {
    folder = folders.next();
  } else {
    folder = DriveApp.createFolder(folderName);
  }

  const spreadsheetId = ss.getId();
  const sheetId = sheet.getSheetId();

  const url =
    "https://docs.google.com/spreadsheets/d/" +
    spreadsheetId +
    "/export?" +
    "format=pdf" +
    "&size=A4" +
    "&portrait=false" +
    "&fitw=true" +
    "&sheetnames=false" +
    "&printtitle=false" +
    "&pagenumbers=false" +
    "&gridlines=false" +
    "&fzr=false" +
    "&gid=" +
    sheetId;

  const token = ScriptApp.getOAuthToken();

  let response;
try {
  response = UrlFetchApp.fetch(url, {
    headers: {
      Authorization: "Bearer " + token
    },
    muteHttpExceptions: true
  });
} catch (e) {
  Logger.log("Erro ao exportar PDF: " + e.message);
  return;
}

if (response.getResponseCode() !== 200) {
  Logger.log("Falha ao gerar PDF. Código: " + response.getResponseCode());
  Logger.log(response.getContentText());
  return;
}

  const agora = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "yyyy-MM-dd_HH-mm"
  );

  const fileName = `Dashboard_Estoque_${agora}.pdf`;
  const file = folder.createFile(response.getBlob().setName(fileName));

  return file.getUrl(); // caso queira usar depois
}



/******************** CARD ********************/
function criarCard(sheet, range, titulo, valor, corFundo, corTexto) {
  const r = sheet.getRange(range);
  r.merge();
  r.setBackground(corFundo);
  r.setFontFamily("Arial");
  r.setHorizontalAlignment("left");
  r.setVerticalAlignment("middle");
  r.setFontColor(corTexto);

  const texto = `${titulo}\n\n${valor}`;

  r.setRichTextValue(
    SpreadsheetApp.newRichTextValue()
      .setText(texto)
      .setTextStyle(0, titulo.length, SpreadsheetApp.newTextStyle()
        .setFontSize(11)
        .setBold(true)
        .build()
      )
      .setTextStyle(titulo.length + 2, texto.length, SpreadsheetApp.newTextStyle()
        .setFontSize(22)
        .setBold(true)
        .build()
      )
      .build()
  );

  // Bordas suaves e profissionais
  r.setBorder(
    true, true, true, true,
    false, false,
    "#D1D5DB",
    SpreadsheetApp.BorderStyle.SOLID
  );
}
