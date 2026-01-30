const BLING_ESTOQUE_URL = "https://www.bling.com.br/Api/v3/estoques";

function atualizarEstoque(produtoId, depositoId, quantidade) {
  const token = getAccessToken();

  const payload = {
    produto: { id: Number(produtoId) },
    deposito: { id: Number(depositoId) },
    quantidade: Number(quantidade),
    tipoOperacao: "B",
    observacoes: "Ajuste automático via Google Sheets"
  };

  const response = UrlFetchApp.fetch(BLING_ESTOQUE_URL, {
    method: "post",
    headers: {
      Authorization: "Bearer " + token,
      "Content-Type": "application/json"
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  return {
    statusCode: response.getResponseCode()
  };
}

function processarPlanilha() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  for (let i = 2; i <= lastRow; i++) {
    const produtoId = sheet.getRange(i, 1).getValue();
    const depositoId = sheet.getRange(i, 2).getValue();
    const quantidade = sheet.getRange(i, 5).getValue();

    if (!produtoId || !depositoId || quantidade === "") continue;

    try {
      const resp = atualizarEstoque(produtoId, depositoId, quantidade);

      if (resp.statusCode === 200 || resp.statusCode === 201) {
        sheet.getRange(i, 7).setValue("OK");
        sheet.getRange(i, 8).setValue("✅ Estoque atualizado com sucesso!");
      } else {
        sheet.getRange(i, 7).setValue("ERRO");
        sheet.getRange(i, 8).setValue("⚠️ Erro ao atualizar estoque");
      }

      sheet.getRange(i, 9).setValue(new Date());
      Utilities.sleep(500);

    } catch (e) {
      sheet.getRange(i, 7).setValue("ERRO");
      sheet.getRange(i, 8).setValue("❌ " + e.message);
      sheet.getRange(i, 9).setValue(new Date());
    }
  }
}
