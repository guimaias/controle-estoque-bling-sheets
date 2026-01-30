function buscarProdutosBling() {
  const token = getAccessToken();
  const produtos = [];
  let pagina = 1;
  const limite = 100;

  while (true) {
    const url = `https://www.bling.com.br/Api/v3/produtos?pagina=${pagina}&limite=${limite}`;

    const response = UrlFetchApp.fetch(url, {
      headers: { Authorization: "Bearer " + token },
      muteHttpExceptions: true
    });

    if (response.getResponseCode() !== 200) break;

    const json = JSON.parse(response.getContentText());
    if (!json.data || json.data.length === 0) break;

    produtos.push(...json.data);
    pagina++;
  }

  return produtos;
}

/*********************************************
 * 🧠 PRODUTOS JÁ EXISTENTES NA PLANILHA
 *********************************************/
function obterProdutoIdsExistentes(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return new Set();

  const ids = sheet.getRange(2, 1, lastRow - 1).getValues().flat();
  return new Set(ids.filter(id => id));
}

/*********************************************
 * 📥 IMPORTAR APENAS PRODUTOS NOVOS
 *********************************************/
function importarProdutosNovos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  const produtosBling = buscarProdutosBling();
  const idsExistentes = obterProdutoIdsExistentes(sheet);

  let novos = 0;
  const linhas = [];

  produtosBling.forEach(prod => {
    if (!idsExistentes.has(prod.id)) {
      linhas.push([
        prod.id,                 // Produto ID
        "",                      // Depósito ID
        prod.nome || "",
        prod.codigo || "",
        prod.estoque?.saldo || 0,
        "", "", ""
      ]);
      novos++;
    }
  });

  if (linhas.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, linhas.length, linhas[0].length)
      .setValues(linhas);
  }

  aplicarDesignProfissional()

  SpreadsheetApp.getUi().alert(
    `Importação concluída ✅\nNovos produtos importados: ${novos}`
  );
}
