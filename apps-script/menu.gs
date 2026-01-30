function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📦 Controle de Estoque")
    .addItem("🔄 Atualizar estoque", "menuAtualizarEstoque")
    .addSeparator()
    .addItem("📥 Importar novos produtos do Bling", "importarProdutosNovos")
    .addSeparator()
    .addItem("🎨 Aplicar design", "aplicarDesignProfissional")
    .addSeparator()
    .addItem("🧹 Limpar mensagens", "limparMensagens")
    .addSeparator()
    .addItem("👁 Mostrar colunas técnicas", "mostrarColunasTecnicas")
    .addItem("🙈 Ocultar colunas técnicas", "ocultarColunasTecnicas")
    .addToUi();
}

function menuAtualizarEstoque() {
  aplicarDesignProfissional();
  processarPlanilha();
  criarOuAtualizarDashboard();
  SpreadsheetApp.getUi().alert("Atualização finalizada com sucesso ✅");
}
