function getListaSites() {
  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('nome da aba na planilha');
  const dados = aba.getRange(2, 1, aba.getLastRow() - 1).getValues();
  return dados.map(l => l[0]).filter(site => site); // Remove vazios
}

function getDocumentosDoLote(lote) {
  const todosLotes = getTodosLotes(); // Usa a função que já contém os dados dos lotes
  return todosLotes[lote] || {};
}

function abrirInterface() {
  const html = HtmlService.createHtmlOutputFromFile('form')
    .setWidth(600)
    .setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, 'Gerar Mala Direta');
}
