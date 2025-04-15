function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🔧 Ferramentas Customizadas')
    .addItem('Mala Direta', 'abrirInterface')
    .addItem('Nome que vai aparecer no menu', 'nomeDaFunção')
    .addItem('Nome que vai aparecer no menu', 'nomeDaFunção')

    .addToUi();
}
