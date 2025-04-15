function abrirInterface() {
  const html = HtmlService.createHtmlOutputFromFile('Index')
    .setWidth(300)
    .setHeight(250);
  SpreadsheetApp.getUi().showModalDialog(html, 'Gerador de Mala Direta');
}

function obterDadosInterface() {

/*Estou usando o APP Script que já esta relacionado com a planilha que eu estou usando, se você quiser usar um arquivo independente terá quemodificar e fazer o relacionamento com a planilha que irá rodar esse código*/
  const planilha = SpreadsheetApp.getActiveSpreadsheet();

/*Aqui você vai colocar o nome da aba que você que usar para fazer a busca. Exemplo, possui uma planilha chamada LEADS e dentro dessa planilha existe a Aba CONTATOS, aó você coloca o nome CONTATO*/
  const aba = planilha.getSheetByName('nome da aba na planilha');
  if (!aba) throw new Error('Aba "nome da aba na planilha" não encontrada.');
  const dados = aba.getDataRange().getValues();
  const cabecalhos = dados[0];
  const sites = dados.slice(1).map(l => l[0]);
/* site é o nome da aba que esou usando como referencia para as linhas */
  return { sites, cabecalhos };
}

function getTodosLotes() {
  return {
/*Substitua pelo o que você achar melhor*/
    "Lote_1_de_arquivos_para_modificar": {
      'ID DO ARQUIVO': 'NOVO NOME DO ARQUIVO QUE VAI APARECER NA LISTA SUSPENSA',
      'ID DO ARQUIVO': 'NOVO NOME DO ARQUIVO QUE VAI APARECER NA LISTA SUSPENSA',
      'ID DO ARQUIVO': 'NOVO NOME DO ARQUIVO QUE VAI APARECER NA LISTA SUSPENSA',
    },
    "Lote_2_de_arquivos_para_modificar": {
      'ID DO ARQUIVO': 'NOVO NOME DO ARQUIVO QUE VAI APARECER NA LISTA SUSPENSA',
      'ID DO ARQUIVO': 'NOVO NOME DO ARQUIVO QUE VAI APARECER NA LISTA SUSPENSA',
      'ID DO ARQUIVO': 'NOVO NOME DO ARQUIVO QUE VAI APARECER NA LISTA SUSPENSA',
    },
    "Lote_3_de_arquivos_para_modificar": {
      'ID DO ARQUIVO': 'NOVO NOME DO ARQUIVO QUE VAI APARECER NA LISTA SUSPENSA',
      'ID DO ARQUIVO': 'NOVO NOME DO ARQUIVO QUE VAI APARECER NA LISTA SUSPENSA',
      'ID DO ARQUIVO': 'NOVO NOME DO ARQUIVO QUE VAI APARECER NA LISTA SUSPENSA',
    },


  };
}

function gerarDocumentos(siteSelecionado, loteSelecionado, documentosSelecionados) {
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const aba = planilha.getSheetByName('nome da aba na planilha');
  const dados = aba.getDataRange().getValues();
  const cabecalhos = dados[0];
  const linha = dados.find((r, idx) => idx > 0 && r[0] === siteSelecionado);
  if (!linha) throw new Error("Site não encontrado");

  const todosLotes = getTodosLotes();
  const docModelos = todosLotes[loteSelecionado] || {};

  function formatarData(valor) {
    return valor instanceof Date
      ? valor.toLocaleDateString('pt-BR', { day: 'numeric', month: 'long', year: 'numeric' })
      : valor;
  }

  // =====================
  // BLOCO DE CRIAÇÃO DE PASTAS (fora do loop)
  // =====================
//Aqui ele buscando a pasta chamada de "CONTRATOS 'Nome_Contido_na_coluna_20'" e busca o nome da pasta de acordo com o nome da coluna A de índice 0
  const nomeSite = linha[0].toString().trim();
  const sufixoRaiz = linha[20];
  const nomeRaiz = `CONTRATOS ${sufixoRaiz}`;

  const pastaBase = DriveApp.getFolderById('1bj3f4XNs0xZAxpDXx3MeTSNpmmS6t666');
  const pastasRaiz = pastaBase.getFoldersByName(nomeRaiz);
  if (!pastasRaiz.hasNext()) throw new Error(`Pasta raiz "${nomeRaiz}" não encontrada.`);
  const pastaRaiz = pastasRaiz.next();

  const pastasSite = pastaRaiz.getFoldersByName(nomeSite);
  if (!pastasSite.hasNext()) throw new Error(`Pasta do site "${nomeSite}" não encontrada.`);
  const pastaSite = pastasSite.next();

//Achando a pasta ele cria uma pasta chamada 'CONTRATOS NOVOS' CASO ELA NÃO EXISTA
  let pastaContratos = pastaSite.getFoldersByName("CONTRATOS NOVOS");
  pastaContratos = pastaContratos.hasNext() ? pastaContratos.next() : pastaSite.createFolder("CONTRATOS NOVOS");

//VERIFICA A DATA DE HOJE
  const dataHoje = new Date();
  const dataFormatada = Utilities.formatDate(dataHoje, Session.getScriptTimeZone(), "dd.MM.yy");

//Conta quantas pastasexistem dentro da pasta 'CONTRATOS NOVOS'.
  let numeroNovaPasta = 1;
  const padraoNumerado = /^\d+\s-\s\d{2}\.\d{2}\.\d{2}$/;
  const pastasInternas = pastaContratos.getFolders();
  while (pastasInternas.hasNext()) {
    const nome = pastasInternas.next().getName();
    if (padraoNumerado.test(nome)) numeroNovaPasta++;
  }

//O nome da pasta fica com o seguinte formato '01 - 25.12.25' = 'Numero_de_pastas+1 - data'
  const nomeNovaPasta = `${String(numeroNovaPasta).padStart(2, '0')} - ${dataFormatada}`;
  const pastaFinal = pastaContratos.createFolder(nomeNovaPasta);

  // =====================
  // LOOP: copia documentos pra pastaFinal e preenche variáveis
  // =====================
  for (const docId of documentosSelecionados) {
    const nomeModelo = docModelos[docId];
    const novoNome = `${linha[0]} - ${nomeModelo}`;
    const novoDoc = DriveApp.getFileById(docId).makeCopy(novoNome, pastaFinal);

    const docNovo = DocumentApp.openById(novoDoc.getId());
    const corpo = docNovo.getBody();
//Para fazer a substituição no arquivo precisa ter o nome do cabeçalho entre chaves. Exemplo: {{Nome do cliente}}
    cabecalhos.forEach((chave, idx) => {
      const regex = new RegExp(`{{${chave}}}`, 'g');
      corpo.replaceText(regex.source, formatarData(linha[idx]));
    });

    docNovo.saveAndClose();
  }

  return true;
}
