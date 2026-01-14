function ui_admin_uploadCSV() {
  const html = HtmlService.createHtmlOutputFromFile('admin/admin_upload_csv')
    .setWidth(500)
    .setHeight(300);

  SpreadsheetApp.getUi().showModalDialog(
    html,
    'Enviar CSV para o Contexto'
  );
}

function admin_salvarCSVNoContexto(nomeArquivo, dataUrl) {

  const contexto = admin_obterContextoAtivo_();
  if (!contexto) {
    throw new Error('Nenhum contexto ativo.');
  }

  // 🔒 Validação rígida por nome
  if (!nomeArquivo || !nomeArquivo.toLowerCase().endsWith('.csv')) {
    throw new Error(
      'Arquivo inválido: "' + nomeArquivo + '". Apenas arquivos CSV são permitidos.'
    );
  }

  // 🔒 Validação mínima do dataURL
  if (!dataUrl || !dataUrl.startsWith('data:')) {
    throw new Error('Conteúdo do arquivo inválido.');
  }

  const pastaCSV = DriveApp.getFolderById(contexto.pastaCSVId);

  const base64 = dataUrl.split(',')[1];
  if (!base64) {
    throw new Error('Falha ao decodificar o arquivo CSV.');
  }

  const bytes = Utilities.base64Decode(base64);

  const blob = Utilities.newBlob(
    bytes,
    'text/csv',
    nomeArquivo
  );

  pastaCSV.createFile(blob);

  return 'CSV salvo com sucesso.';
}

