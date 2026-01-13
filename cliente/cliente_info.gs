/**
 * ============================================================
 * ATUALIZA INFORMAÇÕES DINÂMICAS DA PLANILHA CLIENTE
 * ============================================================
 */
function cliente_atualizarInformacoes_(contexto) {

  if (!contexto) return;

  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('INFORMAÇÕES');
  if (!sheet) return;

  sheet.getRange('B3').setValue(contexto.nome);

  const pasta = DriveApp.getFolderById(contexto.pastaUnidadeId);
  sheet.getRange('B6').setValue(pasta.getUrl());

  const editores = pasta.getEditors().map(u => u.getEmail());
  const adminEmail = contexto.emailAdmin || '';

  sheet.getRange('A10:B30').clearContent();

  let linha = 10;
  sheet.getRange(`A${linha}`).setValue('Admin:');
  sheet.getRange(`B${linha}`).setValue(adminEmail).setFontWeight('bold');
  linha++;

  sheet.getRange(`A${linha}`).setValue('Usuários com acesso:');
  linha++;

  editores.forEach(email => {
    sheet.getRange(`B${linha}`).setValue(email);
    linha++;
  });
}
