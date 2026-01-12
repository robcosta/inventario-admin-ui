/**
 * ============================================================
 * FORMATAÇÃO DA PLANILHA CLIENTE (BOOTSTRAP)
 * Executa UMA ÚNICA VEZ
 * ============================================================
 */
function cliente_formatarPlanilhaInterface_(spreadsheetId) {

  if (!spreadsheetId) {
    throw new Error('ID da planilha cliente não informado.');
  }

  const ss = SpreadsheetApp.openById(spreadsheetId);
  const props = PropertiesService.getDocumentProperties();

  if (props.getProperty('CLIENTE_FORMATADO') === 'true') {
    return;
  }

  let sheet = ss.getSheetByName('INFORMAÇÕES');
  if (!sheet) {
    sheet = ss.insertSheet('INFORMAÇÕES', 0);
  }

  sheet.clear();
  ss.setActiveSheet(sheet);

  sheet.setHiddenGridlines(true);
  sheet.setColumnWidths(1, 2, 520);
  sheet.setRowHeights(1, 30, 28);

  sheet.getRange('A1')
    .setValue('INVENTÁRIO PATRIMONIAL')
    .setFontSize(18)
    .setFontWeight('bold');

  sheet.getRange('A3').setValue('Contexto de Trabalho').setFontWeight('bold');
  sheet.getRange('A6').setValue('📁 Pasta de Trabalho (Drive)').setFontWeight('bold');
  sheet.getRange('A9').setValue('👤 Acessos').setFontWeight('bold');
  sheet.getRange('A13').setValue('▶️ Como utilizar').setFontWeight('bold');

  sheet.getRange('A14').setValue('• Utilize o menu da planilha para executar ações.');
  sheet.getRange('A15').setValue('• Envie fotos somente para a pasta indicada.');
  sheet.getRange('A16').setValue('• Não edite manualmente esta planilha.');

  sheet.setFrozenRows(18);

  const rangeProtegido = sheet.getRange('A1:B18');
  const protection = rangeProtegido.protect();
  protection.setDescription('Área informativa – não editar');

  props.setProperty('CLIENTE_FORMATADO', 'true');
}
