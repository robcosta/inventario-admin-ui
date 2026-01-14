function ui_admin_formatarPlanilhaCliente() {

  Logger.log('[ADMIN][FORMATAR_CLIENTE] Início');

  const ui = SpreadsheetApp.getUi();
  const contextoCliente = admin_obterContextoAtivo_();

  if (!contextoCliente || !contextoCliente.planilhaClienteId) {
    ui.alert('Planilha cliente não encontrada no contextoCliente ativo.');
    return;
  }

  const resp = ui.alert(
    'Formatar Planilha Cliente',
    'Deseja formatar e atualizar a planilha cliente?',
    ui.ButtonSet.YES_NO
  );

  if (resp !== ui.Button.YES) return;

  cliente_formatarPlanilhaInterface_(
    contextoCliente.planilhaClienteId,
    contextoCliente
  );

  cliente_montarInformacoes_(contextoCliente);
  ui.alert('Planilha cliente formatada e atualizada com sucesso.');

  Logger.log('[ADMIN][FORMATAR_CLIENTE] Finalizado');
}
