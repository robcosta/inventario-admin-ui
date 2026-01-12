/**
 * ============================================================
 * UI ADMIN — PLANILHA CLIENTE
 * ============================================================
 */
function ui_admin_formatarPlanilhaCliente() {

  const ui = SpreadsheetApp.getUi();
  const contexto = admin_obterContextoAtivo_();

  if (!contexto || !contexto.planilhaClienteId) {
    ui.alert('Planilha cliente não encontrada no contexto ativo.');
    return;
  }

  const resp = ui.alert(
    'Formatar Planilha Cliente',
    'Isso irá aplicar o layout padrão da planilha cliente.\n\n' +
    'Deseja continuar?',
    ui.ButtonSet.YES_NO
  );

  if (resp !== ui.Button.YES) return;

  try {
    cliente_formatarPlanilhaInterface_(contexto.planilhaClienteId);
    ui.alert('Planilha cliente formatada com sucesso.');
  } catch (e) {
    Logger.log('[ADMIN][FORMATAR_CLIENTE][ERRO]');
    Logger.log(e);
    ui.alert('Erro ao formatar planilha cliente:\n\n' + e.message);
  }
}
