// ============================================================
// UI ADMIN — PLANILHA CLIENTE
// ============================================================

function ui_admin_formatarPlanilhaCliente() {

  const ui = SpreadsheetApp.getUi();
  const contexto = admin_obterContextoAtivo_();

  if (!contexto || !contexto.planilhaClienteId) {
    ui.alert('Planilha cliente não encontrada no contexto ativo.');
    return;
  }

  const resp = ui.alert(
    'Reformatar Planilha Cliente',
    'Isso irá REAPLICAR o layout padrão da planilha cliente,\n' +
    'independentemente do estado atual.\n\n' +
    'Deseja continuar?',
    ui.ButtonSet.YES_NO
  );

  if (resp !== ui.Button.YES) return;

  cliente_reformatarPlanilhaInterface_(contexto.planilhaClienteId);

  ui.alert('Planilha cliente reformatada com sucesso.');
}
