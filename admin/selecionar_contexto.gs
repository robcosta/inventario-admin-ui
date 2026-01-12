/**
 * ============================================================
 * SELECIONAR CONTEXTO DE TRABALHO (ADMIN)
 * ============================================================
 */

function ui_admin_selecionarContextoTrabalho() {

  const ui = SpreadsheetApp.getUi();

  const contextoAtual = admin_obterContextoAtivo_();
  const contextos = admin_listarContextos_();

  if (contextos.length === 0) {
    ui.alert('Nenhum contexto encontrado.');
    return;
  }

  let mensagem =
    'Contexto atual: ' +
    (contextoAtual ? contextoAtual.nome : 'NENHUM') +
    '\n\nSelecione o contexto que deseja abrir:\n\n';

  contextos.forEach((ctx, i) => {
    mensagem += `${i + 1} - ${ctx.nome}\n`;
  });

  const resp = ui.prompt(
    'Selecionar Contexto de Trabalho',
    mensagem,
    ui.ButtonSet.OK_CANCEL
  );

  if (resp.getSelectedButton() !== ui.Button.OK) return;

  const indice = Number(resp.getResponseText());

  if (!indice || indice < 1 || indice > contextos.length) {
    ui.alert('Seleção inválida.');
    return;
  }

  const escolhido = contextos[indice - 1];

  if (contextoAtual && escolhido.nome === contextoAtual.nome) {
    ui.alert('Este já é o contexto atual.');
    return;
  }

  // 🔓 Abrir a planilha do contexto escolhido
  SpreadsheetApp.openById(escolhido.planilhaOperacionalId);

  ui.alert(
    'O contexto "' + escolhido.nome + '" foi aberto.\n\n' +
    'Você pode fechar esta aba se desejar.'
  );
}
