/**
 * ============================================================
 * GERENCIAR ACESSOS DO CONTEXTO (ADMIN)
 * ============================================================
 */

function ui_admin_gerenciarAcessosContexto() {

  const ui = SpreadsheetApp.getUi();

  const contexto = admin_obterContextoAtivo_();
  if (!contexto) {
    ui.alert('Nenhum contexto ativo nesta planilha.');
    return;
  }

  const pastaUnidadeId = contexto.pastaUnidadeId;
  if (!pastaUnidadeId) {
    ui.alert('Pasta da unidade não encontrada no contexto.');
    return;
  }

  const resp = ui.prompt(
    'Gerenciar Acessos do Contexto',
    'Contexto: ' + contexto.nome + '\n\n' +
    'Informe o e-mail do usuário que terá acesso TOTAL\n' +
    'à pasta deste contexto (Editor):',
    ui.ButtonSet.OK_CANCEL
  );

  if (resp.getSelectedButton() !== ui.Button.OK) return;

  const email = (resp.getResponseText() || '').trim();

  if (!email || !email.includes('@')) {
    ui.alert('E-mail inválido.');
    return;
  }

  try {
    const pasta = DriveApp.getFolderById(pastaUnidadeId);
    pasta.addEditor(email);

    const linkPasta = pasta.getUrl();

    // 📢 Mensagem para o admin repassar ao cliente
    const mensagemCliente =
      '✅ Acesso liberado ao Inventário Patrimonial\n\n' +
      'Você recebeu acesso à pasta do inventário referente a:\n' +
      contexto.nome + '\n\n' +
      '📁 Pasta de trabalho:\n' +
      linkPasta + '\n\n' +
      'Nessa pasta você poderá:\n' +
      '- Abrir a planilha do inventário\n' +
      '- Criar subpastas (UOPs)\n' +
      '- Enviar, renomear e apagar fotos\n\n' +
      'Em caso de dúvida, entre em contato com a administração.';

    ui.alert(
      'Acesso concedido com sucesso.\n\n' +
      'Usuário: ' + email + '\n\n' +
      'Mensagem para o cliente (copie e envie):\n\n' +
      mensagemCliente
    );

  } catch (e) {
    Logger.log('[ACESSOS][ERRO]');
    Logger.log(e);

    ui.alert(
      'Erro ao conceder acesso:\n\n' +
      e.message
    );
  }
}
