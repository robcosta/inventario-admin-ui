/**
 * ============================================================
 * BOOTSTRAP DO CONTEXTO DE TRABALHO (ADMIN)
 * ============================================================
 */

function ui_admin_criarContextoTrabalho() {

  const ui = SpreadsheetApp.getUi();

  // 🔒 Bloqueio imediato por estado da planilha
  if (admin_planilhaTemContexto_()) {
    ui.alert(
      'Esta planilha já pertence a um contexto.\n' +
      'Não é permitido criar outro contexto nela.'
    );
    return;
  }

  // 🔎 Listar contextos existentes (informativo)
  const contextosExistentes = admin_listarContextos_();

  let mensagemInfo = '';

  if (contextosExistentes.length > 0) {
    mensagemInfo += 'Contextos já existentes:\n\n';
    contextosExistentes.forEach(ctx => {
      mensagemInfo += '- ' + ctx.nome + '\n';
    });
    mensagemInfo +=
      '\nInforme o nome do NOVO contexto que deseja criar:';
  } else {
    mensagemInfo =
      'Nenhum contexto foi criado até o momento.\n\n' +
      'Informe o nome do primeiro contexto:';
  }

  // 1️⃣ Solicitar nome do contexto
  const resp = ui.prompt(
    'Criar Contexto de Trabalho',
    mensagemInfo,
    ui.ButtonSet.OK_CANCEL
  );

  if (resp.getSelectedButton() !== ui.Button.OK) return;

  const nomeUsuario = (resp.getResponseText() || '').trim();
  if (!nomeUsuario) {
    ui.alert('Nome inválido.');
    return;
  }

  const nomeContexto = nomeUsuario.toUpperCase();

  // 2️⃣ Bloqueio por nome global
  if (admin_contextoComNomeExiste_(nomeContexto)) {
    ui.alert(
      'O contexto "' + nomeContexto + '" já existe.\n\n' +
      'Utilize "Selecionar Contexto de Trabalho".'
    );
    return;
  }

  // 3️⃣ Criar estrutura de pastas
  const raiz = obterPastaInventario_();
  if (!raiz) {
    ui.alert('Pasta "Inventario Patrimonial" não encontrada.');
    return;
  }

  const planilhas = obterOuCriarSubpasta_(raiz, 'PLANILHAS');
  const contextos = obterOuCriarSubpasta_(planilhas, 'CONTEXTOS');
  const pastaContexto = obterOuCriarSubpasta_(contextos, nomeContexto);
  const pastaCSV = obterOuCriarSubpasta_(pastaContexto, 'CSV_ORIGEM');

  const unidades = obterOuCriarSubpasta_(raiz, 'UNIDADES');
  const pastaUnidade = obterOuCriarSubpasta_(unidades, nomeContexto);

  // 4️⃣ Criar Planilha Cliente
  const planilhaCliente = SpreadsheetApp.create('UI ' + nomeUsuario);
  DriveApp.getFileById(planilhaCliente.getId()).moveTo(pastaUnidade);

  // 5️⃣ Planilha Operacional (atual)
  const planilhaOperacional = SpreadsheetApp.getActiveSpreadsheet();
  planilhaOperacional.rename(nomeUsuario);
  DriveApp.getFileById(planilhaOperacional.getId()).moveTo(pastaContexto);

  // 6️⃣ Persistir contexto na planilha
  PropertiesService.getDocumentProperties().setProperty(
    'ADMIN_CONTEXTO_ATIVO',
    JSON.stringify({
      nome: nomeContexto,
      pastaContextoId: pastaContexto.getId(),
      pastaCSVId: pastaCSV.getId(),
      pastaUnidadeId: pastaUnidade.getId(),
      planilhaOperacionalId: planilhaOperacional.getId(),
      planilhaClienteId: planilhaCliente.getId(),
      criadoEm: new Date().toISOString()
    })
  );

  // 7️⃣ Atualizar menu
  criarMenuAdminOperacional_();

  ui.alert(
    'Contexto criado com sucesso.\n\n' +
    'A planilha será reaberta para garantir consistência.'
  );

  // 🔄 Refresh controlado
  SpreadsheetApp.openById(planilhaOperacional.getId());
}
