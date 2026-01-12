/**
 * ============================================================
 * MENU ADMIN — PONTO DE ENTRADA SEGURO
 * ============================================================
 * REGRA:
 * - onOpen deve ser mínimo
 * - TODA lógica fica fora
 * - try/catch evita menu "sumir"
 */

function onOpen() {
  try {
    menuAdmin_onOpen_();
  } catch (e) {
    Logger.log('[MENU_ADMIN][ERRO]');
    Logger.log(e);

    SpreadsheetApp.getUi().alert(
      'Erro ao inicializar o menu de Administração.\n\n' +
      'Detalhes:\n' + e.message
    );
  }
}

/**
 * ============================================================
 * LÓGICA REAL DO MENU
 * ============================================================
 */
function menuAdmin_onOpen_() {

  // 🔎 Verifica estado da planilha atual
  const temContexto = admin_planilhaTemContexto_();

  if (temContexto) {
    criarMenuAdminOperacional_();
  } else {
    criarMenuAdminBootstrap_();
  }
}

/**
 * ============================================================
 * MENU — PLANILHA SEM CONTEXTO (BOOTSTRAP)
 * ============================================================
 */
function criarMenuAdminBootstrap_() {

  const ui = SpreadsheetApp.getUi();

  ui.createMenu('🏛️ Inventário – Administração')
    .addItem('➕ Criar Contexto de Trabalho', 'ui_admin_criarContextoTrabalho')
    .addToUi();
}

/**
 * ============================================================
 * MENU — PLANILHA COM CONTEXTO (OPERACIONAL)
 * ============================================================
 */
function criarMenuAdminOperacional_() {

  const ui = SpreadsheetApp.getUi();

  ui.createMenu('🏛️ Inventário – Administração')
    .addItem('🔁 Selecionar Contexto de Trabalho', 'ui_admin_selecionarContextoTrabalho')
    .addSeparator()

    .addItem('⚙️ Configurar Planilha Base Patrimonial', 'ui_admin_configurarPlanilhaBase')
    .addSeparator()

    .addItem('📥 Importar CSVs do Contexto', 'ui_admin_importarCSVs')
    .addItem('📊 Popular Planilha Operacional', 'ui_admin_popularPlanilhaOperacional')
    .addItem('🎨 Formatar Planilha Operacional', 'ui_admin_formatarPlanilhaOperacional')
    .addSeparator()

    .addItem('🗂️ Pastas de Trabalho', 'ui_admin_pastasTrabalho')
    .addItem('🧪 Diagnóstico', 'ui_admin_diagnostico')
    .addToUi();
}

