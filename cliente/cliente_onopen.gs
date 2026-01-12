/**
 * ============================================================
 * ONOPEN – PLANILHA CLIENTE
 * ============================================================
 */
function onOpen(e) {

  try {
    const props = PropertiesService.getDocumentProperties();

    // Só atua se for planilha cliente formatada
    if (props.getProperty('CLIENTE_FORMATADO') !== 'true') {
      return;
    }

    const contextoRaw = props.getProperty('CONTEXTO_TRABALHO');
    if (!contextoRaw) return;

    const contexto = JSON.parse(contextoRaw);

    cliente_atualizarInformacoes_(contexto);

  } catch (err) {
    Logger.log('[CLIENTE][ONOPEN][ERRO]');
    Logger.log(err);
  }
}
