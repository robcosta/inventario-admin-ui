// ============================================================
// cliente/cliente_onopen.gs
// ============================================================
function onOpen_(e) {
  try {
    const props = PropertiesService.getDocumentProperties();

    if (props.getProperty('CLIENTE_FORMATADO') !== 'true') {
      return;
    }

    const raw = props.getProperty('CONTEXTO_TRABALHO');
    if (!raw) return;

    const contexto = JSON.parse(raw);
    cliente_atualizarInformacoes_(contexto);

  } catch (err) {
    Logger.log('[CLIENTE][ONOPEN][ERRO]');
    Logger.log(err);
  }
}

