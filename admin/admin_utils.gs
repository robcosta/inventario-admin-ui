/**
 * ============================================================
 * UTILITÁRIOS ADMINISTRATIVOS
 * ============================================================
 */

function obterPastaInventario_() {
  const it = DriveApp.getFoldersByName('Inventario Patrimonial');
  return it.hasNext() ? it.next() : null;
}

function obterOuCriarSubpasta_(pai, nome) {
  const it = pai.getFoldersByName(nome);
  return it.hasNext() ? it.next() : pai.createFolder(nome);
}

function admin_planilhaTemContexto_() {
  return !!PropertiesService
    .getDocumentProperties()
    .getProperty('ADMIN_CONTEXTO_ATIVO');
}

function admin_contextoComNomeExiste_(nomeContexto) {
  const raiz = obterPastaInventario_();
  if (!raiz) return false;

  const planilhas = raiz.getFoldersByName('PLANILHAS');
  if (!planilhas.hasNext()) return false;

  const contextos = planilhas.next().getFoldersByName('CONTEXTOS');
  if (!contextos.hasNext()) return false;

  return contextos.next().getFoldersByName(nomeContexto).hasNext();
}

function admin_listarContextos_() {
  const raiz = obterPastaInventario_();
  if (!raiz) return [];

  const planilhas = raiz.getFoldersByName('PLANILHAS');
  if (!planilhas.hasNext()) return [];

  const contextos = planilhas.next().getFoldersByName('CONTEXTOS');
  if (!contextos.hasNext()) return [];

  const it = contextos.next().getFolders();
  const lista = [];

  while (it.hasNext()) {
    const pasta = it.next();
    lista.push({
      nome: pasta.getName(),
      pastaId: pasta.getId()
    });
  }

  return lista;
}
/**
 * Retorna o contexto ativo da planilha (ou null)
 */
function admin_obterContextoAtivo_() {
  const raw = PropertiesService
    .getDocumentProperties()
    .getProperty('ADMIN_CONTEXTO_ATIVO');

  return raw ? JSON.parse(raw) : null;
}