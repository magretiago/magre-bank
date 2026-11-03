/**
 * MAGRE BANK · API do Google Planilhas
 * ------------------------------------------------------------
 * Cole este arquivo em: Planilha → Extensões → Apps Script.
 * 1. Troque a CHAVE abaixo por uma senha só sua.
 * 2. Rode a função "configurar" uma vez (cria as abas).
 * 3. Implantar → Nova implantação → App da Web
 *    - Executar como: Eu
 *    - Quem pode acessar: Qualquer pessoa
 * 4. Copie a URL /exec e cole no site (botão de engrenagem).
 */

const CHAVE = 'troque-por-uma-senha-forte';

const ABAS = {
  lancamentos: { nome: 'Lancamentos', cab: ['id', 'data', 'descricao', 'valor', 'tipo', 'categoria', 'origem', 'origem_nome'] },
  cartoes:     { nome: 'Cartoes',     cab: ['id', 'banco', 'limite', 'vencimento', 'bandeira', 'final'] },
  alvos:       { nome: 'Alvos',       cab: ['id', 'nome', 'alvo', 'guardado', 'icone'] },
  config:      { nome: 'Config',      cab: ['chave', 'valor'] }
};

/* ---------------- Entrada da API ---------------- */

function doGet() {
  return resposta({ ok: true, app: 'MAGRE BANK API', info: 'Use POST com action "load" ou "save".' });
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); }
  catch (err) { return resposta({ ok: false, error: 'Requisição inválida' }); }

  if (body.token !== CHAVE) return resposta({ ok: false, error: 'Chave de acesso inválida' });

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (body.action === 'load') return resposta({ ok: true, data: lerTudo() });
    if (body.action === 'save') { gravarTudo(body.data); return resposta({ ok: true, updatedAt: new Date().toISOString() }); }
    return resposta({ ok: false, error: 'Ação desconhecida' });
  } catch (err) {
    return resposta({ ok: false, error: String(err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

/* ---------------- Configuração inicial ---------------- */

function configurar() {
  Object.keys(ABAS).forEach(k => aba(k));
  SpreadsheetApp.getActive().toast('Abas do MAGRE BANK prontas!', 'MAGRE BANK', 5);
}

function aba(chave) {
  const def = ABAS[chave];
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(def.nome);
  if (!sh) {
    sh = ss.insertSheet(def.nome);
    sh.getRange(1, 1, 1, def.cab.length).setValues([def.cab])
      .setFontWeight('bold').setBackground('#1A1A1A').setFontColor('#FF8C00');
    sh.setFrozenRows(1);
  }
  return sh;
}

/* ---------------- Leitura ---------------- */

function linhas(chave) {
  const sh = aba(chave);
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  const cab = ABAS[chave].cab;
  return sh.getRange(2, 1, n, cab.length).getValues()
    .filter(r => r[0] !== '')
    .map(r => Object.fromEntries(cab.map((c, i) => [c, r[i]])));
}

function dataTexto(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(v);
}

function lerTudo() {
  const cfg = Object.fromEntries(linhas('config').map(r => [r.chave, r.valor]));
  return {
    saldoInicial: Number(cfg.saldoInicial) || 0,
    cards: linhas('cartoes').map(r => ({
      id: String(r.id), banco: String(r.banco), limite: Number(r.limite), venc: Number(r.vencimento),
      bandeira: String(r.bandeira || 'Crédito'), final: String(r.final || '')
    })),
    goals: linhas('alvos').map(r => ({
      id: String(r.id), nome: String(r.nome), alvo: Number(r.alvo), guardado: Number(r.guardado), icone: String(r.icone || 'target')
    })),
    tx: linhas('lancamentos').map(r => ({
      id: String(r.id), data: dataTexto(r.data), desc: String(r.descricao), valor: Number(r.valor),
      tipo: String(r.tipo), cat: String(r.categoria), origem: String(r.origem)
    }))
  };
}

/* ---------------- Escrita ---------------- */

function substituir(chave, valores) {
  const sh = aba(chave);
  const cols = ABAS[chave].cab.length;
  const antigas = sh.getLastRow() - 1;
  if (antigas > 0) sh.getRange(2, 1, antigas, cols).clearContent();
  if (valores.length) sh.getRange(2, 1, valores.length, cols).setValues(valores);
}

function gravarTudo(d) {
  if (!d || !Array.isArray(d.tx) || !Array.isArray(d.cards) || !Array.isArray(d.goals)) {
    throw new Error('Dados incompletos');
  }
  const nomeCartao = Object.fromEntries(d.cards.map(c => [c.id, 'Cartão ' + c.banco]));
  const nomeAlvo = Object.fromEntries(d.goals.map(g => [g.id, g.nome]));
  const origemNome = t => t.origem === 'conta' ? 'Conta Corrente' : (nomeCartao[t.origem] || nomeAlvo[t.origem] || t.origem);

  substituir('cartoes', d.cards.map(c => [c.id, c.banco, c.limite, c.venc, c.bandeira || '', c.final || '']));
  substituir('alvos', d.goals.map(g => [g.id, g.nome, g.alvo, g.guardado, g.icone || '']));
  substituir('lancamentos', d.tx.map(t => [t.id, t.data, t.desc, t.valor, t.tipo, t.cat, t.origem, origemNome(t)]));
  substituir('config', [['saldoInicial', Number(d.saldoInicial) || 0], ['atualizadoEm', new Date()]]);
}

function resposta(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
