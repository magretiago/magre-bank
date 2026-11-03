/**
 * MAGRE BANK · API do Google Planilhas + lembretes de contas por e-mail
 * Versão 3 · contas bancárias, transferências, categorias, faturas de cartão e parcelas
 * ------------------------------------------------------------
 * Cole este arquivo em: Planilha → Extensões → Apps Script.
 * 1. Troque a CHAVE abaixo por uma senha só sua.
 * 2. Rode a função "configurar" uma vez (cria as abas e o lembrete diário das 8h).
 * 3. Implantar → Nova implantação → App da Web
 *    - Executar como: Eu
 *    - Quem pode acessar: Qualquer pessoa
 *    (ao atualizar o código: Implantar → Gerenciar implantações → Editar → Nova versão)
 * 4. Copie a URL /exec e cole no site (botão de engrenagem).
 *
 * Opcional, na aba Config:
 *   email     → para onde enviar os lembretes (padrão: o seu e-mail Google; aceita vários separados por vírgula)
 *   diasAviso → quantos dias antes do vencimento avisar (padrão: 3)
 */

const CHAVE = 'troque-por-uma-senha-forte';
const HORA_LEMBRETE = 8;
const VERSAO = 3;

// Colunas novas sempre entram no FINAL, para os dados antigos continuarem alinhados.
const ABAS = {
  lancamentos: { nome: 'Lancamentos',     cab: ['id', 'data', 'descricao', 'valor', 'tipo', 'categoria', 'origem', 'origem_nome', 'conta_id', 'referencia', 'forma', 'destino', 'conta_aporte', 'cartao', 'parcela', 'parcelas', 'grupo', 'fatura'] },
  bancos:      { nome: 'ContasBancarias', cab: ['id', 'nome', 'tipo', 'saldo_inicial', 'saldo_atual'] },
  cartoes:     { nome: 'Cartoes',         cab: ['id', 'banco', 'limite', 'vencimento', 'bandeira', 'final', 'fechamento', 'fechamento_estimado'] },
  faturas:     { nome: 'Faturas',         cab: ['cartao_id', 'cartao', 'fatura', 'periodo', 'fechamento', 'vencimento', 'total', 'pago', 'em_aberto', 'situacao'] },
  alvos:       { nome: 'Alvos',           cab: ['id', 'nome', 'alvo', 'guardado', 'icone'] },
  contas:      { nome: 'Contas',          cab: ['id', 'nome', 'dia_vencimento', 'valor', 'categoria', 'pago_ate', 'proximo_vencimento', 'situacao'] },
  categorias:  { nome: 'Categorias',      cab: ['tipo', 'nome'] },
  config:      { nome: 'Config',          cab: ['chave', 'valor'] }
};

/* ================= Entrada da API ================= */

function doGet() {
  return resposta({ ok: true, app: 'MAGRE BANK API', versao: VERSAO, info: 'Use POST com action "load" ou "save".' });
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); }
  catch (err) { return resposta({ ok: false, error: 'Requisição inválida' }); }

  if (body.token !== CHAVE) return resposta({ ok: false, error: 'Chave de acesso inválida' });

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (body.action === 'load') return resposta({ ok: true, versao: VERSAO, data: lerTudo() });
    if (body.action === 'save') { gravarTudo(body.data); return resposta({ ok: true, versao: VERSAO, updatedAt: new Date().toISOString() }); }
    return resposta({ ok: false, error: 'Ação desconhecida' });
  } catch (err) {
    return resposta({ ok: false, error: String(err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

/* ================= Configuração inicial ================= */

function configurar() {
  Object.keys(ABAS).forEach(k => aba(k));
  const cfg = lerConfig();
  const padrao = {};
  if (cfg.email === undefined) padrao.email = Session.getEffectiveUser().getEmail();
  if (cfg.diasAviso === undefined) padrao.diasAviso = 3;
  gravarConfig(padrao);

  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'enviarLembretes')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('enviarLembretes').timeBased().everyDays(1).atHour(HORA_LEMBRETE).create();

  SpreadsheetApp.getActive().toast(`Abas prontas e lembrete diário às ${HORA_LEMBRETE}h ativado!`, 'MAGRE BANK', 6);
}

function aba(chave) {
  const def = ABAS[chave];
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(def.nome);
  const criarCabecalho = s => s.getRange(1, 1, 1, def.cab.length).setValues([def.cab])
    .setFontWeight('bold').setBackground('#1A1A1A').setFontColor('#FF8C00');
  if (!sh) {
    sh = ss.insertSheet(def.nome);
    criarCabecalho(sh);
    sh.setFrozenRows(1);
    if (chave === 'contas') sh.getRange('F:F').setNumberFormat('@');
    if (chave === 'faturas') sh.getRange('C:C').setNumberFormat('@');
  } else {
    // aba de uma versão anterior: acrescenta as colunas novas ao cabeçalho
    const atual = sh.getRange(1, 1, 1, def.cab.length).getValues()[0];
    if (atual.join('|') !== def.cab.join('|')) criarCabecalho(sh);
  }
  return sh;
}

/* ================= Leitura ================= */

function linhas(chave) {
  const sh = aba(chave);
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  const cab = ABAS[chave].cab;
  return sh.getRange(2, 1, n, cab.length).getValues()
    .filter(r => r[0] !== '')
    .map(r => Object.fromEntries(cab.map((c, i) => [c, r[i]])));
}

const tz = () => Session.getScriptTimeZone();
function dataTexto(v, fmt) {
  if (v instanceof Date) return Utilities.formatDate(v, tz(), fmt || 'yyyy-MM-dd');
  return v === '' || v === null ? '' : String(v);
}

function lerConfig() {
  return Object.fromEntries(linhas('config').map(r => [String(r.chave), r.valor]));
}

function lerTudo() {
  const cfg = lerConfig();
  const categories = {};
  linhas('categorias').forEach(r => {
    const k = String(r.tipo);
    (categories[k] = categories[k] || []).push(String(r.nome));
  });
  return {
    saldoInicial: Number(cfg.saldoInicial) || 0,        // usado só na migração da versão 1
    diasAviso: Number(cfg.diasAviso) || 3,
    accounts: linhas('bancos').map(r => ({
      id: String(r.id), nome: String(r.nome), tipo: String(r.tipo || 'corrente'), saldoInicial: Number(r.saldo_inicial) || 0
    })),
    categories,
    cards: linhas('cartoes').map(r => ({
      id: String(r.id), banco: String(r.banco), limite: Number(r.limite), venc: Number(r.vencimento),
      bandeira: String(r.bandeira || 'Crédito'), final: String(r.final || ''),
      fech: r.fechamento === '' ? undefined : Number(r.fechamento),
      fechEstimado: r.fechamento_estimado === true || String(r.fechamento_estimado).toLowerCase() === 'sim'
    })),
    goals: linhas('alvos').map(r => ({
      id: String(r.id), nome: String(r.nome), alvo: Number(r.alvo), guardado: Number(r.guardado), icone: String(r.icone || 'target')
    })),
    bills: linhas('contas').map(r => ({
      id: String(r.id), nome: String(r.nome), venc: Number(r.dia_vencimento),
      valor: r.valor === '' ? null : Number(r.valor), cat: String(r.categoria || 'Outros'),
      pagoAte: dataTexto(r.pago_ate, 'yyyy-MM')
    })),
    tx: linhas('lancamentos').map(r => {
      const t = {
        id: String(r.id), data: dataTexto(r.data), desc: String(r.descricao), valor: Number(r.valor),
        tipo: String(r.tipo), cat: String(r.categoria), origem: String(r.origem)
      };
      if (r.conta_id) { t.conta = String(r.conta_id); t.ref = dataTexto(r.referencia, 'yyyy-MM'); }
      if (r.forma) t.forma = String(r.forma);
      if (r.destino) t.destino = String(r.destino);
      if (r.conta_aporte) t.acc = String(r.conta_aporte);
      if (r.cartao) t.cartao = String(r.cartao);
      if (t.tipo === 'fatura') t.ref = dataTexto(r.referencia, 'yyyy-MM');
      if (Number(r.parcelas) > 1) { t.parc = Number(r.parcela); t.parcN = Number(r.parcelas); t.grupo = String(r.grupo); }
      return t;
    })
  };
}

/* ================= Escrita ================= */

function substituir(chave, valores) {
  const sh = aba(chave);
  const cols = ABAS[chave].cab.length;
  const antigas = sh.getLastRow() - 1;
  if (antigas > 0) sh.getRange(2, 1, antigas, cols).clearContent();
  if (valores.length) sh.getRange(2, 1, valores.length, cols).setValues(valores);
}

function gravarConfig(novos) {
  const cfg = Object.assign(lerConfig(), novos);
  substituir('config', Object.keys(cfg).map(k => [k, cfg[k]]));
}

function saldoConta(id, inicial, tx) {
  let v = Number(inicial) || 0;
  tx.forEach(t => {
    if (t.tipo === 'receita' && t.origem === id) v += t.valor;
    else if (t.tipo === 'despesa' && t.origem === id) v -= t.valor;
    else if (t.tipo === 'aporte' && t.acc === id) v -= t.valor;
    else if (t.tipo === 'fatura' && t.origem === id) v -= t.valor;
    else if (t.tipo === 'transferencia') { if (t.origem === id) v -= t.valor; if (t.destino === id) v += t.valor; }
  });
  return Math.round(v * 100) / 100;
}

function gravarTudo(d) {
  if (!d || !Array.isArray(d.tx) || !Array.isArray(d.cards) || !Array.isArray(d.goals)) {
    throw new Error('Dados incompletos');
  }
  const bills = Array.isArray(d.bills) ? d.bills : [];
  // Um site antigo (sem contas bancárias) não apaga as contas nem as categorias já salvas
  const accounts = Array.isArray(d.accounts) ? d.accounts : lerTudo().accounts;

  const nomeConta = Object.fromEntries(accounts.map(a => [a.id, a.nome]));
  const nomeCartao = Object.fromEntries(d.cards.map(c => [c.id, 'Cartão ' + c.banco]));
  const nomeAlvo = Object.fromEntries(d.goals.map(g => [g.id, g.nome]));
  const origemNome = t => {
    if (t.tipo === 'transferencia') return `${nomeConta[t.origem] || t.origem} → ${nomeConta[t.destino] || t.destino}`;
    if (t.tipo === 'cartao') return (nomeCartao[t.origem] || t.origem) + (t.fatura ? ' · fatura ' + mesCurto(t.fatura) : '') + (t.parcN > 1 ? ` · ${t.parc}/${t.parcN}` : '');
    if (t.tipo === 'fatura') return `${nomeConta[t.origem] || t.origem}${t.forma ? ' · ' + t.forma : ''} → Fatura ${(nomeCartao[t.cartao] || '').replace('Cartão ', '')} ${mesCurto(t.ref || '')}`;
    if (t.tipo === 'aporte') return (nomeAlvo[t.origem] || t.origem) + (t.acc ? ' · de ' + (nomeConta[t.acc] || t.acc) : '');
    return (nomeConta[t.origem] || (t.origem === 'conta' ? 'Conta Corrente' : t.origem)) + (t.forma ? ' · ' + t.forma : '');
  };
  const hoje = inicioDoDia(new Date());
  const dias = Number(lerConfig().diasAviso) || 3;

  if (Array.isArray(d.accounts)) {
    substituir('bancos', accounts.map(a => [a.id, a.nome, a.tipo || 'corrente', Number(a.saldoInicial) || 0, saldoConta(a.id, a.saldoInicial, d.tx)]));
  }
  if (d.categories && typeof d.categories === 'object') {
    const rows = [];
    Object.keys(d.categories).forEach(k => (d.categories[k] || []).forEach(n => rows.push([k, n])));
    substituir('categorias', rows);
  }
  substituir('cartoes', d.cards.map(c => [c.id, c.banco, c.limite, c.venc, c.bandeira || '', c.final || '', c.fech || '', c.fechEstimado ? 'sim' : '']));
  if (Array.isArray(d.faturas)) {
    substituir('faturas', d.faturas.map(f => {
      const aberto = Math.max(0, Math.round((f.total - f.pago) * 100) / 100);
      return [f.cartaoId, f.banco, f.ciclo, `${dataBR(f.inicio)} a ${dataBR(f.fechamento)}`, f.fechamento, f.vencimento, f.total, f.pago, aberto, situacaoFatura(f, aberto, hoje)];
    }));
  }
  substituir('alvos', d.goals.map(g => [g.id, g.nome, g.alvo, g.guardado, g.icone || '']));
  substituir('contas', bills.map(b => {
    const s = situacaoConta(b, hoje, dias);
    return [b.id, b.nome, b.venc, b.valor == null ? '' : b.valor, b.cat, b.pagoAte || '', s.vencimento, s.rotulo];
  }));
  substituir('lancamentos', d.tx.map(t => [
    t.id, t.data, t.desc, t.valor, t.tipo, t.cat, t.origem, origemNome(t),
    t.conta || '', t.ref || '', t.forma || '', t.destino || '', t.acc || '',
    t.cartao || '', t.parc || '', t.parcN || '', t.grupo || '', t.fatura || ''
  ]));
  gravarConfig({ atualizadoEm: new Date() });
}

/* ================= Faturas ================= */

const MESES = ['jan', 'fev', 'mar', 'abr', 'mai', 'jun', 'jul', 'ago', 'set', 'out', 'nov', 'dez'];
function mesCurto(c) { const p = String(c).split('-'); return p.length === 2 ? `${MESES[Number(p[1]) - 1]}/${p[0].slice(2)}` : String(c); }
function paraData(v) {
  if (v instanceof Date) return inicioDoDia(v);
  const [y, m, d] = String(v).split('-').map(Number);
  return new Date(y, m - 1, d);
}
function dataBR(v) { return Utilities.formatDate(paraData(v), tz(), 'dd/MM'); }

function situacaoFatura(f, aberto, hoje) {
  if (f.total > 0 && aberto <= 0) return 'Paga';
  const fecha = paraData(f.fechamento), vence = paraData(f.vencimento);
  if (fecha >= hoje) return 'Aberta';
  const diff = Math.round((vence - hoje) / 86400000);
  return diff < 0 ? `ATRASADA há ${-diff} dia(s)` : `Fechada · vence em ${diff} dia(s)`;
}

/* ================= Situação das contas ================= */

function inicioDoDia(d) { const x = new Date(d); x.setHours(0, 0, 0, 0); return x; }
function mesTexto(d) { return Utilities.formatDate(d, tz(), 'yyyy-MM'); }
function somarMeses(s, n) { const [y, m] = s.split('-').map(Number); return mesTexto(new Date(y, m - 1 + n, 1)); }
function dataVenc(ciclo, dia) {
  const [y, m] = ciclo.split('-').map(Number);
  return new Date(y, m - 1, Math.min(dia, new Date(y, m, 0).getDate()));
}

function situacaoConta(b, hoje, dias) {
  const atual = mesTexto(hoje);
  const ciclo = b.pagoAte ? somarMeses(b.pagoAte, 1) : atual;
  const venc = dataVenc(ciclo, Number(b.venc));
  const diff = Math.round((venc - hoje) / 86400000);
  let status = 'pendente';
  if (diff < 0) status = 'vencida';
  else if (diff <= dias) status = 'breve';
  else if (ciclo > atual) status = 'paga';
  const rotulo = {
    vencida: `VENCIDA há ${-diff} dia(s)`,
    breve: diff === 0 ? 'Vence HOJE' : `Vence em ${diff} dia(s)`,
    paga: 'Paga este mês',
    pendente: `Vence em ${diff} dias`
  }[status];
  return { status, diff, ciclo, rotulo, vencimento: Utilities.formatDate(venc, tz(), 'dd/MM/yyyy') };
}

/* ================= Lembretes por e-mail ================= */

function enviarLembretes() { enviar_(false); }
function testarLembrete() { enviar_(true); }

function enviar_(forcar) {
  const cfg = lerConfig();
  const dias = Number(cfg.diasAviso) || 3;
  const para = String(cfg.email || Session.getEffectiveUser().getEmail());
  const hoje = inicioDoDia(new Date());
  const dados = lerTudo();

  const itens = dados.bills
    .map(b => Object.assign({ b }, situacaoConta(b, hoje, dias)))
    .filter(i => i.status === 'vencida' || i.status === 'breve');

  // faturas de cartão em aberto que vencem dentro do aviso (ou já venceram)
  linhas('faturas').forEach(f => {
    const aberto = Number(f.em_aberto) || 0;
    if (aberto <= 0) return;
    const vence = paraData(f.vencimento);
    const diff = Math.round((vence - hoje) / 86400000);
    if (diff > dias) return;
    const status = diff < 0 ? 'vencida' : 'breve';
    itens.push({
      b: { nome: `Fatura ${f.cartao} · ${mesCurto(dataTexto(f.fatura, 'yyyy-MM'))}`, valor: aberto },
      status, diff, vencimento: Utilities.formatDate(vence, tz(), 'dd/MM/yyyy'),
      rotulo: status === 'vencida' ? `VENCIDA há ${-diff} dia(s)` : diff === 0 ? 'Vence HOJE' : `Vence em ${diff} dia(s)`
    });
  });
  itens.sort((a, z) => a.diff - z.diff);

  if (!itens.length && !forcar) return;

  const vencidas = itens.filter(i => i.status === 'vencida').length;
  const assunto = itens.length
    ? `MAGRE BANK · ${vencidas ? vencidas + ' conta(s) vencida(s)' : itens.length + ' conta(s) vencendo em breve'}`
    : 'MAGRE BANK · Teste de lembrete: nenhuma conta pendente';

  const linhasHtml = itens.map(i => {
    const cor = i.status === 'vencida' ? '#FF5C5C' : '#FF8C00';
    const valor = i.b.valor ? Number(i.b.valor).toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' }) : 'Valor variável';
    return `<tr>
      <td style="padding:12px 14px;border-bottom:1px solid #2E2E2E;color:#F5F5F5;font-weight:600">${escapar(i.b.nome)}</td>
      <td style="padding:12px 14px;border-bottom:1px solid #2E2E2E;color:#9A9A9A">${i.vencimento}</td>
      <td style="padding:12px 14px;border-bottom:1px solid #2E2E2E;color:#F5F5F5">${valor}</td>
      <td style="padding:12px 14px;border-bottom:1px solid #2E2E2E;color:${cor};font-weight:700">${i.rotulo}</td>
    </tr>`;
  }).join('');

  const html = `
  <div style="background:#1A1A1A;padding:28px 16px;font-family:Montserrat,Arial,sans-serif">
    <div style="max-width:560px;margin:0 auto;background:#242424;border:1px solid #2E2E2E;border-radius:18px;overflow:hidden">
      <div style="padding:22px 24px;background:linear-gradient(135deg,#FF6A00,#FF8C00)">
        <div style="font-size:18px;font-weight:800;color:#111;letter-spacing:.5px">MAGRE BANK</div>
        <div style="font-size:12px;color:#2a1400">Seu banco. Mais possibilidades.</div>
      </div>
      <div style="padding:22px 24px">
        <p style="margin:0 0 6px;color:#F5F5F5;font-size:16px;font-weight:700">${itens.length ? 'Você tem contas para pagar' : 'Tudo em dia por aqui'}</p>
        <p style="margin:0 0 18px;color:#9A9A9A;font-size:13px">${itens.length ? 'Depois de pagar, marque como liquidada no site para lançar no extrato.' : 'Este é um e-mail de teste. Os lembretes chegam todo dia às ' + HORA_LEMBRETE + 'h quando houver contas a vencer.'}</p>
        ${itens.length ? `<table style="width:100%;border-collapse:collapse;font-size:13px">
          <tr><th align="left" style="padding:8px 14px;color:#9A9A9A;font-size:11px;text-transform:uppercase">Conta</th><th align="left" style="padding:8px 14px;color:#9A9A9A;font-size:11px;text-transform:uppercase">Vencimento</th><th align="left" style="padding:8px 14px;color:#9A9A9A;font-size:11px;text-transform:uppercase">Valor</th><th align="left" style="padding:8px 14px;color:#9A9A9A;font-size:11px;text-transform:uppercase">Situação</th></tr>
          ${linhasHtml}</table>` : ''}
      </div>
    </div>
  </div>`;

  MailApp.sendEmail({ to: para, subject: assunto, htmlBody: html, name: 'MAGRE BANK' });
}

function escapar(s) { return String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c])); }

function resposta(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
