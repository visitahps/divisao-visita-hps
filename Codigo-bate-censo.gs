// Bate censo: lista compartilhada dos hospitais (cole tudo no Apps Script da planilha)
// Cada código dá acesso só às abas do seu hospital.
const CODIGOS = {
  'hps-87m4tdpe': '',        // Hospital Paulo Sacramento: abas "dias" e "pacientes"
  'hjan-k7f4s3ud': 'hjan',   // Hospital Jardim Anália: abas "dias hjan" e "pacientes hjan"
  'hcru-5gmy9j61': 'hcru',   // Hospital Cruzeiro do Sul
  'hbos-9w7mot6r': 'hbos',   // Hospital Bosque da Saúde
  'hros-94teoe1t': 'hros'    // Hospital N. Sra. do Rosário
};
let SUF = '';

// Painel de gestão: '*' vê os 5 hospitais; os coordenadores veem só as unidades listadas. Não passe para as equipes.
const GESTORES = {
  'ger-m7q4x2kd': '*',                // Alexandre: todos os hospitais
  'ger-onzv59uu': ['hps'],            // Coordenação HPS (Fabiano e Ariovaldo)
  'ger-5v95gwby': ['hjan', 'hros'],   // Coordenação Jardim Anália + N. Sra. do Rosário (Felipe Fuchs)
  'ger-03z5zw7j': ['hbos'],           // Coordenação Bosque da Saúde (Fernando Villa Neves)
  'ger-k37ewovs': ['hcru']            // Coordenação Cruzeiro do Sul (Emanoel Nascimento)
};
// aba do bate censo de cada hospital (mesmo sufixo dos códigos acima)
const HOSP_SUF = { hps: '', hjan: 'hjan', hcru: 'hcru', hbos: 'hbos', hros: 'hros' };
const CENSOS_GUARDADOS = 400; // dias de censo guardados para o painel
const DIAS_GUARDADOS = 45;    // apaga da planilha os dias mais antigos que isso

const COLS = ['chave', 'dia', 'id', 'posto', 'leito', 'paciente', 'convenio', 'rp_vs', 'dias_internado', 'medico', 'previsao', 'o_que_falta', 'pendencias_do_dia', 'pendencia_resolvida', 'desfecho', 'atualizado', 'ts', 'json'];
const PREV = { efetivada: 'Alta efetivada', tarde: 'Hoje à tarde', '24h': 'Em 24h', '48h': 'Em 48h', sem: 'Sem previsão' };
const DESF = { alta: 'Alta', obito: 'Óbito', uti: 'Transferência UTI', outra: 'Transferência outra unidade' };

function doGet(e) { return responder((e && e.parameter) || {}); }
function doPost(e) { let p = {}; try { p = JSON.parse(e.postData.contents); } catch (x) {} return responder(p); }
function saida(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function responder(p) {
  const gst = GESTORES[String(p.code || '')];
  if (gst) return gestor(p, gst === '*' || gst === true ? Object.keys(HOSP_SUF) : gst);
  const suf = CODIGOS[String(p.code || '')];
  if (suf === undefined) return saida({ ok: false, erro: 'codigo' });
  SUF = suf ? ' ' + suf : '';
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (p.op === 'ping') { abaDias(ss); abaPac(ss); return saida({ ok: true }); }
    if (p.op === 'days') return saida({ ok: true, days: lerDias(ss) });
    if (p.op === 'day') { const agora = Date.now(); return saida({ ok: true, now: agora, pats: lerDia(ss, String(p.day), Number(p.since || 0)) }); }
    const lock = LockService.getScriptLock(); lock.waitLock(25000);
    try {
      if (p.op === 'putDay') { gravarDia(ss, String(p.day), p.meta || {}); return saida({ ok: true }); }
      if (p.op === 'put') { gravarPacientes(ss, String(p.day), p.items || []); return saida({ ok: true, now: Date.now() }); }
    } finally { lock.releaseLock(); }
    return saida({ ok: false, erro: 'op' });
  } catch (err) { return saida({ ok: false, erro: String(err) }); }
}

function aba(ss, nome, cab) {
  let sh = ss.getSheetByName(nome);
  if (!sh) { sh = ss.insertSheet(nome); sh.getRange(1, 1, 1, cab.length).setValues([cab]).setFontWeight('bold'); sh.setFrozenRows(1); sh.getRange('A:E').setNumberFormat('@'); }
  return sh;
}
function abaDias(ss) { return aba(ss, 'dias' + SUF, ['dia', 'data', 'total', 'atualizado', 'json']); }
function abaPac(ss) { return aba(ss, 'pacientes' + SUF, COLS); }

function lerDias(ss) {
  return abaDias(ss).getDataRange().getValues().slice(1).filter(r => r[0]).map(r => {
    let m = {}; try { m = JSON.parse(r[4]); } catch (x) {}
    m.id = String(r[0]); return m;
  });
}
function lerDia(ss, dia, desde) {
  const out = {};
  abaPac(ss).getDataRange().getValues().slice(1).forEach(r => {
    if (String(r[1]) !== dia || Number(r[16]) <= desde) return;
    try { out[String(r[0]).split('|')[1]] = JSON.parse(r[17]); } catch (x) {}
  });
  return out;
}
function gravarDia(ss, dia, meta) {
  const sh = abaDias(ss), v = sh.getDataRange().getValues();
  const linha = [dia, meta.date || '', meta.total || '', new Date(), JSON.stringify(meta)];
  const i = v.findIndex((r, k) => k > 0 && String(r[0]) === dia);
  if (i > 0) sh.getRange(i + 1, 1, 1, linha.length).setValues([linha]);
  else { sh.appendRow(linha); podar(ss); }
}
function juntar(a, b) {
  const o = Object.assign({}, a);
  Object.keys(b).forEach(k => { const v = b[k]; o[k] = v && typeof v === 'object' && !Array.isArray(v) && a[k] && typeof a[k] === 'object' ? juntar(a[k], v) : v; });
  return o;
}
function gravarPacientes(ss, dia, itens) {
  const sh = abaPac(ss), v = sh.getDataRange().getValues(), idx = {}, nk = {};
  v.forEach((r, k) => { if (k > 0) idx[String(r[0])] = k; });
  const novos = [], ts = Date.now();
  itens.forEach(it => {
    const chave = dia + '|' + it.id, k = idx[chave];
    let antigo = null;
    try { if (k > 0) antigo = JSON.parse(v[k][17]); else if (nk[chave] != null) antigo = JSON.parse(novos[nk[chave]][17]); } catch (x) {}
    const doc = !antigo || it.isNew ? it.body : juntar(antigo, it.body || {});
    const c = doc.c || {};
    const linha = [chave, dia, String(it.id), c.unit || '', c.bed || '', c.name || '', c.conv || '', c.team || '', c.dias == null ? '' : c.dias, c.visitador || '',
      PREV[doc.prev] || '', doc.falta || '', doc.pend || '', doc.resolvido ? 'Sim' : '', DESF[doc.desfecho] || (doc.prev === 'efetivada' ? 'Alta' : ''), new Date(), ts, JSON.stringify(doc)];
    if (k > 0) { sh.getRange(k + 1, 1, 1, COLS.length).setValues([linha]); v[k] = linha; }
    else if (nk[chave] != null) novos[nk[chave]] = linha;
    else { nk[chave] = novos.length; novos.push(linha); }
  });
  if (novos.length) sh.getRange(sh.getLastRow() + 1, 1, novos.length, COLS.length).setValues(novos);
}
function podar(ss) {
  const limite = Utilities.formatDate(new Date(Date.now() - DIAS_GUARDADOS * 864e5), 'America/Sao_Paulo', 'yyyyMMdd');
  [[abaPac(ss), 1], [abaDias(ss), 0]].forEach(([sh, col]) => {
    const v = sh.getDataRange().getValues(); if (v.length < 2) return;
    const manter = v.filter((r, k) => k === 0 || String(r[col]) >= limite);
    if (manter.length === v.length) return;
    sh.getRange(2, 1, v.length - 1, v[0].length).clearContent();
    if (manter.length > 1) sh.getRange(2, 1, manter.length - 1, v[0].length).setValues(manter.slice(1));
  });
}
// Rode esta função uma vez no editor para autorizar e criar as abas.
function preparar() { const ss = SpreadsheetApp.getActiveSpreadsheet(); Object.values(CODIGOS).forEach(x => { SUF = x ? ' ' + x : ''; abaDias(ss); abaPac(ss); }); SUF = ''; abaCensos(ss); abaAcoes(ss); }


/* ===================== painel de gestão ===================== */
function gestor(p, OK) {
  const pode = h => OK.indexOf(String(h)) >= 0;
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (p.op === 'ping') { abaCensos(ss); abaAcoes(ss); return saida({ ok: true, gestor: true, hosps: OK }); }
    if (p.op === 'censos') {
      const lock = LockService.getScriptLock(); lock.waitLock(25000);
      try { sincronizarCensos(ss, String(p.from || '')); } finally { lock.releaseLock(); }
      return saida({ ok: true, now: Date.now(), censos: lerCensos(ss, String(p.from || '')).filter(c => pode(c.hosp)) });
    }
    if (p.op === 'acoes') { const a = lerAcoes(ss), o = {}; Object.keys(a).forEach(k => { if (pode(k.split('|')[0])) o[k] = a[k]; }); return saida({ ok: true, acoes: o }); }
    if (p.op === 'codigos') { const c = {}; Object.keys(CODIGOS).forEach(k => { const h = CODIGOS[k] || 'hps'; if (pode(h)) c[h] = k; }); return saida({ ok: true, codigos: c }); }
    if ((p.op === 'bc' || p.op === 'putCenso' || p.op === 'putAcao') && !pode(p.hosp)) return saida({ ok: false, erro: 'hosp' });
    if (p.op === 'bc') {
      const s = HOSP_SUF[String(p.hosp)]; if (s === undefined) return saida({ ok: false, erro: 'hosp' });
      return saida({ ok: true, bc: lerBC(ss, s ? 'pacientes ' + s : 'pacientes', String(p.from || '')) });
    }
    const lock = LockService.getScriptLock(); lock.waitLock(25000);
    try {
      if (p.op === 'putCenso') { gravarCenso(ss, String(p.day), String(p.hosp), p.meta || {}, p.pats || []); return saida({ ok: true }); }
      if (p.op === 'putAcao') { gravarAcao(ss, String(p.hosp), String(p.id), p.body || {}); return saida({ ok: true }); }
    } finally { lock.releaseLock(); }
    return saida({ ok: false, erro: 'op' });
  } catch (err) { return saida({ ok: false, erro: String(err) }); }
}
function abaCensos(ss) { return aba(ss, 'censos', ['chave', 'dia', 'hospital', 'data', 'total', 'atualizado', 'json']); }
function abaAcoes(ss) { return aba(ss, 'acoes', ['chave', 'hospital', 'id', 'paciente', 'acao', 'atualizado', 'json']); }
function lerCensos(ss, desde) {
  return abaCensos(ss).getDataRange().getValues().slice(1).filter(r => r[0] && String(r[1]) >= desde).map(r => {
    let o = {}; try { o = JSON.parse(r[6]); } catch (x) {}
    return { day: String(r[1]), hosp: String(r[2]), meta: o.meta || {}, pats: o.pats || [] };
  });
}
function gravarCenso(ss, dia, hosp, meta, pats) {
  const sh = abaCensos(ss), v = sh.getDataRange().getValues(), chave = dia + '|' + hosp;
  const linha = [chave, dia, hosp, meta.date || '', pats.length, new Date(), JSON.stringify({ meta: meta, pats: pats })];
  const i = v.findIndex((r, k) => k > 0 && String(r[0]) === chave);
  if (i > 0) sh.getRange(i + 1, 1, 1, linha.length).setValues([linha]);
  else {
    sh.appendRow(linha);
    const limite = Utilities.formatDate(new Date(Date.now() - CENSOS_GUARDADOS * 864e5), 'America/Sao_Paulo', 'yyyyMMdd');
    const w = sh.getDataRange().getValues(), manter = w.filter((r, k) => k === 0 || String(r[1]) >= limite);
    if (manter.length < w.length) { sh.getRange(2, 1, w.length - 1, w[0].length).clearContent(); if (manter.length > 1) sh.getRange(2, 1, manter.length - 1, w[0].length).setValues(manter.slice(1)); }
  }
}
function lerBC(ss, nome, desde) {
  const sh = ss.getSheetByName(nome), out = {}; if (!sh) return out;
  sh.getDataRange().getValues().slice(1).forEach(r => {
    const dia = String(r[1]); if (!r[0] || dia < desde) return;
    let d = {}; try { d = JSON.parse(r[17]); } catch (x) { return; }
    (out[dia] = out[dia] || {})[String(r[0]).split('|')[1]] = { prev: d.prev || '', prevAntes: d.prevAntes || '', desfecho: d.desfecho || '', desfechoEm: d.desfechoEm || '',
      falta: d.falta || '', pend: d.pend || '', naosaiu: d.naosaiu || '', hd: d.hd || '', mot: d.mot || [], por: d.por || '', tocado: d.tocado || '', presente: d.presente !== false, dd: !!d.dd, visitador: (d.c && d.c.visitador) || '', cel: (d.c && d.c.cel) || '', atualizado: d.atualizado || '', dc: d.dc || null };
  });
  return out;
}
function lerAcoes(ss) {
  const out = {};
  abaAcoes(ss).getDataRange().getValues().slice(1).forEach(r => { if (!r[0]) return; try { out[String(r[0])] = JSON.parse(r[6]); } catch (x) {} });
  return out;
}
function gravarAcao(ss, hosp, id, body) {
  const sh = abaAcoes(ss), v = sh.getDataRange().getValues(), chave = hosp + '|' + id;
  const i = v.findIndex((r, k) => k > 0 && String(r[0]) === chave);
  let antigo = {}; if (i > 0) { try { antigo = JSON.parse(v[i][6]); } catch (x) {} }
  const doc = Object.assign({}, antigo, body, { em: new Date().toISOString() });
  const linha = [chave, hosp, id, doc.nome || '', doc.texto || '', new Date(), JSON.stringify(doc)];
  if (i > 0) sh.getRange(i + 1, 1, 1, linha.length).setValues([linha]); else sh.appendRow(linha);
}

/* censos que as equipes já montaram no bate censo entram sozinhos no painel
   (o censo arrastado no próprio painel tem prioridade e não é sobrescrito) */
function sincronizarCensos(ss, desde) {
  const sh = abaCensos(ss), v = sh.getDataRange().getValues(), idx = {};
  v.forEach((r, k) => { if (k > 0 && r[0]) idx[String(r[0])] = k; });
  Object.keys(HOSP_SUF).forEach(hosp => {
    const s = HOSP_SUF[hosp], pac = ss.getSheetByName(s ? 'pacientes ' + s : 'pacientes'); if (!pac) return;
    const porDia = {};
    pac.getDataRange().getValues().slice(1).forEach(r => {
      const dia = String(r[1]); if (!r[0] || dia < desde) return;
      let d = {}; try { d = JSON.parse(r[17]); } catch (x) { return; }
      if (d.presente === false || !d.c) return;
      const c = d.c;
      (porDia[dia] = porDia[dia] || []).push([String(r[0]).split('|')[1], c.unit || '', c.bed || '', c.name || '', c.conv || '', c.team || 'RP', c.dias == null ? null : c.dias, c.intern || '', c.prescriber || '', c.proc || '']);
    });
    Object.keys(porDia).forEach(dia => {
      const chave = dia + '|' + hosp, k = idx[chave];
      let meta = {};
      if (k > 0) { try { meta = (JSON.parse(v[k][6]) || {}).meta || {}; } catch (x) {} if (meta.fonte !== 'bate censo') return; }
      const pats = porDia[dia], json = JSON.stringify({ meta: { date: dia.slice(6, 8) + '/' + dia.slice(4, 6) + '/' + dia.slice(0, 4), fonte: 'bate censo' }, pats: pats });
      if (k > 0 && String(v[k][6]) === json) return;
      const linha = [chave, dia, hosp, dia.slice(6, 8) + '/' + dia.slice(4, 6) + '/' + dia.slice(0, 4), pats.length, new Date(), json];
      if (k > 0) sh.getRange(k + 1, 1, 1, linha.length).setValues([linha]);
      else { sh.appendRow(linha); idx[chave] = sh.getLastRow() - 1; v[idx[chave]] = linha; }
    });
  });
}
