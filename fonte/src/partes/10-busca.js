/* =========================================================
   Busca global e perguntas simples
   Ex.: "quanto paguei de cimento na K3 em 2025", "a pagar Mendonça", "lote K3 19"
   ========================================================= */
const BUSCA_EXTRA = []; // outros módulos acrescentam resultados: fn(qNorm) => [{ titulo, sub, href }]
const PALAVRAS_VAZIAS = new Set('quanto quantos quantas qual quais eu nos de do da dos das na no nas nos em com para pra os as o a e total foi foram que meu minha ja ate desde por ao aos mais todos todas lancamentos lancamento gastos gasto valor valores tenho temos ultimo ultimos'.split(' '));
const MESES_N = MESES.map(normTxt);
function interpretarBusca(q) {
  const n = ' ' + normTxt(q) + ' '; const f = { texto: [] }; let resto = n;
  const tira = s => { resto = resto.replace(' ' + s + ' ', ' '); };
  // situação
  if (/ (paguei|pagamos|pago|pagos|pagas|gastei|gastamos) /.test(n)) { f.st = 'pago'; f.sinal = -1; }
  if (/ (recebi|recebemos|recebido|recebidos|entrou|entraram) /.test(n)) { f.st = 'pago'; f.sinal = 1; }
  if (/ a pagar | devo | vence | vencem | vencendo /.test(n)) { f.st = 'aberto'; f.sinal = -1; }
  if (/ a receber /.test(n)) { f.st = 'aberto'; f.sinal = 1; }
  if (/ vencid/.test(n)) { f.st = 'vencido'; f.sinal = f.sinal || -1; }
  ['paguei', 'pagamos', 'pago', 'pagos', 'pagas', 'gastei', 'gastamos', 'recebi', 'recebemos', 'recebido', 'recebidos', 'entrou', 'entraram', 'a pagar', 'a receber', 'devo', 'vence', 'vencem', 'vencendo', 'vencidas', 'vencidos', 'vencida', 'vencido'].forEach(tira);
  // período
  const h = hoje(); const ano = h.slice(0, 4);
  if (/ este mes | neste mes /.test(n)) { f.de = h.slice(0, 7); f.ate = h.slice(0, 7); ['este mes', 'neste mes'].forEach(tira); }
  else if (/ mes passado /.test(n)) { const m = addMonths(h.slice(0, 7) + '-01', -1).slice(0, 7); f.de = m; f.ate = m; tira('mes passado'); }
  else if (/ ano passado /.test(n)) { f.de = (+ano - 1) + '-01'; f.ate = (+ano - 1) + '-12'; tira('ano passado'); }
  else if (/ este ano | neste ano /.test(n)) { f.de = ano + '-01'; f.ate = ano + '-12'; ['este ano', 'neste ano'].forEach(tira); }
  else {
    const mm = MESES_N.findIndex(m => n.includes(' ' + m + ' ') || n.includes(' ' + m.slice(0, 3) + ' '));
    const aa = (n.match(/ (20\d{2}) /) || [])[1];
    if (mm >= 0) { const y = aa || ano; const k = `${y}-${String(mm + 1).padStart(2, '0')}`; f.de = k; f.ate = k; tira(MESES_N[mm]); tira(MESES_N[mm].slice(0, 3)); }
    else if (aa) { f.de = aa + '-01'; f.ate = aa + '-12'; }
    if (aa) tira(aa);
  }
  // obra, quadra e lote
  for (const o of E().obras) if (n.includes(' ' + normTxt(o) + ' ')) { f.obra = o; tira(normTxt(o)); }
  for (const qd of E().quadras) { const k = normTxt(qd); if (k && (n.includes(' ' + k + ' ') || n.includes(' quadra ' + k + ' '))) { f.quadra = qd; tira('quadra ' + k); tira(k); } }
  const lm = resto.match(/ (?:lote|lt) (\d+[\d ]*?) /); if (lm) { f.lote = lm[1].trim().replace(/ /g, '.'); tira('lote ' + lm[1]); tira('lt ' + lm[1]); }
  else if (f.quadra) { const lx = resto.match(/ (\d{1,3}(?: \d)?) /); if (lx && lotesDe(f.quadra).some(l => normTxt(l) === lx[1])) { f.lote = lotesDe(f.quadra).find(l => normTxt(l) === lx[1]); tira(lx[1]); } }
  // destino e fornecedor (frases inteiras)
  const dests = todosDestinos().slice().sort((a, b) => b.length - a.length);
  for (const d of dests) { const k = normTxt(d); if (k && resto.includes(' ' + k + ' ')) { f.destino = d; tira(k); break; } }
  const forns = [...new Set(S.list.map(t => t.forn).filter(Boolean))].sort((a, b) => b.length - a.length);
  for (const fo of forns) { const k = normTxt(fo); if (k.length > 2 && resto.includes(' ' + k + ' ')) { f.forn = fo; tira(k); break; } }
  f.texto = resto.trim().split(' ').filter(w => w && !PALAVRAS_VAZIAS.has(w) && !/^\d{1,2}$/.test(w));
  f.pergunta = /^ (quanto|quantos|quantas|qual|quais) /.test(n) || /\?/.test(q) || !!(f.st || f.de || f.destino);
  return f;
}
function aplicarBusca(f) {
  const h = hoje();
  return S.list.filter(t => {
    if (f.st === 'pago' && !pago(t)) return false;
    if (f.st === 'aberto' && pago(t)) return false;
    if (f.st === 'vencido' && (pago(t) || !(t.venc < h))) return false;
    if (f.sinal && Math.sign(t.valor) !== f.sinal) return false;
    const d = (dataRef(t) || '').slice(0, 7); if (f.de && (d < f.de || d > f.ate)) return false;
    if (f.obra && obraT(t) !== f.obra) return false; if (f.quadra && t.quadra !== f.quadra) return false; if (f.lote && String(t.lote) !== String(f.lote)) return false;
    if (f.destino && t.destino !== f.destino) return false; if (f.forn && t.forn !== f.forn) return false;
    if (f.texto.length) { const alvo = ' ' + normTxt([t.item, t.forn, t.obs, t.nfNum, t.destino].join(' ')) + ' '; if (!f.texto.every(w => alvo.includes(w))) return false; }
    return true;
  });
}
function respostaBusca(f, ts) {
  if (!ts.length) return 'Nenhum lançamento encontrado com esses filtros.';
  const tot = sum(ts, t => t.valor + (pago(t) ? (t.juros || 0) : 0));
  const verbo = f.st === 'pago' ? (f.sinal === 1 ? 'Recebido' : 'Pago') : f.st === 'aberto' ? (f.sinal === 1 ? 'A receber' : 'A pagar') : f.st === 'vencido' ? 'Vencido em aberto' : 'Total';
  const onde = [f.texto.length ? `"${f.texto.join(' ')}"` : '', f.destino ? f.destino.toLowerCase() : '', f.forn ? 'de ' + f.forn : '', f.quadra ? 'na quadra ' + f.quadra + (f.lote ? ', lote ' + f.lote : '') : '', f.obra ? 'na obra ' + f.obra : '', f.de ? (f.de === f.ate ? 'em ' + mesNome(f.de) : `de ${mesNome(f.de)} a ${mesNome(f.ate)}`) : ''].filter(Boolean).join(' ');
  const qt = sum(ts.filter(t => t.qtd), t => t.qtd); const un = [...new Set(ts.filter(t => t.qtd).map(t => t.unid))];
  const preco = qt > 0 && un.length === 1 ? ` · ${fNum2.format(qt)} ${un[0]}, média de ${brl2(Math.abs(sum(ts.filter(t => t.qtd), t => t.valor)) / qt)} por ${un[0]}` : '';
  return `<b>${verbo}${onde ? ' · ' + esc(onde) : ''}:</b> <span class="${cls(tot)} fig">${brl2(tot)}</span> em ${ts.length} lançamento(s)${esc(preco)}`;
}
ACOES.busca = () => {
  const f = folha('Buscar ou perguntar', `<div class="busca"><input type="search" id="bq" placeholder="Ex.: quanto paguei de cimento na K3 em 2025" autocomplete="off" enterkeyhint="search">
    <p class="muted small">Procura em lançamentos, fornecedores, lotes, contratos e compras. Entende perguntas com: pago, recebido, a pagar, a receber, vencidas; mês ou ano; quadra e lote; destino; fornecedor.</p><div id="bRes"></div></div>`, { largo: true });
  const inp = f.el.querySelector('#bq'), res = f.el.querySelector('#bRes'); setTimeout(() => inp.focus(), 60);
  let tmr = 0;
  const roda = () => {
    const q = inp.value.trim(); if (q.length < 2) { res.innerHTML = ''; return; }
    const fl = interpretarBusca(q); const ts = aplicarBusca(fl).sort((a, b) => (dataRef(b) || '').localeCompare(dataRef(a) || ''));
    const qn = normTxt(q); const out = [];
    if (fl.pergunta || ts.length) out.push(`<div class="busca-resp">${respostaBusca(fl, ts)}</div>`);
    const forns = statsFornecedores().filter(x => normTxt(x.nome).includes(qn)).slice(0, 5);
    if (forns.length) out.push(`<div class="lbl">Fornecedores</div>${forns.map(x => `<button class="busca-item-r" data-act="forn" data-n="${esc(x.nome)}"><b>${esc(x.nome)}</b><small>${x.n} lançamentos · pago ${brl(x.pago)}</small></button>`).join('')}`);
    if (fl.quadra) out.push(`<div class="lbl">Lote</div><a class="busca-item-r" href="#/obras/${encodeURIComponent(fl.quadra)}${fl.lote ? '/' + encodeURIComponent(fl.lote) : ''}" data-close><b>Quadra ${esc(fl.quadra)}${fl.lote ? ', lote ' + esc(fl.lote) : ''}</b><small>abrir resumo financeiro</small></a>${fl.lote ? `<a class="busca-item-r" href="#/etapas/${encodeURIComponent(fl.quadra)}/${encodeURIComponent(fl.lote)}" data-close><b>Fluxo e prazos do lote</b><small>etapas e documentos</small></a>` : ''}`);
    const extra = BUSCA_EXTRA.flatMap(fn => { try { return fn(qn); } catch (e) { return []; } }).slice(0, 8);
    if (extra.length) out.push(`<div class="lbl">Outros</div>${extra.map(x => `<a class="busca-item-r" href="${esc(x.href)}" data-close><b>${esc(x.titulo)}</b><small>${esc(x.sub || '')}</small></a>`).join('')}`);
    if (ts.length) out.push(`<div class="lbl">Lançamentos (${ts.length}${ts.length > 40 ? ', mostrando 40' : ''})</div><div class="tlist">${ts.slice(0, 40).map(t => linhaTitulo(t, { mostrarPago: true })).join('')}</div>`);
    res.innerHTML = out.join('') || '<p class="muted">Nada encontrado.</p>';
  };
  inp.oninput = () => { clearTimeout(tmr); tmr = setTimeout(roda, 180); };
  inp.onkeydown = e => { if (e.key === 'Enter') { e.preventDefault(); roda(); } };
  res.addEventListener('click', e => { if (e.target.closest('[data-act="abrir"],[data-act="forn"]')) f.fechar(); });
};
document.addEventListener('keydown', e => { if (e.key === '/' && S.user && !/input|textarea|select/i.test(document.activeElement?.tagName || '') && !$('.sheet-wrap')) { e.preventDefault(); ACOES.busca(); } });
