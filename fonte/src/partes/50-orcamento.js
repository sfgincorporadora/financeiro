/* =========================================================
   Orçamento por etapa da obra e curva S (previsto x realizado)
   ========================================================= */
usarConfig('orcEtapas');
const ETAPAS_ORC_PADRAO = ['Serviços preliminares', 'Terraplenagem', 'Fundação', 'Estrutura e alvenaria', 'Cobertura', 'Instalações', 'Esquadrias', 'Revestimento e pintura', 'Acabamento e limpeza', 'Infraestrutura (rua, água, esgoto)'];
function etapasOrc() { return (S.orcEtapas && S.orcEtapas.etapas && S.orcEtapas.etapas.length) ? S.orcEtapas.etapas : ETAPAS_ORC_PADRAO; }
const orcQuadra = q => ((S.orcEtapas && S.orcEtapas.quadras) || {})[q] || {};
function salvarOrcEtapas(o) { S.orcEtapas = o; gravar(doc(db, 'config', 'orcEtapas'), o).catch(onErroEscrita); refreshView(); }
let orcQ = '';
function resumoOrc(qs) {
  const ets = etapasOrc(); const custo = t => qs.includes(t.quadra) && t.valor < 0 && secaoDe(t.destino) === 'op';
  const ts = S.list.filter(custo);
  const linhas = ets.map(e => { const orc = sum(qs, q => +(orcQuadra(q)[e]?.v || 0)); const x = ts.filter(t => t.etapaObra === e); const real = -sum(x.filter(pago), t => t.valor + Math.min(0, t.juros || 0)); const comp = -sum(x.filter(t => !pago(t)), t => t.valor); return { e, orc, real: real || 0, comp: comp || 0, tot: (real + comp) || 0 }; });
  const sem = ts.filter(t => !t.etapaObra || !ets.includes(t.etapaObra));
  const semL = { e: 'Sem etapa definida', orc: 0, real: -sum(sem.filter(pago), t => t.valor + Math.min(0, t.juros || 0)), comp: -sum(sem.filter(t => !pago(t)), t => t.valor), sem: true }; semL.tot = semL.real + semL.comp;
  return { linhas, semL, ts };
}
/* Curva S: previsto acumulado (cada etapa distribuída por igual entre os meses de início e fim) x realizado acumulado. */
function curvaS(qs, ts) {
  const prev = {}; let primeiro = '', ultimo = '';
  for (const q of qs) for (const [e, o] of Object.entries(orcQuadra(q))) { if (!(+o.v) || !o.ini) continue; const ms = mesesEntre(o.ini, o.fim && o.fim >= o.ini ? o.fim : o.ini); ms.forEach(m => { prev[m] = (prev[m] || 0) + o.v / ms.length; }); }
  const real = {}, comp = {}; const h = hoje().slice(0, 7);
  for (const t of ts) { const m = (dataRef(t) || '').slice(0, 7); if (!/^\d{4}-\d{2}$/.test(m)) continue; if (pago(t)) real[m] = (real[m] || 0) - t.valor - Math.min(0, t.juros || 0); else { const mm = m < h ? h : m; comp[mm] = (comp[mm] || 0) - t.valor; } }
  const todos = [...Object.keys(prev), ...Object.keys(real), ...Object.keys(comp)].sort(); primeiro = todos[0]; ultimo = todos[todos.length - 1];
  if (!primeiro) return null;
  const meses = mesesEntre(primeiro, ultimo < h ? h : ultimo); let a = 0, b = 0, c = null;
  return meses.map(m => { a += prev[m] || 0; if (m <= h) b += real[m] || 0; let proj = null; if (m >= h) { c = (c === null ? b : c) + (comp[m] || 0); proj = c; } return { m, prev: a, real: m <= h ? b : null, proj }; });
}
function svgCurva(pts) {
  if (!pts || pts.length < 2) return '<p class="muted">Defina meses de início e fim nas etapas para ver a curva.</p>';
  const W = 760, H = 240, L = 64, R = 12, T = 14, B = 28; const max = Math.max(1, ...pts.map(p => Math.max(p.prev, p.real || 0, p.proj || 0)));
  const x = i => L + i * (W - L - R) / (pts.length - 1), y = v => T + (H - T - B) * (1 - v / max);
  const linha = (k, cl) => { const p = pts.map((q, i) => q[k] === null ? null : `${x(i).toFixed(1)},${y(q[k]).toFixed(1)}`); let d = ''; p.forEach((s, i) => { if (s) d += (d && p[i - 1] ? ' L' : ' M') + s; }); return `<path class="${cl}" d="${d.trim()}"/>`; };
  const ticks = [0, .25, .5, .75, 1].map(f => `<line class="cs-grade" x1="${L}" x2="${W - R}" y1="${y(max * f)}" y2="${y(max * f)}"/><text class="cs-eixo" x="${L - 6}" y="${y(max * f) + 4}" text-anchor="end">${short(max * f)}</text>`).join('');
  const passo = Math.ceil(pts.length / 10);
  const xl = pts.map((p, i) => i % passo === 0 ? `<text class="cs-eixo" x="${x(i)}" y="${H - 8}" text-anchor="middle">${MES3[+p.m.slice(5, 7) - 1]}/${p.m.slice(2, 4)}</text>` : '').join('');
  return `<svg class="curva-s" viewBox="0 0 ${W} ${H}" role="img" aria-label="Curva S: previsto e realizado acumulados">${ticks}${xl}${linha('prev', 'cs-prev')}${linha('proj', 'cs-proj')}${linha('real', 'cs-real')}</svg>
    <p class="legenda-cs small"><span class="cs-l prev">Previsto acumulado</span><span class="cs-l real">Realizado (pago)</span><span class="cs-l proj">Com o que já está a pagar</span></p>`;
}
VIEWS.orcamento = {
  titulo: 'Orçamento por etapa',
  refresh(el) {
    const qs0 = E().quadras.filter(q => lotesDe(q).length); if (orcQ && !qs0.includes(orcQ)) orcQ = '';
    const qs = orcQ ? [orcQ] : qs0; const { linhas, semL, ts } = resumoOrc(qs);
    const T = k => sum(linhas, l => l[k]) + (k === 'orc' ? 0 : semL[k]);
    const nSem = ts.filter(t => !t.etapaObra).length;
    el.innerHTML = `<div class="filtros card"><label class="tgl">Quadra <select id="oQ">${opts(qs0, orcQ, 'Todas as quadras')}</select></label>${orcQ ? `<button class="btn primary rw" data-act="orcEd" data-q="${esc(orcQ)}">Editar orçamento da quadra ${esc(orcQ)}</button>` : '<span class="muted small">Escolha uma quadra para editar o orçamento.</span>'}<button class="btn ghost rw" data-act="orcEtapasEd">Etapas</button></div>
      <div class="grid kpis k4">${kpi('Orçado', brl(T('orc')), qs.length + ' quadra(s)', '')}${kpi('Realizado', brl(T('real')), T('orc') ? fPct.format(T('real') / T('orc')) + ' do orçado' : '', '')}${kpi('A pagar', brl(T('comp')), 'comprometido', '')}${kpi('Desvio projetado', brl(T('orc') - T('tot')), T('orc') ? (T('tot') > T('orc') ? 'acima do orçado' : 'dentro do orçado') : 'sem orçamento', T('tot') > T('orc') && T('orc') ? 'neg' : 'pos')}</div>
      <div class="card"><h2>Por etapa da obra</h2><div class="tbl-wrap"><table class="tbl"><thead><tr><th class="l">Etapa</th><th>Orçado</th><th>Pago</th><th>A pagar</th><th>Total</th><th class="l">Consumo do orçado</th><th>Saldo</th></tr></thead><tbody>
      ${[...linhas, semL].filter(l => l.orc || l.tot).map(l => `<tr class="${l.sem ? 'muted' : ''}"><td class="l">${esc(l.e)}${l.sem && nSem ? ` <a class="btn link sm" href="#/arrumar">classificar ${nSem}</a>` : ''}</td><td>${l.orc ? brl(l.orc) : '—'}</td><td>${brl(l.real)}</td><td>${brl(l.comp)}</td><td><b>${brl(l.tot)}</b></td><td class="l">${l.orc ? pctBar(l.tot / l.orc) : '—'}</td><td class="${l.orc && l.tot > l.orc ? 'neg' : ''}">${l.orc ? brl(l.orc - l.tot) : '—'}</td></tr>`).join('') || '<tr><td class="l muted" colspan="7">Sem orçamento nem custos de obra nestas quadras.</td></tr>'}</tbody></table></div>
      <p class="muted small">Entram os custos de obra (destinos operacionais) das quadras. Para separar por etapa, escolha a "Etapa da obra" no lançamento ou em Arrumar lançamentos.</p></div>
      <div class="card"><h2>Curva S</h2>${svgCurva(curvaS(qs, ts))}</div>`;
    $('#oQ', el).onchange = e => { orcQ = e.target.value; this.refresh(el); };
  }
};
ACOES.orcEd = d => {
  const q = d.q; const o = orcQuadra(q); const ets = etapasOrc(); const nL = lotesDe(q).length || 1;
  const outras = E().quadras.filter(x => x !== q && Object.keys(orcQuadra(x)).length);
  const f = folha(`Orçamento por etapa · quadra ${q}`, `<form class="form" id="fx"><p class="muted small">${nL} lote(s). Informe o valor total da quadra por etapa e os meses previstos de início e fim (para a curva S).</p>
    ${outras.length ? `<div class="acts left"><label class="tgl">Copiar de <select id="oCp">${opts(outras, '', 'outra quadra…')}</select></label><span class="muted small">ajusta pelo número de lotes</span></div>` : ''}
    <div class="orc-ed">${ets.map((e, i) => `<div class="orc-row"><span>${esc(e)}</span><input name="v${i}" inputmode="decimal" placeholder="R$" value="${o[e]?.v ? fNum2.format(o[e].v) : ''}"><input name="i${i}" type="month" value="${esc(o[e]?.ini || '')}"><input name="f${i}" type="month" value="${esc(o[e]?.fim || '')}"></div>`).join('')}</div>
    <div class="preview" id="oTot"></div>
    <div class="acts"><button type="button" class="btn" data-close>Cancelar</button><button class="btn primary">Salvar</button></div></form>`, { largo: true });
  const F = f.el.querySelector('#fx');
  const tot = () => { const s = ets.reduce((acc, _, i) => acc + (parseBR(F['v' + i].value) || 0), 0); $('#oTot', F).innerHTML = `Total ${brl2(s)} · ${brl2(s / nL)} por lote`; };
  F.addEventListener('input', tot); tot();
  const cp = F.querySelector('#oCp'); if (cp) cp.onchange = () => { const src = orcQuadra(cp.value), fator = nL / (lotesDe(cp.value).length || 1); ets.forEach((e, i) => { if (src[e]) { F['v' + i].value = src[e].v ? fNum2.format(r2(src[e].v * fator)) : ''; F['i' + i].value = src[e].ini || ''; F['f' + i].value = src[e].fim || ''; } }); tot(); };
  F.onsubmit = e => { e.preventDefault(); const n = {}; ets.forEach((et, i) => { const v = parseBR(F['v' + i].value) || 0; if (v || F['i' + i].value) n[et] = { v: r2(v), ini: F['i' + i].value, fim: F['f' + i].value }; }); const all = clone(S.orcEtapas || {}); all.quadras = { ...(all.quadras || {}), [q]: n }; if (!all.etapas) all.etapas = ets; salvarOrcEtapas(all); f.fechar(); toast('Orçamento salvo'); };
};
ACOES.orcEtapasEd = () => {
  const f = folha('Etapas da obra', `<form class="form" id="fx"><label>Uma etapa por linha<textarea name="e" rows="12">${esc(etapasOrc().join('\n'))}</textarea></label><p class="muted small">Renomear uma etapa não muda os lançamentos já classificados com o nome antigo.</p><div class="acts"><button type="button" class="btn" data-close>Cancelar</button><button class="btn primary">Salvar</button></div></form>`);
  f.el.querySelector('#fx').onsubmit = e => { e.preventDefault(); const ets = [...new Set(e.target.e.value.split('\n').map(s => s.trim()).filter(Boolean))]; if (!ets.length) return; salvarOrcEtapas({ ...(S.orcEtapas || {}), etapas: ets }); f.fechar(); };
};
