/* =========================================================
   Indicadores de obras: rentabilidade (margens, ROI, TIR, MOIC),
   caixa (exposição máxima, payback, prazo) e comercial (VSO, repasse,
   inadimplência), por quadra, por obra e total.
   ========================================================= */
let indEsc = '', indAdm = false;
/* TIR mensal por bisseção (fluxos mensais, o primeiro é o mês 0). */
function tirMensal(f) {
  if (!f.some(v => v < 0) || !f.some(v => v > 0)) return null;
  const vpl = r => f.reduce((s, v, i) => s + v / Math.pow(1 + r, i), 0);
  // procura a troca de sinal mais próxima de zero (fluxos com mais de uma troca de sinal têm várias raízes)
  let melhor = null;
  for (const dir of [1, -1]) { let a = 0, fa = vpl(0); for (let k = 1; k <= 400; k++) { const b = dir * k * 0.0025; if (b <= -0.95) break; const fb = vpl(b); if (fa * fb <= 0) { melhor = melhor === null || Math.abs(b) < Math.abs(melhor[1]) ? [a, b] : melhor; break; } a = b; fa = fb; } }
  if (!melhor) return null;
  let [lo, hi] = melhor[0] < melhor[1] ? melhor : [melhor[1], melhor[0]]; let flo = vpl(lo);
  for (let k = 0; k < 100; k++) { const m = (lo + hi) / 2, fm = vpl(m); if (Math.abs(hi - lo) < 1e-7) return m; if (flo * fm <= 0) hi = m; else { lo = m; flo = fm; } }
  return (lo + hi) / 2;
}
const difMeses = (a, b) => { const [y1, m1] = a.split('-').map(Number), [y2, m2] = b.split('-').map(Number); return (y2 - y1) * 12 + (m2 - m1); };
function indicadoresDe(qs) {
  const set = new Set(qs); const h = hoje();
  const ts = S.list.filter(t => set.has(t.quadra) && secaoDe(t.destino) !== 'aporte');
  const real = ts.filter(pago), ab = ts.filter(t => !pago(t));
  const v = t => t.valor + (pago(t) ? (t.juros || 0) : 0);
  const recReal = sum(real.filter(t => v(t) > 0), v), recAb = sum(ab.filter(t => t.valor > 0), t => t.valor);
  const custoReal = -sum(real.filter(t => v(t) < 0), v), custoAb = -sum(ab.filter(t => t.valor < 0), t => t.valor);
  const adm = indAdm && typeof admQuadra === 'function' ? sum(qs, q => admQuadra(q)) : 0;
  const receita = recReal + recAb, custo = custoReal + custoAb + adm, lucro = receita - custo;
  // margens no padrão da viabilidade (operacional x líquida)
  const C = byDest(ts);
  const iC = indicadores(C);
  const prev = qs.reduce((o, q) => { const p = previstoLote(q), n = lotesDe(q).length; for (const k in p) o[k] = (o[k] || 0) + p[k] * n; return o; }, {});
  const iP = indicadores(prev); const custoPrev = -(iP.op + iP.non);
  // fluxo mensal (realizado no mês do pagamento; aberto no vencimento, vencidos no mês atual)
  const fl = {}; for (const t of ts) { const d = pago(t) ? (t.pag || t.venc) : (t.venc < h ? h : t.venc); if (!d) continue; const m = d.slice(0, 7); fl[m] = (fl[m] || 0) + v(t); }
  const ms = Object.keys(fl).sort(); let tir = null, expo = 0, expoMes = '', payback = '', prazo = null; const curva = [];
  if (ms.length) { const ini = ms[0], fim = ms[ms.length - 1]; const n = difMeses(ini, fim) + 1; const f = Array(n).fill(0); ms.forEach(m => { f[difMeses(ini, m)] += fl[m]; });
    if (adm) f[0] -= adm; tir = tirMensal(f); let ac = 0;
    f.forEach((x, i) => { ac += x; const m = addMonths(ini + '-01', i).slice(0, 7); curva.push({ m, ac, real: m <= h.slice(0, 7) }); if (ac < expo) { expo = ac; expoMes = m; } });
    const iMin = curva.findIndex(c => c.m === expoMes); const pb = curva.findIndex((c, i) => i > iMin && c.ac >= 0); payback = pb >= 0 ? curva[pb].m : ''; prazo = n; }
  // comercial
  const ls = qs.flatMap(q => lotesDe(q).map(l => infoVendaLote(q, l)));
  const vend = ls.filter(x => x.etapa !== 'disp'), rep = ls.filter(x => x.etapa === 'recebido');
  const dur = ls.map(x => { const hs = ((x.v && x.v.hist) || []).filter(e => e.data); const a = hs.find(e => e.etapa !== 'disp'), b = hs.find(e => e.etapa === 'recebido'); return a && b ? (new Date(b.data) - new Date(a.data)) / 864e5 : null; }).filter(x => x !== null && x >= 0);
  const vencRec = sum(ab.filter(t => t.valor > 0 && t.venc < addDays(h, -5)), t => t.valor);
  const area = sum(qs.flatMap(q => lotesDe(q).map(l => +String((typeof loteInfo === 'function' ? loteInfo(q, l).construida : '') || '').replace(',', '.') || 0)), x => x);
  return { qs, n: ls.length, vend: vend.length, rep: rep.length, receita, recReal, recAb, custo, custoReal, custoAb, adm, lucro,
    margem: receita ? lucro / receita : null, mop: iC.mop, ml: iC.ml, roi: custo ? lucro / custo : null, moic: custo ? receita / custo : null,
    tirM: tir, tirA: tir === null ? null : Math.pow(1 + tir, 12) - 1, expo, expoMes, payback, prazo, curva,
    custoLote: ls.length ? custo / ls.length : null, ticket: vend.length ? receita / vend.length : null, custoM2: area ? custo / area : null, area,
    vso: ls.length ? vend.length / ls.length : null, repasseDias: dur.length ? sum(dur, x => x) / dur.length : null,
    inad: receita ? vencRec / (recAb || 1) : null, vencRec, recebidoPct: receita ? recReal / receita : null,
    prevMargem: iP.rec ? (iP.rec + iP.op + iP.non) / iP.rec : null, prevRoi: iP.roi, custoPrev, desvioCusto: custoPrev ? (custo - adm) / custoPrev - 1 : null, avanco: custoPrev ? custoReal / custoPrev : null };
}
VIEWS.indicadores = {
  titulo: 'Indicadores das obras',
  refresh(el) {
    const obras = E().obras.filter(o => quadrasDaObra(o).some(q => lotesDe(q).length));
    const escopos = [['', 'Todas as obras'], ...obras.map(o => ['O:' + o, o]), ...E().quadras.filter(q => lotesDe(q).length).map(q => ['Q:' + q, 'Quadra ' + q])];
    if (!escopos.some(e => e[0] === indEsc)) indEsc = '';
    const qsDe = e => e.startsWith('Q:') ? [e.slice(2)] : e.startsWith('O:') ? quadrasDaObra(e.slice(2)).filter(q => lotesDe(q).length) : E().quadras.filter(q => lotesDe(q).length);
    const I = indicadoresDe(qsDe(indEsc));
    const linhas = [...(obras.length > 1 ? obras.map(o => ({ nome: o, obra: true, i: indicadoresDe(quadrasDaObra(o).filter(q => lotesDe(q).length)) })) : []), ...E().quadras.filter(q => lotesDe(q).length).map(q => ({ nome: q, i: indicadoresDe([q]) }))];
    const pp = v => pct(v), mes = m => m ? mesNome(m) : '—';
    const dlt = (a, b) => a === null || b === null || !isFinite(a) || !isFinite(b) ? '' : `<small class="${a >= b ? 'pos' : 'neg'}">previsto ${pct(b)}</small>`;
    el.innerHTML = `<div class="filtros card"><select id="iEsc">${opts(escopos, indEsc)}</select><label class="tgl"><input type="checkbox" id="iAdm" ${indAdm ? 'checked' : ''}> Incluir administração rateada</label><span class="muted small">Projetado = realizado + o que falta receber e pagar</span></div>
      <div class="lbl">Rentabilidade</div>
      <div class="grid kpis k4">${kpi('Lucro projetado', brl(I.lucro), `receita ${short(I.receita)} − custo ${short(I.custo)}`, cls(I.lucro))}${kpi('Margem líquida', pp(I.margem), dlt(I.margem, I.prevMargem) || 'lucro ÷ receita', cls(I.margem))}${kpi('ROI', pp(I.roi), dlt(I.roi, I.prevRoi) || 'lucro ÷ custo', cls(I.roi))}${kpi('TIR', I.tirA === null ? '—' : pct(I.tirA) + ' a.a.', I.tirM === null ? 'sem fluxo suficiente' : pct(I.tirM) + ' ao mês', cls(I.tirA))}</div>
      <div class="grid kpis k4">${kpi('MOIC', I.moic === null ? '—' : I.moic.toLocaleString('pt-BR', { maximumFractionDigits: 2 }) + 'x', 'receita ÷ custo: quanto volta para cada R$ 1 gasto', '')}${kpi('Margem operacional', pp(I.mop), 'antes de despesas não operacionais', '')}${kpi('Custo por lote', I.custoLote === null ? '—' : brl(I.custoLote), I.custoM2 ? `${brl(I.custoM2)} por m² construído` : 'cadastre a área construída do lote para ver por m²', '')}${kpi('Desvio de custo', I.desvioCusto === null ? '—' : (I.desvioCusto > 0 ? '+' : '') + pct(I.desvioCusto), I.custoPrev ? `orçado ${short(I.custoPrev)}` : 'sem orçamento', I.desvioCusto > 0.02 ? 'neg' : I.desvioCusto !== null ? 'pos' : '')}</div>
      <div class="lbl">Caixa</div>
      <div class="grid kpis k4">${kpi('Exposição máxima', brl(-I.expo), I.expoMes ? 'pico de dinheiro aplicado em ' + mes(I.expoMes) : '', 'neg')}${kpi('Payback', I.payback ? mes(I.payback) : '—', I.payback ? 'quando o caixa acumulado volta a ficar positivo' : 'o caixa ainda não volta ao positivo', '')}${kpi('Prazo da operação', I.prazo ? I.prazo + ' meses' : '—', 'do primeiro ao último movimento', '')}${kpi('Avanço financeiro', pp(I.avanco), 'custo pago ÷ custo orçado', '')}</div>
      <div class="lbl">Comercial</div>
      <div class="grid kpis k4">${kpi('VSO (vendidos)', pp(I.vso), `${I.vend} de ${I.n} lotes · ${I.rep} com repasse recebido`, '')}${kpi('Ticket médio', I.ticket === null ? '—' : brl(I.ticket), 'receita por lote vendido', '')}${kpi('Prazo médio de repasse', I.repasseDias === null ? '—' : Math.round(I.repasseDias) + ' dias', 'da venda ao repasse recebido', '')}${kpi('Inadimplência', pp(I.inad), I.vencRec ? `${brl(I.vencRec)} vencidos há mais de 5 dias` : 'nada vencido', I.vencRec ? 'neg' : '')}</div>
      <div class="card"><h2>Caixa acumulado da operação <span class="hint">a curva J: quanto dinheiro fica aplicado e quando volta</span></h2><div class="chartbox"><canvas id="cJ"></canvas></div></div>
      <div class="card"><h2>Comparativo</h2><div class="tbl-wrap"><table class="tbl"><thead><tr><th class="l">${obras.length > 1 ? 'Obra / quadra' : 'Quadra'}</th><th>Lotes</th><th>VSO</th><th>Receita</th><th>Custo</th><th>Lucro</th><th>Margem líq.</th><th>ROI</th><th>TIR a.a.</th><th>MOIC</th><th>Exposição máx.</th><th>Payback</th><th>Custo/lote</th><th>Desvio custo</th></tr></thead><tbody>
      ${linhas.map(({ nome, obra, i }) => `<tr class="${obra ? 'sec' : 'click'}" ${obra ? '' : `data-act="ir" data-href="#/obras/${encodeURIComponent(nome)}"`}><td class="l"><b>${esc(nome)}</b></td><td>${i.n}</td><td>${pp(i.vso)}</td><td>${short(i.receita)}</td><td>${short(i.custo)}</td><td class="${cls(i.lucro)}">${short(i.lucro)}</td><td>${pp(i.margem)}</td><td>${pp(i.roi)}</td><td>${i.tirA === null ? '—' : pct(i.tirA)}</td><td>${i.moic === null ? '—' : i.moic.toLocaleString('pt-BR', { maximumFractionDigits: 2 }) + 'x'}</td><td class="neg">${short(-i.expo)}</td><td>${i.payback ? mesCurto(i.payback) : '—'}</td><td>${i.custoLote === null ? '—' : short(i.custoLote)}</td><td class="${i.desvioCusto > 0.02 ? 'neg' : ''}">${i.desvioCusto === null ? '—' : (i.desvioCusto > 0 ? '+' : '') + pct(i.desvioCusto)}</td></tr>`).join('')}</tbody></table></div>
      <p class="muted small">Receita e custo projetados = já recebido/pago + o que está lançado a receber/pagar. Lotes ainda sem venda lançada entram só com o custo, então o lucro final tende a ser maior. TIR calculada sobre o fluxo mensal da operação (custos e receitas das quadras); impostos e despesas lançados na empresa ficam de fora, a não ser a administração rateada se marcada.</p></div>`;
    $('#iEsc', el).onchange = e => { indEsc = e.target.value; this.refresh(el); };
    $('#iAdm', el).onchange = e => { indAdm = e.target.checked; this.refresh(el); };
    const cv = I.curva; if (cv.length) grafico('cJ', { type: 'line', data: { labels: cv.map(c => mesCurto(c.m)), datasets: [
      { label: 'Realizado', data: cv.map(c => c.real ? c.ac : null), borderColor: css('--navy') || '#1F3A6B', backgroundColor: 'transparent', borderWidth: 2, pointRadius: 0, tension: 0 },
      { label: 'Projetado', data: cv.map((c, i) => !c.real || (cv[i + 1] && !cv[i + 1].real) ? c.ac : null), borderColor: css('--muted') || '#8A96A8', borderDash: [5, 4], backgroundColor: 'transparent', borderWidth: 2, pointRadius: 0, tension: 0 }] },
      options: { responsive: true, maintainAspectRatio: false, animation: false, interaction: { mode: 'index', intersect: false }, plugins: { legend: { display: true, labels: { boxWidth: 12, color: css('--muted') } }, tooltip: { callbacks: { label: c => c.dataset.label + ': ' + brl(c.parsed.y) } } },
        scales: { x: { ticks: { color: css('--muted'), maxTicksLimit: 10 }, grid: { display: false } }, y: { ticks: { color: css('--muted'), callback: v => short(v) }, grid: { color: css('--line') } } } } });
  }
};
