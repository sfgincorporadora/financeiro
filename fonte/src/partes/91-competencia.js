/* =========================================================
   DRE por competência (incorporação): receita das unidades vendidas
   reconhecida pelo avanço da obra (POC = custo incorrido ÷ custo total
   previsto do lote), a partir da data da venda. Custo dos lotes não
   vendidos fica em estoque. Despesas e demais receitas pelo vencimento.
   ========================================================= */
let _pocCache = { L: null, base: null };
const DESP_VENDA = ['Corretagem', 'Impostos'];
function basePoc(L) {
  const H = hoje(); if (_pocCache.L === L && _pocCache.ref === S.list && _pocCache.h === H) return _pocCache.base;
  const lotes = {}; const ger = {};
  const ehVenda = t => secaoDe(t.destino) === 'rec' && !FIN_REC.includes(t.destino);
  const valida = t => t.quadra && t.quadra !== '-' && E().quadras.includes(t.quadra) && secaoDe(t.destino) !== 'aporte' && !t.fin;
  const temLote = t => t.lote && t.lote !== '-';
  // quadra vendida inteira (venda sem lote): a quadra toda vira uma unidade só
  const inteira = new Set(L.filter(t => valida(t) && ehVenda(t) && !temLote(t)).map(t => t.quadra));
  const dv = {}; // corretagem e impostos das vendas, por quadra: reconhecidos junto com a receita
  const un = k => lotes[k] = lotes[k] || { vgv: 0, venda: '', ev: [], prev: 0, abertos: 0, ult: '', ultP: '' };
  for (const t of L) {
    if (!valida(t)) continue;
    const k = inteira.has(t.quadra) ? t.quadra + '|*' : t.quadra + '|' + t.lote;
    const comLote = inteira.has(t.quadra) || temLote(t);
    if (ehVenda(t)) { if (!comLote) continue; const o = un(k); o.vgv += t.valor; const d = t.venc || t.pag; if (d && (!o.venda || d < o.venda)) o.venda = d;
      if (pago(t)) { if ((t.pag || t.venc) > o.ult) o.ult = t.pag || t.venc; } else o.abertos++;
      const dr = pago(t) ? (t.pag || t.venc) : (t.venc && t.venc > H ? t.venc : H); if (dr > o.ultP) o.ultP = dr; continue; }
    if (DESP_VENDA.includes(t.destino)) { const q = dv[t.quadra] = dv[t.quadra] || {}; q[t.destino] = (q[t.destino] || 0) + t.valor; continue; }
    if (secaoDe(t.destino) !== 'op' || !t.valor) continue;
    const ev = { d: pago(t) ? (t.pag || t.venc) : (t.venc && t.venc > H ? t.venc : H), v: -t.valor, dest: t.destino, pago: pago(t) };
    if (comLote) un(k).ev.push(ev); else (ger[t.quadra] = ger[t.quadra] || []).push(ev);
  }
  // custo geral da quadra dividido entre os lotes (sem lotes cadastrados: a quadra vira uma unidade)
  for (const q in ger) { const ls = lotesDe(q); const ks = ls.length ? ls.map(l => q + '|' + l) : [q + '|*'];
    for (const k of ks) { const o = un(k); for (const e of ger[q]) o.ev.push({ ...e, v: e.v / ks.length }); } }
  for (const k in lotes) { const o = lotes[k]; const [q, l] = k.split('|');
    // data da venda: primeira etapa de venda registrada ou primeiro vencimento de recebimento
    const v = ((S.vendas && S.vendas.lotes) || {})[k]; const h = ((v && v.hist) || []).find(e => e.etapa && e.etapa !== 'disp' && e.data); if (h && (!o.venda || h.data < o.venda)) o.venda = h.data;
    const tot = sum(o.ev, e => e.v); const pl = l === '*' ? {} : previstoLote(q); const prevCusto = -sum(Object.keys(pl).filter(d => secaoDe(d) === 'op'), d => pl[d] || 0);
    o.custoTotal = Math.max(tot, prevCusto || 0) || tot;
    // conclusão: último recebimento ou pagamento (realizado ou previsto) -> 100% reconhecido (ajuste final da estimativa)
    const pagos = o.ev.filter(e => e.pago); o.abertos += o.ev.length - pagos.length;
    for (const e of pagos) if (e.d > o.ult) o.ult = e.d;
    for (const e of o.ev) if (e.d > o.ultP) o.ultP = e.d;
    o.fim = o.vgv && o.ultP ? o.ultP : '';
    o.ev.sort((a, b) => a.d.localeCompare(b.d)); }
  const vgvQ = {}; for (const k in lotes) { const q = k.split('|')[0]; vgvQ[q] = (vgvQ[q] || 0) + lotes[k].vgv; }
  for (const q in dv) if (!(vgvQ[q] > 0)) delete dv[q];
  Object.defineProperty(lotes, '_dv', { value: { dv, vgvQ } });
  _pocCache = { L, ref: S.list, h: H, base: lotes }; return lotes;
}
/* Receita e custo reconhecidos acumulados até a data D. */
function pocAte(base, D) {
  const rec = {}, custo = {}, vend = {}, rq = {}; let r = 0;
  for (const k in base) { const o = base[k]; if (!o.venda || o.venda > D || !o.vgv) continue;
    let inc = 0; const porDest = {}; for (const e of o.ev) { if (e.d > D) break; inc += e.v; porDest[e.dest] = (porDest[e.dest] || 0) + e.v; }
    const poc = o.fim && D >= o.fim ? 1 : o.custoTotal > 0 ? Math.min(1, inc / o.custoTotal) : 0; r += o.vgv * poc; const q = k.split('|')[0]; rq[q] = (rq[q] || 0) + o.vgv * poc;
    for (const d in porDest) custo[d] = (custo[d] || 0) - porDest[d]; }
  const { dv, vgvQ } = base._dv; for (const q in dv) { const f = Math.min(1, (rq[q] || 0) / vgvQ[q]); for (const d in dv[q]) vend[d] = (vend[d] || 0) + dv[q][d] * f; }
  rec['Venda'] = r; return { rec, custo, vend };
}
function dreCompetencia(L, ini, fim) {
  const base = basePoc(L); const antes = addDays(ini, -1);
  const A = pocAte(base, fim), B = pocAte(base, antes);
  const dif = (a, b) => { const o = {}; for (const k of new Set([...Object.keys(a), ...Object.keys(b)])) { const v = (a[k] || 0) - (b[k] || 0); if (Math.abs(v) > 0.005) o[k] = v; } return o; };
  const rec = dif(A.rec, B.rec), custo = dif(A.custo, B.custo), vend = dif(A.vend, B.vend);
  const { dv } = base._dv; const daVenda = t => DESP_VENDA.includes(t.destino) && dv[t.quadra] && secaoDe(t.destino) !== 'aporte' && !t.fin;
  // demais itens pelo vencimento: receitas sem lote (ex.: venda de quadra inteira), despesas, impostos, financeiras
  const noPer = t => t.venc >= ini && t.venc <= fim;
  const ts = L.filter(t => noPer(t) && secaoDe(t.destino) !== 'aporte' && !daVenda(t));
  for (const t of ts) if (secaoDe(t.destino) === 'rec' && !FIN_REC.includes(t.destino) && (!t.quadra || t.quadra === '-' || !E().quadras.includes(t.quadra) || t.fin)) rec[t.destino] = (rec[t.destino] || 0) + t.valor;
  const by = f => { const o = {}; for (const t of ts.filter(f)) o[t.destino] = (o[t.destino] || 0) + t.valor; return o; };
  const imp = sum(ts.filter(t => t.destino === 'Impostos'), t => t.valor) + (vend['Impostos'] || 0);
  const desp = by(t => secaoDe(t.destino) === 'nonop' && t.destino !== 'Impostos'); for (const d in vend) if (d !== 'Impostos') desp[d] = (desp[d] || 0) + vend[d];
  const fin = sum(ts.filter(t => FIN_REC.includes(t.destino)), t => t.valor);
  const juros = sum(L.filter(t => pago(t) && t.venc >= ini && t.venc <= fim), t => t.juros || 0);
  const S_ = o => sum(Object.values(o), v => v);
  const rb = S_(rec), rl = rb + imp, cu = S_(custo), lb = rl + cu, de = S_(desp), ro = lb + de, res = ro + fin + juros;
  return { rec, imp, custo, desp, fin, juros, rb, rl, cu, lb, de, ro, res };
}
