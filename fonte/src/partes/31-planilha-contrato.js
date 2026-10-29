/* =========================================================
   Planilha orçamentária do empreiteiro dentro do contrato:
   etapas e itens (unid., quant., unitário, total) e medições por % de etapa,
   no mesmo formato das planilhas de medição (aba Planilha + aba MEDIÇÕES).
   ========================================================= */
function lerPlanilhaEmpreiteiro(buf, nomeArq) {
  const wb = XLSX.read(buf, { type: 'array' });
  const linhas = n => { const ws = wb.Sheets[n]; if (!ws || !ws['!ref']) return []; return XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: true }); };
  const s = x => x === null || x === undefined ? '' : String(x).trim();
  const num = x => typeof x === 'number' && isFinite(x) ? x : (typeof x === 'string' && /^-?[\d.]+,?\d*$/.test(x.trim()) ? parseBR(x) : null);
  const out = { etapas: [], medicoes: [], objeto: '', local: '', forn: '', banco: '', pix: '', arquivo: nomeArq || '' };
  // --- aba com ITEM / DISCRIMINAÇÃO
  for (const nome of wb.SheetNames) {
    const R = linhas(nome); const hi = R.findIndex(r => r && r.some(x => /^ITEM$/i.test(s(x))) && r.some(x => /DISCRIMINA/i.test(s(x))) && r.some(x => /^UNID/i.test(s(x))));
    if (hi < 0) continue;
    const H = R[hi].map(s), H2 = (R[hi + 1] || []).map(s);
    const ci = re => H.findIndex(x => re.test(x));
    const cItem = ci(/^ITEM$/i), cDesc = ci(/DISCRIMINA/i), cUn = ci(/^UNID/i), cQt = ci(/^QUANT/i);
    let cUnit = H2.findIndex(x => /UNIT/i.test(x)); if (cUnit < 0) cUnit = ci(/UNIT|VALOR/i);
    let cTot = H2.findIndex(x => /^TOTAL/i.test(x)); if (cTot < 0) cTot = ci(/TOTAL/i); if (cTot < 0) cTot = cUnit + 1;
    for (let i = 0; i < hi; i++) { const r = R[i] || []; const lab = r.map(s).join(' ');
      if (/OBRA OU SERVI/i.test(lab)) out.objeto = (R[i + 1] || []).map(s).filter(Boolean)[0] || '';
      if (/LOCAL:/i.test(lab)) out.local = (R[i + 1] || []).map(s).filter(Boolean)[0] || ''; }
    let et = null;
    for (const r of R.slice(hi + 1)) {
      if (!r) continue; const it = s(r[cItem]); const desc = s(r[cDesc]) || s(r[cItem + 1]);
      if (/VALOR TOTAL/i.test(r.map(s).join(' '))) break;
      if (/^\d+(\.0+)?$/.test(it) && desc) { et = { n: String(parseInt(it)), nome: desc, valor: num(r[cTot]) || 0, itens: [] }; out.etapas.push(et); continue; }
      if (et && /^\d+\.\d+/.test(it) && desc) et.itens.push({ item: it, desc, un: s(r[cUn]), qt: num(r[cQt]) || 0, unit: num(r[cUnit]) || 0, total: num(r[cTot]) || 0 });
    }
    for (const e of out.etapas) if (!e.valor) e.valor = r2(sum(e.itens, x => x.total));
    if (out.etapas.length) break;
  }
  // --- aba de medições (colunas 1MED, 2MED…)
  for (const nome of wb.SheetNames) {
    const R = linhas(nome); const hi = R.findIndex(r => r && r.some(x => /^\d+\s*MED$/i.test(s(x))));
    if (hi < 0) continue;
    const H = R[hi].map(s); const cols = H.map((x, j) => /^\d+\s*MED$/i.test(x) ? [parseInt(x), j] : null).filter(Boolean);
    const cVal = H.findIndex(x => /VL\.?\s*TOTAL/i.test(x));
    const meds = Object.fromEntries(cols.map(([k]) => [k, {}]));
    let depois = false;
    for (const r of R.slice(hi + 1)) {
      if (!r) continue; const a = s(r[0]);
      const m = a.match(/^(\d+)\s*[-–.]\s*(.+)$/);
      if (m && !depois) {
        const n = m[1]; if (!out.etapas.find(e => e.n === n)) out.etapas.push({ n, nome: m[2], valor: num(r[cVal]) || 0, itens: [] });
        for (const [k, j] of cols) { const p = num(r[j]); if (p) meds[k][n] = p > 1.0001 ? p / 100 : p; }
        continue;
      }
      if (/^TOTAL ACUMULADO/i.test(a)) { depois = true; continue; }
      if (depois && a) {
        if (/AG\.?\s*\d|C\/C|AGÊNCIA|AGENCIA/i.test(a)) out.banco = a;
        else if (/PIX/i.test(a)) out.pix = a.replace(/^.*PIX\s*/i, '').trim();
        else if (!out.forn && /[A-Za-z]{3}/.test(a) && !/^TOTAL/i.test(a)) out.forn = a;
      }
    }
    const total = sum(out.etapas, e => e.valor);
    for (const [k] of cols) { const et = meds[k]; if (!Object.keys(et).length) continue;
      out.medicoes.push({ n: k, etapas: et, valor: r2(sum(out.etapas, e => e.valor * (et[e.n] || 0))) }); }
    let ac = 0; out.medicoes.forEach(m => { ac += m.valor; m.fisico = total ? r2(ac / total * 100) : 0; });
    break;
  }
  out.etapas.sort((a, b) => +a.n - +b.n);
  out.total = r2(sum(out.etapas, e => e.valor));
  if (!out.etapas.length) throw new Error('Não encontrei a planilha orçamentária (colunas ITEM e DISCRIMINAÇÃO) nem a aba de medições (1MED, 2MED…).');
  return out;
}
function escolherArquivo(accept) { return new Promise(res => { const i = document.createElement('input'); i.type = 'file'; i.accept = accept; i.onchange = () => res(i.files[0] || null); i.click(); }); }
async function importarPlanilhaEmpreiteiro(c) {
  const arq = await escolherArquivo('.xlsx,.xls,.xlsm'); if (!arq) return;
  await carregarLib('xlsx'); let p;
  try { p = lerPlanilhaEmpreiteiro(await arq.arrayBuffer(), arq.name); } catch (e) { toast(e.message || 'Não consegui ler a planilha.', 'err'); return; }
  const medTot = sum(p.medicoes, m => m.valor);
  const ok = await confirmar(`<b>${esc(arq.name)}</b><br>${p.etapas.length} etapas, ${sum(p.etapas, e => e.itens.length)} itens, total ${brl2(p.total)}.<br>${p.medicoes.length} medição(ões) somando ${brl2(medTot)} (${p.total ? fPct.format(medTot / p.total) : '—'}).${p.forn ? '<br>Empreiteiro: ' + esc(p.forn) : ''}<p class="muted small">As medições importadas não criam contas a pagar: os pagamentos já lançados podem ser vinculados ao contrato.</p>`, c ? 'Atualizar contrato' : 'Criar contrato', false);
  if (!ok) return;
  const planilha = { etapas: p.etapas, arquivo: p.arquivo, importado: hoje() };
  const medsImp = p.medicoes.map(m => ({ id: 'md' + uid(), n: m.n, data: '', ref: 'Importada da planilha', valor: m.valor, ret: 0, fisico: m.fisico, venc: '', quadra: '-', importada: true, etapas: m.etapas }));
  if (c) {
    const meds = (c.medicoes || []).map(m => { const x = p.medicoes.find(y => y.n === m.n); return x ? { ...m, etapas: x.etapas, fisico: x.fisico } : m; });
    for (const m of medsImp) if (!meds.some(x => x.n === m.n)) meds.push(m);
    salvarContrato({ ...c, planilha, medicoes: meds, banco: c.banco || p.banco, pix: c.pix || p.pix, valor: c.valor || p.total }); toast('Planilha importada no contrato'); ACOES.ctAbrir({ id: c.id }); return;
  }
  formContrato(null, { forn: p.forn.replace(/\s+(LTDA|ME|EIRELI|S\/?A)\.?$/i, m => m), objeto: p.objeto || 'Construção conforme planilha orçamentária', valor: p.total, retPct: 0, planilha, medicoes: medsImp, banco: p.banco, pix: p.pix, destino: 'Mão de Obra' });
}
ACOES.ctImport = () => importarPlanilhaEmpreiteiro(null);
const acumEtapa = (c, n, ateN = Infinity) => sum((c.medicoes || []).filter(m => m.n <= ateN && m.etapas), m => +m.etapas[n] || 0);
function quadroPlanilha(c) {
  const ets = c.planilha.etapas; const tot = sum(ets, e => e.valor) || 1;
  const meds = (c.medicoes || []).filter(m => m.etapas).sort((a, b) => a.n - b.n);
  const medido = e => e.valor * Math.min(1, acumEtapa(c, e.n));
  const pc = v => v ? (v * 100).toLocaleString('pt-BR', { maximumFractionDigits: 1 }) + '%' : '—';
  return `<div class="lbl" style="margin-top:12px">Planilha orçamentária e medições ${c.planilha.arquivo ? `<span class="muted small">· ${esc(c.planilha.arquivo)}</span>` : ''}</div>
  <div class="tbl-wrap"><table class="tbl plan-med"><thead><tr><th class="l">Etapa</th><th>Valor</th><th>Inc.</th>${meds.map(m => `<th>${m.n}ª med.</th>`).join('')}<th>Acum.</th><th>Medido</th><th>Saldo</th></tr></thead><tbody>
  ${ets.map(e => { const a = acumEtapa(c, e.n); return `<tr><td class="l"><details><summary>${esc(e.n)} · ${esc(e.nome)}</summary>${e.itens.length ? `<table class="itens">${e.itens.map(x => `<tr><td class="l">${esc(x.item)}</td><td class="l wrap">${esc(x.desc)}</td><td>${esc(x.un)}</td><td>${(x.qt || 0).toLocaleString('pt-BR')}</td><td>${brl2(x.unit)}</td><td>${brl2(x.total)}</td></tr>`).join('')}</table>` : '<span class="muted small">sem itens</span>'}</details></td><td>${brl(e.valor)}</td><td>${pc(e.valor / tot)}</td>${meds.map(m => `<td>${pc(+m.etapas[e.n] || 0)}</td>`).join('')}<td class="${a >= .999 ? 'pos' : ''}"><b>${pc(a)}</b></td><td>${brl(medido(e))}</td><td>${brl(e.valor - medido(e))}</td></tr>`; }).join('')}
  <tr class="tot"><td class="l">Total</td><td>${brl(tot)}</td><td>100%</td>${meds.map(m => `<td>${brl(sum(ets, e => e.valor * (+m.etapas[e.n] || 0)))}</td>`).join('')}<td><b>${pc(sum(ets, medido) / tot)}</b></td><td>${brl(sum(ets, medido))}</td><td>${brl(tot - sum(ets, medido))}</td></tr></tbody></table></div>
  <p class="muted small">Cada coluna é o percentual de cada etapa medido naquela medição; toque na etapa para ver os itens (unidade, quantidade, unitário e total).</p>`;
}
function camposEtapasMedicao(c) {
  return `<div class="lbl">Percentual medido por etapa nesta medição</div><div class="tbl-wrap"><table class="tbl med-et"><thead><tr><th class="l">Etapa</th><th>Valor</th><th>Já medido</th><th>Nesta (%)</th><th>R$</th></tr></thead><tbody>
  ${c.planilha.etapas.map(e => { const a = acumEtapa(c, e.n); return `<tr><td class="l">${esc(e.n)} · ${esc(e.nome)}</td><td>${brl(e.valor)}</td><td>${(a * 100).toLocaleString('pt-BR', { maximumFractionDigits: 1 })}%</td><td><input name="et_${esc(e.n)}" inputmode="decimal" placeholder="0" ${a >= .999 ? 'disabled' : ''} style="width:70px"></td><td id="etv_${esc(e.n)}">—</td></tr>`; }).join('')}
  </tbody></table></div><p class="muted small">Preencha o % de cada etapa executado nesta medição: o valor e o físico acumulado são calculados sozinhos.</p>`;
}
function etapasDoForm(F, c) { const o = {}; for (const e of c.planilha.etapas) { const i = F['et_' + e.n]; const p = i ? parseBR(i.value) || 0 : 0; if (p) o[e.n] = p / 100; } return o; }
function calcEtapasMedicao(F, c) {
  if (!document.activeElement || !/^et_/.test(document.activeElement.name || '')) return;
  const et = etapasDoForm(F, c); let v = 0;
  for (const e of c.planilha.etapas) { const x = e.valor * (et[e.n] || 0); v += x; const td = F.querySelector('#etv_' + CSS.escape(e.n)); if (td) td.textContent = x ? brl2(x) : '—'; }
  const tot = sum(c.planilha.etapas, e => e.valor) || 1; const ant = sum(c.planilha.etapas, e => e.valor * acumEtapa(c, e.n));
  F.valor.value = fNum2.format(r2(v)); F.fisico.value = String(r2((ant + v) / tot * 100)).replace('.', ',');
}
