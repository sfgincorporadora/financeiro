/* =========================================================
   Compras e cotações: pedido → cotações lado a lado → aprovação → contas a pagar → recebimento
   Guardado em compras/{AAAA-MM} (um documento por mês de abertura).
   ========================================================= */
usarBlocos('compras', o => { const l = []; for (const b in o) for (const k in (o[b].itens || {})) { try { const c = JSON.parse(o[b].itens[k]); c._b = b; l.push(c); } catch (e) { /* ignora */ } } S.compras = l; });
const ST_COMPRA = [['cotando', 'Cotando'], ['aprovacao', 'Aguardando aprovação'], ['aprovada', 'Aprovada, aguardando entrega'], ['recebida', 'Recebida'], ['cancelada', 'Cancelada']];
const stCompraNome = k => (ST_COMPRA.find(x => x[0] === k) || ['', k])[1];
const compraPorId = id => (S.compras || []).find(c => c.id === id);
function salvarCompra(c, acao) {
  const b = c._b || (c.data || hoje()).slice(0, 7); const o = { ...c }; delete o._b;
  if (acao) o.hist = [...(o.hist || []), { em: Date.now(), por: S.user?.email || '', a: acao }];
  S.compras = [...(S.compras || []).filter(x => x.id !== c.id), { ...o, _b: b }];
  gravar(doc(db, 'compras', b), { itens: { [o.id]: JSON.stringify(o) } }, { merge: true }).catch(onErroEscrita);
}
const totalCot = q => (+q.valor || 0) + (+q.frete || 0);
/* Preço médio pago antes por unidade, para itens com quantidade registrada. */
function precoHistorico(desc) {
  const k = normTxt(desc).split(' ').filter(w => w.length > 3).slice(0, 2); if (!k.length) return null;
  const ts = S.list.filter(t => t.qtd > 0 && t.valor < 0 && k.every(w => normTxt(t.item).includes(w))).sort((a, b) => (dataRef(b) || '').localeCompare(dataRef(a) || '')).slice(0, 20);
  if (!ts.length) return null; const qt = sum(ts, t => t.qtd); return { unit: -sum(ts, t => t.valor) / qt, unid: ts[0].unid || '', n: ts.length, ult: dataRef(ts[0]), forn: ts[0].forn };
}
let compSt = 'abertas';
VIEWS.compras = {
  titulo: 'Compras e cotações',
  refresh(el) {
    const cs = (S.compras || []).slice().sort((a, b) => (b.data || '').localeCompare(a.data || '') || (b.n || 0) - (a.n || 0));
    const filtro = c => compSt === 'todas' ? true : compSt === 'abertas' ? !['recebida', 'cancelada'].includes(c.status) : c.status === compSt;
    const cont = k => cs.filter(c => c.status === k).length;
    el.innerHTML = `<div class="grid kpis k4">${kpi('Cotando', cont('cotando'), '', '')}${kpi('Aguardando aprovação', cont('aprovacao'), '', cont('aprovacao') ? 'neg' : '')}${kpi('Aguardando entrega', cont('aprovada'), '', '')}${kpi('Comprado no mês', brl(sum(cs.filter(c => ['aprovada', 'recebida'].includes(c.status) && (c.aprovadaEm || '').slice(0, 7) === hoje().slice(0, 7)), c => totalCot(c.cotacoes[c.esc] || {}))), 'aprovado em ' + mesNome(hoje().slice(0, 7)), '')}</div>
      <div class="filtros card"><select id="cSt">${opts([['abertas', 'Em andamento'], ...ST_COMPRA, ['todas', 'Todas']], compSt)}</select><button class="btn primary rw" data-act="cpNova">Novo pedido de compra</button></div>
      <div class="card"><div class="tbl-wrap"><table class="tbl"><thead><tr><th class="l">Nº</th><th class="l">Pedido</th><th class="l">Obra / quadra</th><th>Cotações</th><th class="l">Melhor preço</th><th class="l">Situação</th></tr></thead><tbody>
      ${cs.filter(filtro).map(c => { const qs = c.cotacoes || []; const melhor = qs.length ? qs.reduce((a, b) => totalCot(b) < totalCot(a) ? b : a) : null; const esc_ = c.esc !== undefined && qs[c.esc]; return `<tr class="click" data-act="cpAbrir" data-id="${esc(c.id)}"><td class="l">${c.n || ''}</td><td class="l"><b>${esc(c.desc)}</b><br><span class="muted small">${dBR(c.data)}${c.itens?.length ? ' · ' + c.itens.length + ' item(ns)' : ''}</span></td><td class="l">${esc([c.obra, c.quadra && c.quadra !== '-' ? 'Quadra ' + c.quadra : ''].filter(Boolean).join(' · '))}</td><td>${qs.length}</td><td class="l">${esc_ ? `${esc(esc_.forn)} · ${brl2(totalCot(esc_))}` : melhor ? `${esc(melhor.forn)} · ${brl2(totalCot(melhor))}` : '—'}</td><td class="l"><span class="chip ${c.status === 'recebida' ? 'ok' : c.status === 'aprovacao' ? 'warn' : c.status === 'cancelada' ? '' : 'fx-andamento'}">${stCompraNome(c.status)}</span></td></tr>`; }).join('') || '<tr><td class="l muted" colspan="6">Nenhum pedido aqui.</td></tr>'}</tbody></table></div>
      <p class="muted small">Peça três cotações, compare lado a lado com o preço que você já pagou antes e aprove: a compra aprovada vira conta a pagar com fornecedor, quantidade e parcelas.</p></div>`;
    $('#cSt', el).onchange = e => { compSt = e.target.value; this.refresh(el); };
  }
};
ACOES.cpNova = () => formCompra(null);
function formCompra(c) {
  const itens = (c?.itens || [{ desc: '', qtd: '', unid: '' }]).map(x => ({ ...x }));
  const f = folha(c ? 'Editar pedido' : 'Novo pedido de compra', `<form class="form" id="fx"><label>O que comprar<input name="desc" required value="${esc(c?.desc || '')}" placeholder="Ex.: cimento para fundação K3"></label>
    ${campoObra(c?.obra || '', c?.emp || '')}
    <div class="g2"><label>Quadra<select name="quadra">${opts(quadrasOpts(), c?.quadra || '-')}</select></label><label>Destino<select name="destino">${optsDestino(c?.destino || 'Material')}</select></label></div>
    <div class="g2"><label>Precisa até<input type="date" name="ate" value="${esc(c?.ate || '')}"></label><label>Etapa da obra<select name="etapaObra">${opts(etapasOrc(), c?.etapaObra || '', 'Sem etapa')}</select></label></div>
    <div class="lbl">Itens</div><div id="cpIt"></div><button type="button" class="btn sm" id="cpAdd">+ Item</button>
    <label>Observação<input name="obs" value="${esc(c?.obs || '')}"></label>
    <div class="acts"><button type="button" class="btn" data-close>Cancelar</button><button class="btn primary">Salvar</button></div></form>`, { largo: true });
  const F = f.el.querySelector('#fx'); const box = F.querySelector('#cpIt');
  const des = () => { box.innerHTML = itens.map((x, i) => `<div class="cp-item"><input data-i="${i}" data-k="desc" placeholder="Item" value="${esc(x.desc)}"><input data-i="${i}" data-k="qtd" inputmode="decimal" placeholder="Qtd" value="${esc(x.qtd)}"><input data-i="${i}" data-k="unid" placeholder="Unid." list="dlUnid2" value="${esc(x.unid)}"><button type="button" class="x" data-rm="${i}">×</button></div>`).join('') + '<datalist id="dlUnid2"><option value="sc"><option value="m³"><option value="m²"><option value="un"><option value="kg"><option value="t"><option value="milheiro"><option value="m"></datalist>';
    $$('[data-k]', box).forEach(x => x.oninput = () => { itens[+x.dataset.i][x.dataset.k] = x.value; }); $$('[data-rm]', box).forEach(b => b.onclick = () => { itens.splice(+b.dataset.rm, 1); des(); }); };
  F.querySelector('#cpAdd').onclick = () => { itens.push({ desc: '', qtd: '', unid: '' }); des(); }; des();
  F.onsubmit = e => {
    e.preventDefault(); const n = { ...(c || { id: 'cp' + uid(), n: Math.max(0, ...(S.compras || []).map(x => x.n || 0)) + 1, data: hoje(), status: 'cotando', cotacoes: [] }), desc: F.desc.value.trim(), obra: F.obra ? F.obra.value : E().obras[0], emp: F.emp ? F.emp.value : '', quadra: F.quadra.value, destino: F.destino.value, ate: F.ate.value, etapaObra: F.etapaObra.value, obs: F.obs.value.trim(), itens: itens.filter(x => x.desc.trim()).map(x => ({ desc: x.desc.trim(), qtd: parseBR(String(x.qtd)) || 0, unid: x.unid.trim() })) };
    salvarCompra(n, c ? 'editou' : 'abriu o pedido'); f.fechar(); ACOES.cpAbrir({ id: n.id });
  };
}
ACOES.cpAbrir = d => {
  const c = compraPorId(d.id); if (!c) return; const qs = c.cotacoes || []; const tots = qs.map(totalCot); const min = Math.min(...tots);
  const um = c.itens && c.itens.length === 1 && c.itens[0].qtd > 0 ? c.itens[0] : null; const ph = um ? precoHistorico(um.desc) : null;
  const pode = !['aprovada', 'recebida', 'cancelada'].includes(c.status);
  const f = folha(`Pedido ${c.n} · ${c.desc}`, `<p class="muted small">${dBR(c.data)} · ${esc([c.obra, c.quadra !== '-' ? 'Quadra ' + c.quadra : '', c.destino, c.etapaObra].filter(Boolean).join(' · '))}${c.ate ? ' · precisa até ' + dBR(c.ate) : ''} · <b>${stCompraNome(c.status)}</b></p>
    ${c.itens?.length ? `<ul class="lista-simples">${c.itens.map(x => `<li>${esc(x.desc)}${x.qtd ? ` · ${fNum2.format(x.qtd)} ${esc(x.unid)}` : ''}</li>`).join('')}</ul>` : ''}
    ${ph ? `<p class="small">Histórico: você pagou em média <b>${brl2(ph.unit)} por ${esc(ph.unid)}</b> nas últimas ${ph.n} compras (última em ${dBR(ph.ult)}, ${esc(ph.forn || '')}).</p>` : ''}
    <div class="lbl">Cotações</div>
    ${qs.length ? `<div class="tbl-wrap"><table class="tbl cot"><thead><tr><th class="l">Fornecedor</th><th>Valor</th><th>Frete</th><th>Total</th>${um ? `<th>Por ${esc(um.unid || 'unid.')}</th>` : ''}<th class="l">Pagamento</th><th class="l">Entrega</th><th></th></tr></thead><tbody>
    ${qs.map((q, i) => `<tr class="${tots[i] === min ? 'melhor' : ''} ${c.esc === i ? 'escolhida' : ''}"><td class="l"><b>${esc(q.forn)}</b>${q.obs ? `<br><span class="muted small">${esc(q.obs)}</span>` : ''}</td><td>${brl2(q.valor)}</td><td>${q.frete ? brl2(q.frete) : '—'}</td><td><b>${brl2(tots[i])}</b>${tots[i] === min && qs.length > 1 ? '<br><span class="chip ok">menor</span>' : ''}</td>${um ? `<td>${brl2(tots[i] / um.qtd)}${ph ? `<br><span class="small ${tots[i] / um.qtd > ph.unit * 1.05 ? 'neg' : 'pos'}">${fPct.format(tots[i] / um.qtd / ph.unit - 1)} vs. histórico</span>` : ''}</td>` : ''}<td class="l">${q.parc > 1 ? `${q.parc}x a cada ${q.int || 30} dias` : 'à vista'}, 1º em ${q.dias || 0} dias</td><td class="l">${q.entrega ? q.entrega + ' dias' : '—'}</td><td>${pode ? `<button class="btn sm rw" data-esc="${i}">${c.esc === i ? 'Escolhida' : 'Escolher'}</button> <button class="btn sm ghost rw" data-qrm="${i}">×</button>` : c.esc === i ? '<span class="chip ok">comprada</span>' : ''}</td></tr>`).join('')}</tbody></table></div>` : '<p class="muted">Nenhuma cotação ainda.</p>'}
    <div class="acts left rw">${pode ? `<button class="btn" id="qAdd">+ Cotação</button>` : ''}${c.status === 'aprovacao' && papel() === 'dono' ? '<button class="btn primary" id="qApr">Aprovar e lançar a pagar</button>' : ''}${c.status === 'aprovacao' && papel() !== 'dono' ? '<span class="muted small">Aguardando aprovação do dono.</span>' : ''}${c.status === 'aprovada' ? '<button class="btn primary" id="qRec">Marcar como recebida</button>' : ''}${pode ? '<button class="btn ghost" id="qEd">Editar pedido</button><button class="btn danger ghost" id="qCan">Cancelar pedido</button>' : ''}</div>
    ${(c.titulos || []).length ? `<div class="lbl">Contas a pagar geradas</div><div class="tlist">${c.titulos.map(id => S.byId.get(id)).filter(Boolean).map(t => linhaTitulo(t, { mostrarPago: true })).join('')}</div>` : ''}
    ${(c.hist || []).length ? `<details class="hist"><summary>Histórico do pedido</summary><ol>${c.hist.slice().reverse().map(h => `<li><span class="muted small">${quando(h.em)}</span> <b>${quem(h.por)}</b> ${esc(h.a)}</li>`).join('')}</ol></details>` : ''}`, { largo: true });
  const re = () => { f.fechar(); ACOES.cpAbrir(d); };
  $$('[data-esc]', f.el).forEach(b => b.onclick = () => { salvarCompra({ ...c, esc: +b.dataset.esc, status: 'aprovacao' }, 'escolheu a cotação de ' + qs[+b.dataset.esc].forn); re(); });
  $$('[data-qrm]', f.el).forEach(b => b.onclick = () => { const i = +b.dataset.qrm; const n = { ...c, cotacoes: qs.filter((_, j) => j !== i) }; if (c.esc === i) { delete n.esc; n.status = 'cotando'; } else if (c.esc > i) n.esc = c.esc - 1; salvarCompra(n, 'removeu uma cotação'); re(); });
  const on = (id, fn) => { const b = f.el.querySelector('#' + id); if (b) b.onclick = fn; };
  on('qAdd', () => { f.fechar(); formCotacao(c); });
  on('qEd', () => { f.fechar(); formCompra(c); });
  on('qCan', async () => { if (await confirmar('Cancelar este pedido?', 'Cancelar pedido')) { salvarCompra({ ...c, status: 'cancelada' }, 'cancelou'); re(); } });
  on('qRec', () => { salvarCompra({ ...c, status: 'recebida', recebidaEm: hoje() }, 'marcou como recebida'); re(); });
  on('qApr', () => { f.fechar(); aprovarCompra(c); });
  f.el.addEventListener('click', e => { if (e.target.closest('[data-act="abrir"]')) f.fechar(); });
};
function formCotacao(c) {
  const fs = [...new Set([...S.list.map(t => t.forn), ...Object.keys((S.fornecedores && S.fornecedores.lista) || {})].filter(Boolean))].sort();
  const f = folha('Nova cotação · ' + c.desc, `<form class="form" id="fx"><label>Fornecedor<input name="forn" list="dlFq" required><datalist id="dlFq">${fs.map(x => `<option value="${esc(x)}">`).join('')}</datalist></label>
    <div class="g3"><label>Valor dos itens (R$)<input name="valor" inputmode="decimal" required></label><label>Frete (R$)<input name="frete" inputmode="decimal"></label><label>Entrega em (dias)<input name="entrega" type="number" min="0" inputmode="numeric"></label></div>
    <div class="g3"><label>Parcelas<input name="parc" type="number" min="1" value="1" inputmode="numeric"></label><label>1º vencimento em (dias)<input name="dias" type="number" min="0" value="0" inputmode="numeric"></label><label>Intervalo (dias)<input name="int" type="number" min="1" value="30" inputmode="numeric"></label></div>
    <label>Observação<input name="obs" placeholder="Ex.: validade da proposta, marca"></label>
    <div class="acts"><button type="button" class="btn" data-close>Cancelar</button><button class="btn primary">Adicionar</button></div></form>`);
  f.el.querySelector('#fx').onsubmit = e => { e.preventDefault(); const F = e.target; const v = parseBR(F.valor.value); if (!(v > 0)) return;
    const q = { forn: F.forn.value.trim(), valor: r2(v), frete: r2(parseBR(F.frete.value) || 0), entrega: +F.entrega.value || 0, parc: Math.max(1, +F.parc.value || 1), dias: +F.dias.value || 0, int: +F.int.value || 30, obs: F.obs.value.trim() };
    salvarCompra({ ...c, cotacoes: [...(c.cotacoes || []), q] }, 'adicionou a cotação de ' + q.forn); f.fechar(); ACOES.cpAbrir({ id: c.id }); };
}
function aprovarCompra(c) {
  const q = c.cotacoes[c.esc]; if (!q) return; const tot = totalCot(q); const um = c.itens && c.itens.length === 1 && c.itens[0].qtd > 0 ? c.itens[0] : null;
  const f = folha('Aprovar compra', `<form class="form" id="fx"><p><b>${esc(q.forn)}</b> · ${brl2(tot)} · ${q.parc > 1 ? q.parc + ' parcelas' : 'à vista'}</p>
    ${variasEmpresas() ? `<label>CNPJ da nota / pagamento<select name="emp" required>${opts(empOpts(), c.emp || '', 'Escolha o CNPJ…')}</select></label>` : ''}
    <label class="chk"><input type="checkbox" name="ratear"> Dividir entre os lotes da quadra</label>
    <div class="acts"><button type="button" class="btn" data-close>Cancelar</button><button class="btn primary">Aprovar e lançar</button></div></form>`);
  const F = f.el.querySelector('#fx');
  F.onsubmit = e => {
    e.preventDefault(); if (F.emp && !F.emp.value) { toast('Escolha o CNPJ.', 'err'); return; }
    const L = c.quadra && c.quadra !== '-' && F.ratear.checked ? lotesDe(c.quadra) : ['-']; const out = []; const grupo = uid();
    const cents = Math.round(tot * 100), bp = Math.floor(cents / q.parc); let rp = cents - bp * q.parc;
    for (let i = 0; i < q.parc; i++) {
      const vp = bp + (rp-- > 0 ? 1 : 0); const bl = Math.floor(vp / L.length); let rl = vp - bl * L.length; const venc = proximoDiaUtil(addDays(hoje(), q.dias + i * q.int));
      for (const l of L) { const v = (bl + (rl-- > 0 ? 1 : 0)) / 100; if (!v) continue;
        const t = { id: uid(), item: `${c.desc}${q.parc > 1 ? ` (${i + 1}/${q.parc})` : ''} - ${q.forn}${c.quadra !== '-' ? ' - ' + c.quadra + (l !== '-' ? ' ' + l : '') : ''}`, valor: -v, juros: 0, nf: 0, nfNum: '', local: c.quadra !== '-' ? obraDe(c.quadra) : (c.obra || E().obras[0]), quadra: c.quadra || '-', lote: l, destino: c.destino || 'Material', venc, status: 'aberto', forn: q.forn, obs: 'Pedido de compra ' + c.n, grupo, compra: c.id, criado: Date.now() };
        if (c.etapaObra) t.etapaObra = c.etapaObra; if (F.emp) t.emp = F.emp.value; else if (c.emp) t.emp = c.emp;
        if (um) { t.qtd = Math.round(um.qtd / (q.parc * L.length) * 1000) / 1000; t.unid = um.unid; }
        out.push(t); }
    }
    salvarTitulos(out); salvarCompra({ ...c, status: 'aprovada', aprovadaEm: hoje(), titulos: out.map(t => t.id) }, 'aprovou e lançou ' + brl2(tot) + ' a pagar');
    f.fechar(); toast('Compra aprovada e lançada em Contas a pagar'); ACOES.cpAbrir({ id: c.id });
  };
}
const comprasAprovacao = () => (S.compras || []).filter(c => c.status === 'aprovacao');
BUSCA_EXTRA.push(q => (S.compras || []).filter(c => normTxt(c.desc + ' ' + (c.cotacoes || []).map(x => x.forn).join(' ')).includes(q)).slice(0, 5).map(c => ({ titulo: `Compra ${c.n} · ${c.desc}`, sub: stCompraNome(c.status), href: '#/compras' })));
