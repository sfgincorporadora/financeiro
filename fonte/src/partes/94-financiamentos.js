/* =========================================================
   Financiamentos (Apoio à Produção e outras dívidas):
   contrato, liberações, juros/encargos, amortizações, saldo devedor
   e amortização sugerida a cada repasse de lote financiado.
   Os movimentos são lançamentos comuns marcados com fin (contrato) e finTipo.
   ========================================================= */
usarConfig('financiamentos');
const fins = () => ((S.financiamentos && S.financiamentos.lista) || []);
const finPorId = id => fins().find(c => c.id === id);
function salvarFins(lista) { S.financiamentos = { ...(S.financiamentos || {}), lista }; gravar(doc(db, 'config', 'financiamentos'), { lista }).catch(onErroEscrita); refreshView(); }
const FIN_TIPO = { lib: 'Liberação', juros: 'Juros / encargos', amort: 'Amortização' };
function resumoFin(c) {
  const ts = S.list.filter(t => t.fin === c.id); const v = t => t.valor;
  const lib = sum(ts.filter(t => t.finTipo === 'lib' && pago(t)), v), libPrev = sum(ts.filter(t => t.finTipo === 'lib' && !pago(t)), v);
  const amort = -sum(ts.filter(t => t.finTipo === 'amort' && pago(t)), v), amortPrev = -sum(ts.filter(t => t.finTipo === 'amort' && !pago(t)), v);
  const juros = -sum(ts.filter(t => pago(t)), t => t.finTipo === 'juros' ? t.valor : (t.juros || 0)), jurosPrev = -sum(ts.filter(t => !pago(t) && t.finTipo === 'juros'), v);
  const saldo = lib - amort;
  // repasses de lotes do contrato sem amortização lançada
  const lotesAmort = new Set(ts.filter(t => t.finTipo === 'amort' && t.finLote).map(t => t.finLote));
  const pend = (c.quadras || []).flatMap(q => lotesDe(q).map(l => infoVendaLote(q, l))).filter(x => x.etapa === 'recebido' && !lotesAmort.has(x.k) && (!c.inicio || (x.ult || '') >= c.inicio));
  return { ts, lib, libPrev, amort, amortPrev, juros, jurosPrev, saldo, aLiberar: Math.max(0, (+c.valor || 0) - lib - libPrev), pend };
}
VIEWS.financiamentos = {
  titulo: 'Financiamentos',
  refresh(el) {
    const cs = fins(); const rs = cs.map(c => ({ c, r: resumoFin(c) })); const at = rs.filter(x => x.c.status !== 'quitado');
    el.innerHTML = `<div class="grid kpis k4">${kpi('Saldo devedor', brl(sum(at, x => x.r.saldo)), `${at.length} contrato(s) ativo(s)`, sum(at, x => x.r.saldo) > 0.005 ? 'neg' : '')}${kpi('Liberado', brl(sum(rs, x => x.r.lib)), 'recebido dos bancos', '')}${kpi('A liberar', brl(sum(at, x => x.r.libPrev + x.r.aLiberar)), 'do valor contratado', '')}${kpi('Juros e encargos pagos', brl(sum(rs, x => x.r.juros)), 'custo da dívida', 'neg')}</div>
      <div class="card"><h2>Contratos de financiamento <button class="btn sm primary rw" data-act="finEd">Novo financiamento</button></h2>
      ${cs.length ? `<div class="tbl-wrap"><table class="tbl"><thead><tr><th class="l">Contrato</th><th class="l">Obra / quadras</th><th>Contratado</th><th>Liberado</th><th>Amortizado</th><th>Saldo devedor</th><th>Juros pagos</th><th class="l">Pendências</th></tr></thead><tbody>
      ${rs.map(({ c, r }) => `<tr class="click" data-act="finAbrir" data-id="${esc(c.id)}"><td class="l"><b>${esc(c.nome)}</b><br><span class="muted small">${esc(c.banco || '')}${c.taxa ? ' · ' + esc(c.taxa) : ''}${c.status === 'quitado' ? ' · quitado' : ''}</span></td><td class="l">${esc(c.obra || '')}${(c.quadras || []).length ? '<br><span class="muted small">' + esc(c.quadras.join(', ')) + '</span>' : ''}</td><td>${brl(+c.valor || 0)}</td><td class="pos">${brl(r.lib)}</td><td>${brl(r.amort)}</td><td class="${r.saldo > 0.005 ? 'neg' : ''}"><b>${brl(r.saldo)}</b></td><td class="neg">${brl(r.juros)}</td><td class="l">${r.pend.length ? `<span class="chip warn">${r.pend.length} repasse(s) sem amortização</span>` : '<span class="muted">—</span>'}</td></tr>`).join('')}</tbody></table></div>`
      : '<p class="muted">Nenhum financiamento cadastrado. Cadastre o contrato (ex.: Apoio à Produção da quadra K3) e registre as liberações, os juros e as amortizações aqui: o app lança tudo no caixa, no BP (passivo) e no custo da obra.</p>'}</div>
      <p class="muted small">Liberações e amortizações entram como "Empréstimo / financiamento de terceiros": no BP ficam no passivo e não são receita nem despesa. Juros e encargos entram como despesa ("Juros e encargos") na obra e na DRE.</p>`;
  }
};
ACOES.finEd = d => {
  const c = finPorId(d.id) || { quadras: [], tipo: 'apoio' };
  const obras = E().obras; const ob = c.obra || obras[0];
  const f = folha(d.id ? 'Editar financiamento' : 'Novo financiamento', `<form class="form" id="fx">
    <div class="g2"><label>Nome<input name="nome" required value="${esc(c.nome || '')}" placeholder="Ex.: Apoio à Produção K3"></label><label>Banco / credor<input name="banco" value="${esc(c.banco || '')}" placeholder="Ex.: Caixa Econômica"></label></div>
    <div class="g2"><label>Tipo<select name="tipo">${opts([['apoio', 'Apoio à Produção / plano empresário'], ['capital', 'Capital de giro'], ['outro', 'Outro empréstimo']], c.tipo || 'apoio')}</select></label><label>Taxa / indexador<input name="taxa" value="${esc(c.taxa || '')}" placeholder="Ex.: TR + 8% a.a."></label></div>
    <div class="g3"><label>Valor contratado (R$)<input name="valor" inputmode="decimal" required value="${c.valor ? fNum2.format(c.valor) : ''}"></label><label>Data do contrato<input type="date" name="inicio" value="${esc(c.inicio || hoje())}"></label><label>Prazo (meses)<input name="prazo" inputmode="numeric" value="${esc(String(c.prazo || ''))}"></label></div>
    ${variasObras() ? `<label>Obra<select name="obra">${opts(obras, ob)}</select></label>` : ''}
    <div class="lbl">Quadras financiadas</div><div class="chips" id="fq">${quadrasDaObra(variasObras() ? ob : '').map(q => `<label class="pick"><input type="checkbox" name="qs" value="${esc(q)}" ${(c.quadras || []).includes(q) ? 'checked' : ''}><span>${esc(q)}</span></label>`).join('')}</div>
    <div class="g2">${variasEmpresas() ? `<label>CNPJ que contratou<select name="emp">${opts(empOpts(), c.emp || (ob ? empDaObra(ob) : matrizId()))}</select></label>` : ''}<label>Conta onde o banco credita<select name="conta">${opts(E().contas.map(x => x.nome), c.conta || E().contas[0]?.nome)}</select></label></div>
    <label>Amortização por unidade no repasse (R$) <span class="muted">(opcional: o app sugere esse valor a cada repasse de lote financiado)</span><input name="amortUn" inputmode="decimal" value="${c.amortUn ? fNum2.format(c.amortUn) : ''}"></label>
    <label>Observação<input name="obs" value="${esc(c.obs || '')}"></label>
    <div class="acts">${d.id ? `<button type="button" class="btn ghost" id="qt">${c.status === 'quitado' ? 'Reabrir' : 'Marcar como quitado'}</button><button type="button" class="btn danger ghost" id="rm">Remover</button>` : ''}<button type="button" class="btn" data-close>Cancelar</button><button class="btn primary">Salvar</button></div></form>`, { largo: true });
  const F = f.el.querySelector('#fx');
  if (F.obra) F.obra.onchange = () => { $('#fq', F).innerHTML = quadrasDaObra(F.obra.value).map(q => `<label class="pick"><input type="checkbox" name="qs" value="${esc(q)}"><span>${esc(q)}</span></label>`).join(''); if (F.emp) F.emp.value = empDaObra(F.obra.value); };
  F.onsubmit = e => { e.preventDefault(); const v = parseBR(F.valor.value); if (!(v > 0)) { toast('Informe o valor contratado.', 'err'); return; }
    const n = { ...c, id: c.id || 'fn' + uid(), nome: F.nome.value.trim(), banco: F.banco.value.trim(), tipo: F.tipo.value, taxa: F.taxa.value.trim(), valor: r2(v), inicio: F.inicio.value, prazo: parseInt(F.prazo.value) || 0, obra: F.obra ? F.obra.value : E().obras[0], quadras: $$('[name=qs]:checked', F).map(x => x.value), emp: F.emp ? F.emp.value : matrizId(), conta: F.conta.value, amortUn: parseBR(F.amortUn.value) || 0, obs: F.obs.value.trim(), status: c.status || 'ativo' };
    salvarFins(d.id ? fins().map(x => x.id === c.id ? n : x) : [...fins(), n]); f.fechar(); toast('Financiamento salvo'); ACOES.finAbrir({ id: n.id }); };
  const rm = f.el.querySelector('#rm'); if (rm) rm.onclick = async () => { if (!(await confirmar('Remover o contrato? Os lançamentos já feitos continuam no extrato.', 'Remover'))) return; salvarFins(fins().filter(x => x.id !== c.id)); f.fechar(); };
  const qt = f.el.querySelector('#qt'); if (qt) qt.onclick = () => { salvarFins(fins().map(x => x.id === c.id ? { ...x, status: c.status === 'quitado' ? 'ativo' : 'quitado' } : x)); f.fechar(); };
};
ACOES.finAbrir = d => {
  const c = finPorId(d.id); if (!c) return; const r = resumoFin(c);
  // saldo devedor mês a mês
  const meses = {}; for (const t of r.ts) { const dt = (pago(t) ? (t.pag || t.venc) : t.venc) || ''; const m = dt.slice(0, 7); if (!m) continue; const o = meses[m] = meses[m] || { lib: 0, amort: 0, juros: 0, prev: !pago(t) }; if (pago(t)) o.prev = false;
    if (t.finTipo === 'lib') o.lib += t.valor; else if (t.finTipo === 'amort') { o.amort += -t.valor; o.juros += -(t.juros || 0); } else if (t.finTipo === 'juros') o.juros += -t.valor; }
  let ac = 0; const linhas = Object.keys(meses).sort().map(m => { const o = meses[m]; ac += o.lib - o.amort; return { m, ...o, saldo: ac, futuro: m > hoje().slice(0, 7) }; });
  const f = folha(c.nome, `<div class="grid kpis k4">${kpi('Saldo devedor', brl(r.saldo), c.banco || '', r.saldo > 0.005 ? 'neg' : '')}${kpi('Liberado', brl(r.lib), `de ${brl(+c.valor || 0)} contratados`, 'pos')}${kpi('A liberar', brl(r.libPrev + r.aLiberar), r.libPrev ? `${brl(r.libPrev)} com data prevista` : 'sem cronograma lançado', '')}${kpi('Juros e encargos', brl(r.juros), r.jurosPrev ? `+ ${brl(r.jurosPrev)} previstos` : 'pagos até hoje', 'neg')}</div>
    <p class="muted small">${esc([c.taxa, c.prazo ? c.prazo + ' meses' : '', c.inicio ? 'contrato em ' + dBR(c.inicio) : '', (c.quadras || []).length ? 'quadras ' + c.quadras.join(', ') : '', c.amortUn ? 'amortização de ' + brl(c.amortUn) + ' por repasse' : ''].filter(Boolean).join(' · '))}</p>
    <div class="acts left rw"><button class="btn primary" data-fm="lib">Registrar liberação</button><button class="btn" data-fm="juros">Juros / encargos</button><button class="btn" data-fm="amort">Amortização</button><button class="btn" id="fCr">Cronograma de liberações</button><button class="btn ghost" id="fEd">Editar contrato</button></div>
    ${r.pend.length ? `<div class="alert">Repasses de lotes deste financiamento sem amortização lançada: ${r.pend.slice(0, 12).map(x => `<button class="btn sm" data-am="${esc(x.k)}">${esc(x.q)} ${esc(x.l)} · ${dBR(x.ult)}</button>`).join(' ')}</div>` : ''}
    <div class="lbl" style="margin-top:12px">Saldo devedor mês a mês</div>
    <div class="tbl-wrap"><table class="tbl"><thead><tr><th class="l">Mês</th><th>Liberações</th><th>Amortizações</th><th>Juros e encargos</th><th>Saldo devedor</th></tr></thead><tbody>
    ${linhas.map(x => `<tr class="${x.prev ? 'muted' : ''}"><td class="l">${mesNome(x.m)}${x.prev ? ' <span class="chip">previsto</span>' : ''}</td><td class="pos">${x.lib ? brl(x.lib) : '—'}</td><td>${x.amort ? brl(x.amort) : '—'}</td><td class="neg">${x.juros ? brl(x.juros) : '—'}</td><td><b>${brl(x.saldo)}</b></td></tr>`).join('') || '<tr><td colspan="5" class="l muted">Nenhum movimento ainda.</td></tr>'}</tbody></table></div>
    <div class="lbl">Lançamentos do contrato (${r.ts.length})</div><div class="tlist">${r.ts.slice().sort((a, b) => (dataRef(b) || '').localeCompare(dataRef(a) || '')).slice(0, 200).map(t => linhaTitulo(t, { mostrarPago: true })).join('') || '<p class="muted">—</p>'}</div>`, { largo: true });
  $$('[data-fm]', f.el).forEach(b => b.onclick = () => { f.fechar(); formMovFin(c, b.dataset.fm); });
  $$('[data-am]', f.el).forEach(b => b.onclick = () => { f.fechar(); const x = r.pend.find(p => p.k === b.dataset.am); formMovFin(c, 'amort', { lote: x.k, data: x.ult }); });
  f.el.querySelector('#fEd').onclick = () => { f.fechar(); ACOES.finEd({ id: c.id }); };
  f.el.querySelector('#fCr').onclick = () => { f.fechar(); formCronoFin(c); };
  f.el.addEventListener('click', e => { if (e.target.closest('[data-act="abrir"]')) f.fechar(); });
};
/* Lançamento de um movimento do financiamento. */
function formMovFin(c, tipo, pre = {}) {
  const r = resumoFin(c); const qs = c.quadras || [];
  const sugest = tipo === 'lib' ? r.aLiberar : tipo === 'amort' ? (pre.lote ? c.amortUn || 0 : 0) : 0;
  const f = folha(`${FIN_TIPO[tipo]} · ${c.nome}`, `<form class="form" id="fx">
    <div class="g2"><label>Valor (R$)<input name="valor" inputmode="decimal" required value="${sugest ? fNum2.format(sugest) : ''}"></label><label>Data<input type="date" name="data" value="${esc(pre.data || hoje())}" required></label></div>
    ${tipo === 'amort' ? '<label>Juros cobrados junto (R$) <span class="muted">(opcional; entram como despesa)</span><input name="juros" inputmode="decimal" placeholder="0,00"></label>' : ''}
    <label>Quadra<select name="quadra">${opts([...qs.map(q => [q, 'Quadra ' + q]), ['-', 'Geral da obra']], pre.lote ? pre.lote.split('|')[0] : qs[0] || '-')}</select></label>
    ${pre.lote ? `<p class="muted small">Amortização do repasse do lote ${esc(pre.lote.replace('|', ' '))}.</p>` : ''}
    <label class="chk"><input type="checkbox" name="pg" ${tipo === 'lib' && !pre.data ? '' : 'checked'}> Já ${tipo === 'lib' ? 'creditado' : 'pago'} (lança direto no extrato)</label>
    <label>Conta<select name="conta">${opts(E().contas.map(x => x.nome), c.conta)}</select></label>
    <label>Observação<input name="obs" placeholder="${tipo === 'lib' ? 'Ex.: liberação da 2ª medição' : tipo === 'juros' ? 'Ex.: juros de obra de outubro' : ''}"></label>
    <div class="acts"><button type="button" class="btn" data-close>Cancelar</button><button class="btn primary">Lançar</button></div></form>`);
  const F = f.el.querySelector('#fx');
  F.onsubmit = e => { e.preventDefault(); const v = parseBR(F.valor.value); if (!(v > 0)) { toast('Informe o valor.', 'err'); return; }
    if (tipo === 'juros' && typeof garantirDestinoJuros === 'function') garantirDestinoJuros();
    const q = F.quadra.value; const pg = F.pg.checked;
    const t = { id: uid(), item: `${FIN_TIPO[tipo]} ${c.nome}${pre.lote ? ' - ' + pre.lote.replace('|', ' ') : ''}`, valor: (tipo === 'lib' ? 1 : -1) * r2(v), juros: 0, destino: tipo === 'juros' ? (typeof DEST_JUROS !== 'undefined' ? DEST_JUROS : 'Juros e encargos') : 'Aporte', ...(tipo !== 'juros' ? { aporteTipo: 'terceiro' } : {}),
      quadra: q, lote: '-', local: q !== '-' ? obraDe(q) : (c.obra || E().obras[0]), emp: c.emp || matrizId(), forn: c.banco || '', venc: F.data.value, status: pg ? 'pago' : 'aberto', ...(pg ? { pag: F.data.value, conta: F.conta.value } : {}), obs: F.obs.value.trim(), fin: c.id, finTipo: tipo, ...(pre.lote ? { finLote: pre.lote } : {}), criado: Date.now() };
    if (tipo === 'amort') { const j = parseBR(F.juros.value) || 0; if (j > 0) t.juros = -r2(j); }
    salvarTitulos([t]); f.fechar(); toast(`${FIN_TIPO[tipo]} lançada`); ACOES.finAbrir({ id: c.id });
  };
}
/* Cronograma de liberações previstas (contas a receber). */
function formCronoFin(c) {
  const r = resumoFin(c);
  const f = folha('Cronograma de liberações · ' + c.nome, `<form class="form" id="fx"><p class="muted small">Cria as liberações previstas como contas a receber, para o fluxo de caixa contar com elas. Ajuste as datas e valores depois, se precisar.</p>
    <div class="g3"><label>Valor total (R$)<input name="tot" inputmode="decimal" value="${fNum2.format(r.aLiberar)}"></label><label>Parcelas<input name="n" type="number" min="1" max="60" value="${Math.max(1, Math.min(12, c.prazo || 6))}"></label><label>Primeira data<input type="date" name="d" value="${addMonths(hoje(), 1)}"></label></div>
    <label>Quadra<select name="quadra">${opts([...(c.quadras || []).map(q => [q, 'Quadra ' + q]), ['-', 'Geral da obra']], (c.quadras || [])[0] || '-')}</select></label>
    <div class="acts"><button type="button" class="btn" data-close>Cancelar</button><button class="btn primary">Criar cronograma</button></div></form>`);
  const F = f.el.querySelector('#fx');
  F.onsubmit = e => { e.preventDefault(); const tot = parseBR(F.tot.value) || 0, n = Math.max(1, parseInt(F.n.value) || 1); if (!(tot > 0)) return;
    const cents = Math.round(tot * 100), b = Math.floor(cents / n); let rr = cents - b * n; const g = uid(); const q = F.quadra.value;
    const ts = Array.from({ length: n }, (_, i) => ({ id: uid(), item: `Liberação ${c.nome} (${i + 1}/${n})`, valor: (b + (rr-- > 0 ? 1 : 0)) / 100, juros: 0, destino: 'Aporte', aporteTipo: 'terceiro', quadra: q, lote: '-', local: q !== '-' ? obraDe(q) : (c.obra || E().obras[0]), emp: c.emp || matrizId(), forn: c.banco || '', venc: proximoDiaUtil(addMonths(F.d.value, i)), status: 'aberto', grupo: g, obs: '', fin: c.id, finTipo: 'lib', criado: Date.now() }));
    salvarTitulos(ts); f.fechar(); toast(`${n} liberação(ões) prevista(s) criadas`); ACOES.finAbrir({ id: c.id }); };
}
BUSCA_EXTRA.push(q => fins().filter(c => normTxt(c.nome + ' ' + (c.banco || '')).includes(q)).map(c => ({ titulo: 'Financiamento · ' + c.nome, sub: c.banco || '', href: '#/financiamentos' })));
