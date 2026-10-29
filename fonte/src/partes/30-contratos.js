/* =========================================================
   Contratos e medições de empreiteiros
   Cada contrato é um documento em contratos/{id}; as medições ficam dentro dele.
   ========================================================= */
usarBlocos('contratos', o => { S.contratos = Object.values(o).filter(x => x && x.id); S.loaded.contratos = true; });
const contratos = () => (S.contratos || []).slice().sort((a, b) => (a.status === 'encerrado') - (b.status === 'encerrado') || (a.forn || '').localeCompare(b.forn || ''));
const contratoPorId = id => (S.contratos || []).find(c => c.id === id);
function salvarContrato(c) {
  S.contratos = [...(S.contratos || []).filter(x => x.id !== c.id), c];
  gravar(doc(db, 'contratos', c.id), JSON.parse(JSON.stringify(c))).catch(onErroEscrita);
}
function resumoContrato(c) {
  const ts = S.list.filter(t => t.contrato === c.id);
  const contratado = (+c.valor || 0) + sum(c.aditivos || [], a => +a.valor || 0);
  const medido = sum(c.medicoes || [], m => +m.valor || 0);
  const retido = sum(c.medicoes || [], m => +m.ret || 0) - sum(c.liberacoes || [], l => +l.valor || 0);
  const pg = -sum(ts.filter(pago), t => t.valor + Math.min(0, t.juros || 0));
  const aPagar = -sum(ts.filter(t => !pago(t)), t => t.valor);
  const avulso = -sum(ts.filter(t => !t.medicao && !t.liberacao), t => t.valor);
  const fisico = (c.medicoes || []).map(m => +m.fisico || 0).filter(Boolean).pop() || null;
  return { ts, contratado, medido, retido, pago: pg || 0, aPagar: aPagar || 0, avulso, saldo: contratado - medido - avulso, pct: contratado ? (medido + avulso) / contratado : null, fisico };
}
function titulosMedicao(c, m, venc, valorLiq, quadra, ratear, emp, extra) {
  const L = quadra && quadra !== '-' && ratear ? lotesDe(quadra) : ['-'];
  const cents = Math.round(valorLiq * 100), b = Math.floor(cents / L.length); let r = cents - b * L.length; const grupo = uid();
  return L.map(l => {
    const v = (b + (r-- > 0 ? 1 : 0)) / 100;
    const t = { id: uid(), item: `${extra.titulo} - ${c.forn}${quadra && quadra !== '-' ? ` - ${quadra}${l !== '-' ? ' ' + l : ''}` : ''}`, valor: -v, juros: 0, nf: 0, nfNum: '', local: quadra && quadra !== '-' ? obraDe(quadra) : (c.obra || E().obras[0]), quadra: quadra || '-', lote: l, destino: c.destino || 'Mão de Obra', venc, status: 'aberto', forn: c.forn, obs: extra.obs || '', grupo, contrato: c.id, criado: Date.now() };
    if (extra.medicao) t.medicao = extra.medicao; if (extra.liberacao) t.liberacao = extra.liberacao; if (emp) t.emp = emp;
    return t;
  }).filter(t => t.valor);
}
VIEWS.contratos = {
  titulo: 'Contratos',
  refresh(el) {
    const cs = contratos(); const rs = cs.map(c => ({ c, r: resumoContrato(c) })); const at = rs.filter(x => x.c.status !== 'encerrado');
    const T = k => sum(at, x => x.r[k]);
    el.innerHTML = `<div class="grid kpis k4">${kpi('Contratado (ativos)', brl(T('contratado')), at.length + ' contrato(s)', '')}${kpi('Medido', brl(T('medido')), T('contratado') ? fPct.format(T('medido') / T('contratado')) + ' do contratado' : '', '')}${kpi('Retenção acumulada', brl(T('retido')), 'a liberar no fim', '')}${kpi('Saldo a medir', brl(T('saldo')), '', '')}</div>
      <div class="card"><h2>Contratos de empreiteiros <button class="btn sm primary rw" data-act="ctNovo">Novo contrato</button></h2>
      <div class="tbl-wrap"><table class="tbl"><thead><tr><th class="l">Empreiteiro</th><th class="l">Objeto</th><th>Contratado</th><th class="l">Andamento</th><th>Pago</th><th>A pagar</th><th>Retido</th><th>Saldo</th></tr></thead><tbody>
      ${rs.map(({ c, r }) => `<tr class="click" data-act="ctAbrir" data-id="${esc(c.id)}"><td class="l"><b>${esc(c.forn)}</b>${c.status === 'encerrado' ? ' <span class="chip">encerrado</span>' : ''}<br><span class="muted small">${esc((c.quadras || []).join(', ') || c.obra || '')}</span></td><td class="l wrap">${esc(c.objeto || '')}</td><td>${brl(r.contratado)}</td><td class="l">${pctBar(r.pct)}${r.fisico !== null ? `<small class="muted">físico ${fPct.format(r.fisico / 100)}</small>` : ''}</td><td class="neg">${brl(-r.pago)}</td><td>${brl(-r.aPagar)}</td><td>${brl(r.retido)}</td><td class="${r.saldo < 0 ? 'neg' : ''}">${brl(r.saldo)}</td></tr>`).join('') || '<tr><td class="l muted" colspan="8">Nenhum contrato. Cadastre o contrato de cada empreiteiro para controlar medições, retenção e saldo.</td></tr>'}</tbody></table></div>
      <p class="muted small">Cada medição gera a conta a pagar do valor medido menos a retenção. A retenção acumulada vira conta a pagar quando você a libera. Pagamentos antigos do empreiteiro podem ser vinculados ao contrato.</p></div>`;
  }
};
ACOES.ctNovo = () => formContrato(null);
function formContrato(c) {
  const fs = [...new Set([...S.list.filter(t => t.valor < 0).map(t => t.forn), ...Object.keys((S.fornecedores && S.fornecedores.lista) || {})].filter(Boolean))].sort();
  const f = folha(c ? 'Editar contrato' : 'Novo contrato', `<form class="form" id="fx">
    <div class="g2"><label>Empreiteiro<input name="forn" list="dlFC" required value="${esc(c?.forn || '')}"><datalist id="dlFC">${fs.map(x => `<option value="${esc(x)}">`).join('')}</datalist></label><label>Objeto<input name="objeto" required value="${esc(c?.objeto || '')}" placeholder="Ex.: mão de obra completa das casas"></label></div>
    <div class="g2"><label>Valor contratado (R$)<input name="valor" inputmode="decimal" required value="${c ? fNum2.format(c.valor || 0) : ''}"></label><label>Retenção técnica (%)<input name="ret" inputmode="decimal" value="${c ? String(c.retPct || 0).replace('.', ',') : '5'}"></label></div>
    <div class="g2"><label>Início<input type="date" name="inicio" value="${esc(c?.inicio || hoje())}"></label><label>Término previsto<input type="date" name="fim" value="${esc(c?.fim || '')}"></label></div>
    <label>Destino dos pagamentos<select name="destino">${optsDestino(c?.destino || 'Mão de Obra')}</select></label>
    ${campoObra(c?.obra || '', c?.emp || '')}
    <div class="lbl">Quadras atendidas</div><div class="chips">${E().quadras.map(q => `<label class="chip"><input type="checkbox" name="qs" value="${esc(q)}" ${(c?.quadras || []).includes(q) ? 'checked' : ''}> ${esc(q)}</label>`).join('')}</div>
    <label>Observação<input name="obs" value="${esc(c?.obs || '')}"></label>
    <div class="acts">${c ? `<button type="button" class="btn ghost" id="enc">${c.status === 'encerrado' ? 'Reabrir' : 'Encerrar contrato'}</button>` : ''}<button type="button" class="btn" data-close>Cancelar</button><button class="btn primary">Salvar</button></div></form>`, { largo: true });
  const F = f.el.querySelector('#fx');
  if (F.emp && !c) F.emp.querySelector('option').textContent = 'CNPJ que contrata…';
  F.onsubmit = e => {
    e.preventDefault(); const v = parseBR(F.valor.value); if (!(v > 0)) { toast('Informe o valor do contrato.', 'err'); return; }
    const n = { ...(c || { id: 'ct' + uid(), medicoes: [], aditivos: [], liberacoes: [], status: 'ativo', criado: Date.now(), por: S.user?.email || '' }), forn: F.forn.value.trim(), objeto: F.objeto.value.trim(), valor: r2(v), retPct: parseBR(F.ret.value) || 0, inicio: F.inicio.value, fim: F.fim.value, destino: F.destino.value, obra: F.obra ? F.obra.value : E().obras[0], emp: F.emp ? F.emp.value : '', quadras: $$('[name=qs]:checked', F).map(x => x.value), obs: F.obs.value.trim() };
    salvarContrato(n); f.fechar(); toast('Contrato salvo'); ACOES.ctAbrir({ id: n.id });
  };
  const enc = F.querySelector('#enc'); if (enc) enc.onclick = () => { salvarContrato({ ...c, status: c.status === 'encerrado' ? 'ativo' : 'encerrado' }); f.fechar(); refreshView(); };
}
ACOES.ctAbrir = d => {
  const c = contratoPorId(d.id); if (!c) return; const r = resumoContrato(c);
  const meds = (c.medicoes || []).slice().sort((a, b) => a.n - b.n);
  const f = folha(`${c.forn} · ${c.objeto || 'contrato'}`, `<div class="grid kpis k4">${kpi('Contratado', brl(r.contratado), (c.aditivos || []).length ? (c.aditivos.length + ' aditivo(s)') : 'sem aditivos', '')}${kpi('Medido', brl(r.medido), r.pct !== null ? fPct.format(r.pct) + (r.fisico !== null ? ` · físico ${fPct.format(r.fisico / 100)}` : '') : '', r.fisico !== null && r.pct * 100 > r.fisico + 10 ? 'neg' : '')}${kpi('Retido', brl(r.retido), c.retPct ? c.retPct + '% por medição' : 'sem retenção', '')}${kpi('Saldo a medir', brl(r.saldo), r.avulso ? 'inclui ' + brl(r.avulso) + ' de pagamentos vinculados' : '', r.saldo < 0 ? 'neg' : '')}</div>
    <div class="acts left rw"><button class="btn primary" id="cM">Nova medição</button><button class="btn" id="cA">Aditivo</button>${r.retido > 0.5 ? '<button class="btn" id="cR">Liberar retenção</button>' : ''}<button class="btn" id="cV">Vincular pagamentos</button><button class="btn ghost" id="cE">Editar contrato</button></div>
    <div class="lbl" style="margin-top:12px">Medições</div>
    <div class="tbl-wrap"><table class="tbl"><thead><tr><th class="l">Nº</th><th class="l">Data</th><th class="l">Referência</th><th>Bruto</th><th>Retenção</th><th>Líquido</th><th>Físico</th><th class="l">Situação</th><th></th></tr></thead><tbody>
    ${meds.map(m => { const ts = S.list.filter(t => t.medicao === m.id); const pg = ts.length && ts.every(pago); return `<tr><td class="l">${m.n}</td><td class="l">${dBR(m.data)}</td><td class="l wrap">${esc(m.ref || '')}${m.quadra && m.quadra !== '-' ? ` <span class="muted small">· ${esc(m.quadra)}</span>` : ''}</td><td>${brl2(m.valor)}</td><td>${brl2(m.ret || 0)}</td><td>${brl2(m.valor - (m.ret || 0))}</td><td>${m.fisico ? fPct.format(m.fisico / 100) : '—'}</td><td class="l"><span class="chip ${pg ? 'ok' : ''}">${pg ? 'paga' : ts.some(pago) ? 'paga em parte' : 'a pagar, vence ' + dBR(m.venc)}</span></td><td>${ts.some(pago) ? '' : `<button class="btn sm danger ghost rw" data-mdel="${esc(m.id)}">Excluir</button>`}</td></tr>`; }).join('') || '<tr><td class="l muted" colspan="9">Nenhuma medição ainda.</td></tr>'}</tbody></table></div>
    ${(c.aditivos || []).length ? `<div class="lbl">Aditivos</div><ul class="lista-simples">${c.aditivos.map(a => `<li>${dBR(a.d)} · ${brl2(a.valor)} · ${esc(a.desc || '')}</li>`).join('')}</ul>` : ''}
    ${(c.liberacoes || []).length ? `<div class="lbl">Retenção liberada</div><ul class="lista-simples">${c.liberacoes.map(a => `<li>${dBR(a.data)} · ${brl2(a.valor)}</li>`).join('')}</ul>` : ''}
    ${r.ts.filter(t => !t.medicao && !t.liberacao).length ? `<div class="lbl">Pagamentos vinculados sem medição</div><div class="tlist">${r.ts.filter(t => !t.medicao && !t.liberacao).slice(0, 30).map(t => linhaTitulo(t, { mostrarPago: true })).join('')}</div>` : ''}`, { largo: true });
  const on = (id, fn) => { const b = f.el.querySelector('#' + id); if (b) b.onclick = () => { f.fechar(); fn(); }; };
  on('cM', () => formMedicao(c)); on('cA', () => formAditivo(c)); on('cR', () => formLiberar(c)); on('cV', () => vincularContrato(c)); on('cE', () => formContrato(c));
  $$('[data-mdel]', f.el).forEach(b => b.onclick = async () => { const m = meds.find(x => x.id === b.dataset.mdel); if (!(await confirmar(`Excluir a medição ${m.n} e as contas a pagar dela?`, 'Excluir'))) return; excluirTitulos(S.list.filter(t => t.medicao === m.id).map(t => t.id)); salvarContrato({ ...c, medicoes: c.medicoes.filter(x => x.id !== m.id) }); f.fechar(); ACOES.ctAbrir(d); });
  f.el.addEventListener('click', e => { if (e.target.closest('[data-act="abrir"]')) f.fechar(); });
};
function formMedicao(c) {
  const r = resumoContrato(c); const n = Math.max(0, ...(c.medicoes || []).map(m => m.n || 0)) + 1;
  const qs = (c.quadras || []).length ? c.quadras : ['-'];
  const f = folha(`Medição ${n} · ${c.forn}`, `<form class="form" id="fx"><p class="muted small">Contratado ${brl2(r.contratado)} · já medido ${brl2(r.medido)} · saldo ${brl2(r.saldo)}</p>
    <div class="g2"><label>Data da medição<input type="date" name="data" value="${hoje()}" required></label><label>Referência<input name="ref" placeholder="Ex.: 2ª quinzena de outubro, alvenaria"></label></div>
    <div class="g3"><label>Valor medido (R$)<input name="valor" inputmode="decimal" required></label><label>ou % do contrato<input name="pct" inputmode="decimal" placeholder="Ex.: 10"></label><label>Físico acumulado (%)<input name="fisico" inputmode="decimal" placeholder="opcional"></label></div>
    <div class="g3"><label>Retenção (%)<input name="ret" inputmode="decimal" value="${String(c.retPct || 0).replace('.', ',')}"></label><label>Vencimento<input type="date" name="venc" value="${proximoDiaUtil(addDays(hoje(), 7))}" required></label><label>Quadra<select name="quadra">${opts(qs.map(q => [q, q === '-' ? 'Geral' : 'Quadra ' + q]), qs[0])}</select></label></div>
    <label class="chk"><input type="checkbox" name="ratear" checked> Dividir entre os lotes da quadra</label>
    ${variasEmpresas() ? `<label>CNPJ da nota / pagamento<select name="emp" required>${opts(empOpts(), c.emp || '', 'Escolha o CNPJ…')}</select></label>` : ''}
    <div class="preview" id="mPrev"></div>
    <div class="acts"><button type="button" class="btn" data-close>Cancelar</button><button class="btn primary">Lançar medição</button></div></form>`);
  const F = f.el.querySelector('#fx');
  const calc = () => { if (F.pct.value && document.activeElement === F.pct) F.valor.value = fNum2.format(r2(r.contratado * parseBR(F.pct.value) / 100)); const v = parseBR(F.valor.value) || 0, rt = r2(v * (parseBR(F.ret.value) || 0) / 100); $('#mPrev', F).innerHTML = v ? `Bruto ${brl2(v)} − retenção ${brl2(rt)} = <b>${brl2(v - rt)} a pagar</b>${v > r.saldo + 0.5 ? '<br><span class="neg">Passa do saldo do contrato: registre um aditivo.</span>' : ''}` : ''; };
  F.addEventListener('input', calc);
  F.onsubmit = e => {
    e.preventDefault(); const v = parseBR(F.valor.value); if (!(v > 0)) { toast('Informe o valor medido.', 'err'); return; }
    if (F.emp && !F.emp.value) { toast('Escolha o CNPJ.', 'err'); return; }
    const rt = r2(v * (parseBR(F.ret.value) || 0) / 100); const venc = proximoDiaUtil(F.venc.value);
    const m = { id: 'md' + uid(), n, data: F.data.value, ref: F.ref.value.trim(), valor: r2(v), ret: rt, fisico: parseBR(F.fisico.value) || 0, venc, quadra: F.quadra.value };
    const ts = titulosMedicao(c, m, venc, v - rt, m.quadra, F.ratear.checked, F.emp ? F.emp.value : c.emp, { titulo: `Medição ${n}${c.objeto ? ' ' + c.objeto : ''}`, medicao: m.id, obs: m.ref });
    salvarTitulos(ts); salvarContrato({ ...c, medicoes: [...(c.medicoes || []), m] }); f.fechar(); toast(`Medição ${n} lançada: ${brl2(v - rt)} a pagar`); ACOES.ctAbrir({ id: c.id });
  };
}
function formAditivo(c) {
  const f = folha('Aditivo · ' + c.forn, `<form class="form" id="fx"><div class="g2"><label>Data<input type="date" name="d" value="${hoje()}"></label><label>Valor (R$, negativo para reduzir)<input name="v" inputmode="decimal" required></label></div><label>Descrição<input name="desc" placeholder="Ex.: muro de divisa"></label><div class="acts"><button type="button" class="btn" data-close>Cancelar</button><button class="btn primary">Salvar</button></div></form>`);
  f.el.querySelector('#fx').onsubmit = e => { e.preventDefault(); const F = e.target; const v = parseBR(F.v.value.replace(/^-/, '')) * (F.v.value.trim().startsWith('-') ? -1 : 1); if (!v) return; salvarContrato({ ...c, aditivos: [...(c.aditivos || []), { d: F.d.value, valor: r2(v), desc: F.desc.value.trim() }] }); f.fechar(); ACOES.ctAbrir({ id: c.id }); };
}
function formLiberar(c) {
  const r = resumoContrato(c);
  const f = folha('Liberar retenção · ' + c.forn, `<form class="form" id="fx"><div class="g2"><label>Valor (R$)<input name="v" inputmode="decimal" value="${fNum2.format(r.retido)}"></label><label>Vencimento<input type="date" name="venc" value="${proximoDiaUtil(addDays(hoje(), 7))}"></label></div>${variasEmpresas() ? `<label>CNPJ<select name="emp" required>${opts(empOpts(), c.emp || '', 'Escolha o CNPJ…')}</select></label>` : ''}<div class="acts"><button type="button" class="btn" data-close>Cancelar</button><button class="btn primary">Gerar conta a pagar</button></div></form>`);
  f.el.querySelector('#fx').onsubmit = e => { e.preventDefault(); const F = e.target; const v = parseBR(F.v.value); if (!(v > 0)) return; const lib = { id: 'lb' + uid(), data: hoje(), valor: r2(v) }; const venc = proximoDiaUtil(F.venc.value);
    salvarTitulos(titulosMedicao(c, null, venc, v, (c.quadras || [])[0] || '-', false, F.emp ? F.emp.value : c.emp, { titulo: 'Liberação de retenção', liberacao: lib.id })); salvarContrato({ ...c, liberacoes: [...(c.liberacoes || []), lib] }); f.fechar(); ACOES.ctAbrir({ id: c.id }); };
}
function vincularContrato(c) {
  const k = normTxt(c.forn).split(' ')[0];
  const cand = S.list.filter(t => !t.contrato && t.valor < 0 && (t.forn === c.forn || (k.length > 2 && normTxt(t.item + ' ' + (t.forn || '')).includes(k)))).sort((a, b) => (dataRef(b) || '').localeCompare(dataRef(a) || ''));
  const f = folha('Vincular pagamentos · ' + c.forn, `<p class="muted small">Lançamentos com o fornecedor ou o nome "${esc(k)}" na descrição, ainda sem contrato. Marque os que pertencem a este contrato.</p>
    <div class="acts left"><button class="btn sm" id="vT">Marcar todos</button></div>
    <div class="vinc-lista">${cand.slice(0, 400).map(t => `<label class="vinc"><input type="checkbox" value="${esc(t.id)}"><span>${dBR(dataRef(t))}</span><span class="wrap">${esc(t.item)}</span><b class="neg">${brl2(t.valor)}</b></label>`).join('') || '<p class="muted">Nada encontrado.</p>'}</div>
    <div class="acts"><button class="btn" data-close>Cancelar</button><button class="btn primary" id="vOk">Vincular</button></div>`, { largo: true });
  f.el.querySelector('#vT').onclick = () => $$('.vinc input', f.el).forEach(x => x.checked = true);
  f.el.querySelector('#vOk').onclick = () => { const ids = new Set($$('.vinc input:checked', f.el).map(x => x.value)); salvarTitulos(S.list.filter(t => ids.has(t.id)).map(t => ({ ...t, contrato: c.id }))); f.fechar(); toast(ids.size + ' lançamento(s) vinculados'); ACOES.ctAbrir({ id: c.id }); };
}
BUSCA_EXTRA.push(q => (S.contratos || []).filter(c => normTxt(c.forn + ' ' + (c.objeto || '')).includes(q)).map(c => ({ titulo: 'Contrato · ' + c.forn, sub: c.objeto || '', href: '#/contratos' })));
