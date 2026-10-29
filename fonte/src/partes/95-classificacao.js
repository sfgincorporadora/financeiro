/* =========================================================
   Classificação: despesas da empresa x obra, regras automáticas,
   juntar lançamentos rateados em lotes, transferências entre contas
   e rateio gerencial da administração entre os lotes ativos.
   ========================================================= */
usarConfig('regras');
usarConfig('rateio');

/* ---------- Palavras que indicam despesa da empresa (não de um lote) ---------- */
const PALAVRAS_EMPRESA = ['cim', 'strada', 'container', 'claude', 'chatgpt', 'contador', 'contabilidade', 'honorarios contabeis', 'internet', 'aluguel', 'ipva', 'licenciamento', 'combustivel', 'abastecimento', 'gasolina', 'diesel', 'software', 'assinatura', 'tarifa', 'pro labore', 'escritorio', 'despachante', 'plotagem', 'seguro do carro', 'celular', 'telefone', 'energia escritorio', 'google workspace', 'dominio', 'anuidade', 'crea', 'cau', 'sindicato'];
const casaTxt = (txt, termo) => { const a = ' ' + normTxt(txt) + ' ', b = normTxt(termo); return !!b && a.includes(' ' + b + ' '); };
const regras = () => ((S.regras && S.regras.lista) || []);
const regraDe = t => regras().find(r => r.txt && (r.so !== 'saida' || t.valor < 0) && (r.so !== 'entrada' || t.valor > 0) && casaTxt(t.item + ' ' + (t.forn || ''), r.txt));
const pareceEmpresa = t => t.valor < 0 && !ehEmpresa(t) && t.destino !== 'Impostos' && secaoDe(t.destino) !== 'aporte' && (PALAVRAS_EMPRESA.some(p => casaTxt(baseItem(t), p)) || t.destino === 'Administrativo' || (regraDe(t) || {}).ambito === 'empresa');

/* Aplica a regra que casar com a descrição do formulário de novo lançamento */
function aplicarRegrasForm(F, tipo) {
  const r = regraDe({ item: F.item.value, forn: F.forn.value, valor: tipo === 'saida' ? -1 : 1 }); if (!r) return;
  const feito = [];
  if (r.ambito && F.ambito && F.ambito.value !== r.ambito) { F.ambito.value = r.ambito; F.ambito.dispatchEvent(new Event('change')); feito.push(r.ambito === 'empresa' ? 'despesa da empresa' : 'obra'); }
  if (r.destino && F.destino && todosDestinos().includes(r.destino)) { F.destino.value = r.destino; F.destino.dispatchEvent(new Event('change')); feito.push(r.destino); }
  if (r.forn && F.forn && !F.forn.value) { F.forn.value = r.forn; feito.push(r.forn); }
  if (r.emp && F.emp && empPorId(r.emp)) { F.emp.value = r.emp; F.emp.dispatchEvent(new Event('change')); feito.push(empPorId(r.emp).nome); }
  if (r.etapa && F.etapaObra && r.ambito !== 'empresa') { F.etapaObra.value = r.etapa; feito.push(r.etapa); }
  if (feito.length) toast('Regra "' + r.txt + '" aplicada: ' + feito.join(', '));
}

/* ---------- Juntar parcelas de um pagamento que foi dividido entre lotes ---------- */
const chaveJuntar = t => [normTxt(baseItem(t).replace(new RegExp('\\s*[-–]\\s*' + reEsc(t.quadra || '§') + '\\s*$'), '')), pago(t) ? (t.pag || t.venc) : t.venc, t.status, pago(t) ? (t.conta || '') : '', t.destino, Math.sign(t.valor), t.grupo || ''].join('|');
function gruposJuntar(ts) {
  const g = {}; for (const t of ts) (g[chaveJuntar(t)] = g[chaveJuntar(t)] || []).push(t);
  return Object.values(g);
}
function juntarEmEmpresa(ts, extra = {}) {
  const out = [], apagar = [];
  for (const xs of gruposJuntar(ts)) {
    const t0 = xs[0]; const anexos = []; const vistos = new Set();
    for (const t of xs) for (const a of anexosDe(t)) { const k = a.id || a.link || a.drive; if (!vistos.has(k)) { vistos.add(k); anexos.push(a); } }
    const base = baseItem(t0).replace(new RegExp('\\s*[-–]\\s*' + reEsc(t0.quadra || '§') + '\\s*$'), '').trim();
    const n = { ...t0, item: base, valor: r2(sum(xs, t => t.valor)), juros: r2(sum(xs, t => t.juros || 0)), quadra: '-', lote: '-', local: '', ambito: 'empresa', ...extra };
    if (anexos.length) { n.anexos = anexos; delete n.anexo; delete n.anexoLink; delete n.anexoDrive; }
    const nfs = [...new Set(xs.map(t => t.nfNum).filter(Boolean))]; if (nfs.length) n.nfNum = nfs.join(', '); if (xs.some(t => t.nf)) n.nf = 1;
    if (xs.some(t => t.qtd)) n.qtd = r2(sum(xs, t => t.qtd || 0));
    if (!extra.emp && !(empDef(t0) && empPorId(t0.emp)?.tipo === 'matriz')) n.emp = matrizId();
    delete n.etapaObra; if (xs.length > 1) delete n.grupo;
    out.push(n); apagar.push(...xs.slice(1).map(t => t.id));
  }
  salvarTitulos(out, { acao: 'reclassificou como despesa da empresa' });
  if (apagar.length) excluirTitulos(apagar);
  return { grupos: out.length, removidos: apagar.length };
}

/* ---------- Transferência entre contas (mesma empresa ou entre CNPJs) ---------- */
ACOES.transferir = d => {
  const cs = E().contas.map(c => c.nome); if (cs.length < 2) { toast('Cadastre pelo menos duas contas para transferir.', 'err'); return; }
  const f = folha('Transferência entre contas', `<form class="form" id="fx">
    <div class="g2"><label>Sai da conta<select name="de">${opts(cs, d.de || cs[0])}</select></label><label>Entra na conta<select name="para">${opts(cs, cs.find(c => c !== (d.de || cs[0])))}</select></label></div>
    <div class="g2"><label>Valor (R$)<input name="valor" inputmode="decimal" required placeholder="0,00"></label><label>Data<input type="date" name="data" value="${hoje()}" required></label></div>
    <label>Descrição<input name="item" placeholder="Ex.: Transferência para a filial, aporte do sócio"></label>
    <p class="hint-data info" id="trHint"></p>
    <div class="acts"><button type="button" class="btn" data-close>Cancelar</button><button class="btn primary">Transferir</button></div></form>`);
  const F = f.el.querySelector('#fx');
  const info = () => { const a = empConta(F.de.value), b = empConta(F.para.value); const ea = empPorId(a), eb = empPorId(b);
    $('#trHint', F).textContent = F.de.value === F.para.value ? 'Escolha contas diferentes.' : a === b ? `Mesma empresa (${ea?.nome || ''}): fica registrada como transferência, não entra em receita nem despesa.` : `Entre empresas: sai de ${ea?.nome || ''} e entra em ${eb?.nome || ''} como aporte entre empresas do grupo. No consolidado as duas pontas se anulam.`; };
  F.de.onchange = info; F.para.onchange = info; info();
  F.onsubmit = e => {
    e.preventDefault(); const v = parseBR(F.valor.value); if (!(v > 0) || F.de.value === F.para.value) { toast('Confira o valor e as contas.', 'err'); return; }
    const a = empConta(F.de.value), b = empConta(F.para.value); const mesma = a === b; const g = uid();
    const desc = F.item.value.trim() || (mesma ? 'Transferência' : 'Transferência entre empresas') + ` ${F.de.value} → ${F.para.value}`;
    const base = { item: desc, juros: 0, destino: mesma ? 'Transferência' : 'Aporte', aporteTipo: mesma ? 'transferencia' : 'intragrupo', venc: F.data.value, pag: F.data.value, status: 'pago', quadra: '-', lote: '-', local: '', ambito: 'empresa', grupo: g, forn: '', obs: '', criado: Date.now() };
    salvarTitulos([{ ...base, id: uid(), valor: -r2(v), conta: F.de.value, emp: a }, { ...base, id: uid(), valor: r2(v), conta: F.para.value, emp: b }]);
    f.fechar(); toast('Transferência registrada');
  };
};

/* ---------- Rateio gerencial da administração ---------- */
const ehAdm = t => pago(t) && t.valor < 0 && secaoDe(t.destino) !== 'aporte' && t.destino !== 'Impostos' && (ehEmpresa(t) || (t.destino === 'Administrativo' && (!t.quadra || t.quadra === '-')));
let _rat = { ref: null, crit: '' };
/* Só gasto de construção marca o lote como "em obra" no mês (terreno, ITBI, cartório e registro não contam). */
const DEST_OBRA_ATIVA = ['Mão de Obra', 'Material', 'Pavimentação', 'Operacional'];
function rateioAdm() {
  const crit = (S.rateio && S.rateio.criterio) || 'igual';
  if (_rat.ref === S.list && _rat.crit === crit) return _rat.r;
  const meses = {}; // m -> { adm, lotes: Map(k -> peso) }
  const M = m => (meses[m] = meses[m] || { adm: 0, lotes: new Map(), quadras: new Map() });
  for (const t of S.list) {
    const d = (t.pag || t.venc || '').slice(0, 7); if (!d || !pago(t)) continue;
    if (ehAdm(t)) { M(d).adm += -(t.valor + Math.min(0, t.juros || 0)); continue; }
    if (t.valor < 0 && DEST_OBRA_ATIVA.includes(t.destino) && t.quadra && t.quadra !== '-' && E().quadras.includes(t.quadra)) {
      const m = M(d);
      if (t.lote && t.lote !== '-') m.lotes.set(t.quadra + '|' + t.lote, (m.lotes.get(t.quadra + '|' + t.lote) || 0) - t.valor);
      else m.quadras.set(t.quadra, (m.quadras.get(t.quadra) || 0) - t.valor);
    }
  }
  // lote concluído (repasse recebido): fica fora a partir do mês seguinte ao último recebimento
  const fim = {}; for (const q of E().quadras) for (const l of lotesDe(q)) { const i = infoVendaLote(q, l); if (i.etapa === 'recebido' && i.ult) fim[q + '|' + l] = i.ult.slice(0, 7); }
  const porLote = {}, porMes = [];
  for (const m of Object.keys(meses).sort()) {
    const x = meses[m]; const peso = new Map(x.lotes);
    for (const [q, v] of x.quadras) { const ls = lotesDe(q); for (const l of ls) peso.set(q + '|' + l, (peso.get(q + '|' + l) || 0) + v / ls.length); }
    for (const k of [...peso.keys()]) if (fim[k] && fim[k] < m) peso.delete(k);
    const tot = sum([...peso.values()], v => v); const n = peso.size;
    let rateado = 0;
    if (x.adm > 0 && n) for (const [k, v] of peso) { const parte = crit === 'custo' && tot > 0 ? x.adm * v / tot : x.adm / n; porLote[k] = (porLote[k] || 0) + parte; rateado += parte; }
    porMes.push({ m, adm: x.adm, n, rateado, sem: x.adm - rateado });
  }
  const r = { porLote, porMes, crit, total: sum(porMes, x => x.adm), semLotes: sum(porMes, x => x.sem) };
  _rat = { ref: S.list, crit, r }; return r;
}
const admLote = (q, l) => (rateioAdm().porLote[q + '|' + l] || 0);
const admQuadra = q => sum(lotesDe(q), l => admLote(q, l));

/* ---------- Tela "Revisar e classificar" ---------- */
const TIPOS_REVISAR = [['empresa', 'Despesas da empresa em lotes'], ['nf', 'Pagos sem NF'], ['forn', 'Sem fornecedor'], ['etapa', 'Sem etapa'], ['cnpj', 'Sem CNPJ'], ['aporte', 'Aportes sem tipo'], ['contas', 'Contas bancárias'], ['anexo', 'Sem anexo'], ['regras', 'Regras automáticas'], ['rateio', 'Rateio da administração']];
let revBusca = '', revNFPer = '12';
function candidatosRevisar(tipo) {
  if (tipo === 'empresa') return S.list.filter(t => pareceEmpresa(t) && t.quadra && t.quadra !== '-');
  if (tipo === 'nf') { const lim = revNFPer === 'tudo' ? '' : addMonths(hoje(), -(+revNFPer)); return S.list.filter(t => pago(t) && t.valor < 0 && !t.nfNum && !t.nf && !t.semNota && !anexosDe(t).length && ['op', 'nonop'].includes(secaoDe(t.destino)) && t.destino !== 'Impostos' && (t.pag || t.venc) >= lim); }
  if (tipo === 'aporte') return S.list.filter(t => secaoDe(t.destino) === 'aporte' && !t.aporteTipo);
  if (tipo === 'contas') return E().contas.filter(c => !c.num || (variasEmpresas() && !c.emp));
  if (tipo === 'regras' || tipo === 'rateio') return regras();
  return candidatosArrumar(tipo);
}
VIEWS.arrumar = {
  titulo: 'Revisar e classificar',
  refresh(el) {
    const tipos = TIPOS_REVISAR.filter(([k]) => k !== 'cnpj' || variasEmpresas());
    if (!tipos.some(([k]) => k === arrTipo)) arrTipo = 'empresa';
    const cont = k => k === 'regras' ? regras().length : k === 'rateio' ? '' : ['nf', 'empresa'].includes(k) ? gruposJuntar(candidatosRevisar(k)).length : candidatosRevisar(k).length;
    const abas = `<div class="seg rev-tabs">${tipos.map(([k, n]) => { const c = cont(k); return `<button class="${k === arrTipo ? 'on' : ''}" data-act="arrTipo" data-v="${k}">${n}${c !== '' ? ` <span class="cnt${c && !['regras'].includes(k) ? ' on' : ''}">${c.toLocaleString('pt-BR')}</span>` : ''}</button>`; }).join('')}</div>`;
    const fns = { empresa: revEmpresa, nf: revNF, aporte: revAporte, contas: revContas, regras: revRegras, rateio: revRateio };
    if (fns[arrTipo]) { el.innerHTML = abas + '<div id="revBox"></div>'; fns[arrTipo]($('#revBox', el)); return; }
    VIEWS_ARRUMAR_ANTIGO.refresh(el, abas);
  }
};
function revEmpresa(box) {
  const q = normTxt(revBusca);
  const ts = candidatosRevisar('empresa').filter(t => !q || (' ' + normTxt(t.item + ' ' + (t.forn || '')) + ' ').includes(q));
  const gs = {}; for (const t of ts) (gs[chaveGrupo(t)] = gs[chaveGrupo(t)] || []).push(t);
  const lista = Object.entries(gs).map(([k, xs]) => ({ k, xs, tot: sum(xs, t => t.valor), pag: gruposJuntar(xs).length, qs: [...new Set(xs.map(t => t.quadra))] })).sort((a, b) => a.tot - b.tot);
  box.innerHTML = `<div class="card"><h2>Despesas da empresa lançadas em lotes <span class="muted small">${ts.length.toLocaleString('pt-BR')} lançamento(s), ${brl(-sum(ts, t => t.valor))}</span></h2>
    <p class="muted small">Contas como CIM, financiamento e despesas do carro, frete do container, assinaturas, contador e escritório são da empresa, não de um lote. Aqui elas aparecem divididas entre lotes porque foram lançadas assim. <b>Juntar</b> soma as partes de cada pagamento em um único lançamento da empresa (CNPJ da matriz, sem quadra e lote). Os lançamentos originais vão para a lixeira e o histórico registra a mudança.</p>
    <p class="muted small">Se a despesa for mesmo de uma obra específica (ex.: frete de material para um lote), use <b>Não é da empresa</b>: ela some desta lista.</p>
    <div class="rev-busca"><input type="search" id="rvQ" placeholder="Filtrar: strada, container, cim…" value="${esc(revBusca)}">${lista.length ? `<button class="btn sm primary rw" id="rvTodos">Juntar todos os ${lista.length} grupos</button>` : ''}</div>
    <div class="arr-lista">${lista.slice(0, 200).map((g, i) => `<div class="arr-g"><div class="arr-txt"><b>${esc(baseItem(g.xs[0]))}</b><small>${g.xs.length} parte(s) em ${g.qs.length} quadra(s) (${esc(g.qs.slice(0, 5).join(', '))}) · ${g.pag} pagamento(s) · <b class="neg">${brl2(g.tot)}</b> · ${esc(g.xs[0].destino)}</small></div>
      <button class="btn sm ghost rw" data-nao="${i}">Não é da empresa</button><button class="btn sm rw" data-ver="${i}">Ver</button><button class="btn sm primary rw" data-jt="${i}">${g.pag === g.xs.length ? 'Mover para a empresa' : `Juntar em ${g.pag}`}</button></div>`).join('') || '<p class="muted">Nenhuma despesa da empresa espalhada em lotes.</p>'}</div></div>`;
  const rq = $('#rvQ', box); rq.oninput = () => { revBusca = rq.value; clearTimeout(rq._t); rq._t = setTimeout(() => { revEmpresa(box); const n = $('#rvQ', box); n.focus(); n.setSelectionRange(n.value.length, n.value.length); }, 250); };
  const juntar = async gs => { const xs = gs.flatMap(g => g.xs); const np = sum(gs, g => g.pag); if (!(await confirmar(`Juntar <b>${xs.length}</b> lançamento(s) em <b>${np}</b> lançamento(s) da empresa, total ${brl2(sum(xs, t => t.valor))}?`, 'Juntar', false))) return; const r = juntarEmEmpresa(xs); toast(`${r.grupos} lançamento(s) da empresa criados, ${r.removidos} parte(s) movida(s) para a lixeira`); };
  $$('[data-jt]', box).forEach(b => b.onclick = () => juntar([lista[+b.dataset.jt]]));
  $$('[data-ver]', box).forEach(b => b.onclick = () => { const g = lista[+b.dataset.ver]; folha(baseItem(g.xs[0]), `<div class="tlist">${g.xs.slice(0, 120).map(t => linhaTitulo(t, { mostrarPago: true })).join('')}</div>`); });
  $$('[data-nao]', box).forEach(b => b.onclick = () => { const g = lista[+b.dataset.nao]; salvarTitulos(g.xs.map(t => ({ ...t, ambito: 'obra', destino: t.destino === 'Administrativo' ? 'Operacional' : t.destino })), { acao: 'confirmou como custo da obra' }); toast('Mantido como custo da obra'); });
  const bt = $('#rvTodos', box); if (bt) bt.onclick = () => juntar(lista);
}
function revNF(box) {
  const ts = candidatosRevisar('nf'); const gs = gruposJuntar(ts).map(xs => ({ xs, t: xs[0], tot: sum(xs, t => t.valor) })).sort((a, b) => (b.t.pag || b.t.venc).localeCompare(a.t.pag || a.t.venc)).slice(0, 150);
  box.innerHTML = `<div class="card"><h2>Pagamentos sem número de NF <span class="muted small">${ts.length.toLocaleString('pt-BR')} lançamento(s) em ${gruposJuntar(ts).length.toLocaleString('pt-BR')} pagamento(s)</span></h2>
    <div class="rev-busca"><select id="nfPer">${opts([['3', 'Últimos 3 meses'], ['12', 'Últimos 12 meses'], ['tudo', 'Todo o histórico']], revNFPer)}</select></div>
    <p class="muted small">Pagamentos de obra e administrativos sem número de nota, sem a marcação "tem nota fiscal" e sem anexo. Digite o número da nota (vale para todas as partes do mesmo pagamento) ou marque <b>Sem nota</b> quando a despesa não tem nota mesmo (tarifa, guia, recibo). Para anexar o arquivo, toque em <b>Abrir</b>.</p>
    <div class="arr-lista">${gs.map((g, i) => `<div class="arr-g"><div class="arr-txt"><b>${esc(baseItem(g.t))}</b><small>${dBR(g.t.pag || g.t.venc)} · ${esc(g.t.forn || 'sem fornecedor')} · <b class="neg">${brl2(g.tot)}</b>${g.xs.length > 1 ? ` · ${g.xs.length} partes` : ''}${anexosDe(g.t).length ? ' · com anexo' : ''}</small></div>
      <input data-nf="${i}" inputmode="numeric" placeholder="Nº da NF" style="max-width:120px"><button class="btn sm primary rw" data-snf="${i}">Salvar</button><button class="btn sm ghost rw" data-sem="${i}">Sem nota</button><button class="btn sm" data-ab="${i}">Abrir</button></div>`).join('') || '<p class="muted">Todos os pagamentos têm NF ou foram marcados sem nota.</p>'}</div></div>`;
  $('#nfPer', box).onchange = e => { revNFPer = e.target.value; refreshView(); };
  $$('[data-snf]', box).forEach(b => b.onclick = () => { const g = gs[+b.dataset.snf]; const v = $(`[data-nf="${b.dataset.snf}"]`, box).value.trim(); if (!v) { toast('Digite o número da NF', 'err'); return; } salvarTitulos(g.xs.map(t => ({ ...t, nf: 1, nfNum: v }))); toast('NF registrada'); });
  $$('[data-sem]', box).forEach(b => b.onclick = () => { const g = gs[+b.dataset.sem]; salvarTitulos(g.xs.map(t => ({ ...t, semNota: 1 }))); toast('Marcado como sem nota'); });
  $$('[data-ab]', box).forEach(b => b.onclick = () => { const g = gs[+b.dataset.ab]; detalhe(g.t); });
}
function revAporte(box) {
  const ts = candidatosRevisar('aporte'); const gs = {}; for (const t of ts) (gs[chaveGrupo(t) + (t.valor > 0 ? '+' : '-')] = gs[chaveGrupo(t) + (t.valor > 0 ? '+' : '-')] || []).push(t);
  const lista = Object.values(gs).map(xs => ({ xs, tot: sum(xs, t => t.valor) })).sort((a, b) => Math.abs(b.tot) - Math.abs(a.tot));
  const sug = g => { const s = normTxt(g.xs[0].item); return /transf/.test(s) ? 'transferencia' : /emprest/.test(s) ? (g.tot > 0 ? 'emprestimo' : 'devolucao') : /retirad|distrib|lucro/.test(s) ? 'retirada' : /filial|matriz|ret /.test(s) ? 'intragrupo' : g.tot > 0 ? 'capital' : 'retirada'; };
  box.innerHTML = `<div class="card"><h2>Aportes sem tipo <span class="muted small">${ts.length} lançamento(s)</span></h2>
    <p class="muted small">Aporte não é receita: é dinheiro que entra ou sai do caixa por decisão dos sócios ou entre empresas do grupo. O tipo define onde ele aparece no BP e na DFC: <b>capital</b> aumenta o patrimônio líquido; <b>empréstimo do sócio</b> é uma dívida da empresa; <b>entre empresas do grupo</b> se anula no consolidado; <b>retirada</b> reduz o patrimônio.</p>
    <div class="arr-lista">${lista.map((g, i) => `<div class="arr-g"><div class="arr-txt"><b>${esc(g.xs[0].item)}</b><small>${g.xs.length} lançamento(s) · <b class="${cls(g.tot)}">${brl2(g.tot)}</b> · ${dBR(g.xs[0].pag || g.xs[0].venc)}</small></div><select data-g="${i}">${opts(TIPOS_APORTE, sug(g))}</select><button class="btn sm rw" data-ap="${i}">Aplicar</button></div>`).join('') || '<p class="muted">Todos os aportes têm tipo.</p>'}</div></div>`;
  $$('[data-ap]', box).forEach(b => b.onclick = () => { const g = lista[+b.dataset.ap]; const v = $(`[data-g="${b.dataset.ap}"]`, box).value; salvarTitulos(g.xs.map(t => ({ ...t, aporteTipo: v }))); toast('Tipo de aporte salvo'); });
}
function revContas(box) {
  const cs = E().contas; const sd = saldos().contas;
  box.innerHTML = `<div class="card"><h2>Contas bancárias <button class="btn sm rw" data-act="cadConta">Adicionar conta</button><button class="btn sm rw" data-act="transferir">Transferir entre contas</button></h2>
    <p class="muted small">Cada conta precisa do banco, agência, número e da empresa (CNPJ) titular. Com o número cadastrado, o app reconhece a conta ao importar o extrato OFX.</p>
    <div class="tbl-wrap"><table class="tbl"><thead><tr><th class="l">Conta</th><th class="l">Banco · agência · conta</th><th class="l">Empresa</th><th>Saldo</th><th></th></tr></thead><tbody>
    ${cs.map(c => `<tr><td class="l"><b>${esc(c.nome)}</b></td><td class="l">${c.num ? esc([c.banco, c.ag ? 'ag. ' + c.ag : '', 'cc ' + c.num].filter(Boolean).join(' · ')) : '<span style="color:var(--warn)">falta o número</span>'}</td><td class="l">${c.emp || !variasEmpresas() ? esc(empRotulo(empPorId(empConta(c.nome)), true)) : '<span style="color:var(--warn)">falta a empresa</span>'}</td><td>${brl2((sd.find(x => x.nome === c.nome) || {}).saldo || 0)}</td><td><button class="btn sm" data-act="cadConta" data-nome="${esc(c.nome)}">Editar</button></td></tr>`).join('')}</tbody></table></div></div>`;
}
function revRegras(box) {
  const rs = regras();
  const alvo = r => S.list.filter(t => regraDe(t) === r && (
    (r.ambito === 'empresa' && !ehEmpresa(t)) || (r.destino && t.destino !== r.destino) || (r.forn && !t.forn) || (r.emp && t.emp !== r.emp) || (r.etapa && r.ambito !== 'empresa' && !t.etapaObra)));
  box.innerHTML = `<div class="card"><h2>Regras automáticas <button class="btn sm primary rw" id="rgNova">Nova regra</button></h2>
    <p class="muted small">Uma regra diz: <b>quando a descrição ou o fornecedor contém</b> certa palavra, o lançamento é classificado de um jeito. Ela é aplicada sozinha quando você digita a descrição de um lançamento novo, e pode ser aplicada aos lançamentos que já existem (com prévia antes de mudar).</p>
    <div class="arr-lista">${rs.map((r, i) => { const n = alvo(r).length; return `<div class="arr-g"><div class="arr-txt"><b>“${esc(r.txt)}”</b><small>${[r.ambito === 'empresa' ? 'Despesa da empresa' : r.ambito === 'obra' ? 'Obra' : '', r.destino, r.forn, r.emp && empPorId(r.emp) ? empPorId(r.emp).nome : '', r.etapa].filter(Boolean).map(esc).join(' · ')}${n ? ` · <b>${n} lançamento(s) fora da regra</b>` : ''}</small></div>
      ${n ? `<button class="btn sm primary rw" data-apl="${i}">Aplicar (${n})</button>` : ''}<button class="btn sm rw" data-ed="${i}">Editar</button></div>`; }).join('') || '<p class="muted">Nenhuma regra ainda. Sugestões: “cim”, “strada”, “container”, “claude”, “contador”, todas como despesa da empresa.</p>'}</div>
    ${rs.length ? '' : '<button class="btn rw" id="rgSug">Criar as regras sugeridas</button>'}</div>`;
  const salvar = lista => gravar(doc(db, 'config', 'regras'), { lista }).catch(onErroEscrita);
  const editar = i => {
    const r = i === undefined ? { id: uid(), so: 'saida' } : rs[i];
    const f = folha(i === undefined ? 'Nova regra' : 'Editar regra', `<form class="form" id="fx"><label>Quando a descrição ou o fornecedor contém<input name="txt" required value="${esc(r.txt || '')}" placeholder="Ex.: strada"></label>
      <label>Vale para<select name="so">${opts([['saida', 'Saídas'], ['entrada', 'Entradas'], ['', 'Saídas e entradas']], r.so ?? 'saida')}</select></label>
      <label>A quem pertence<select name="ambito">${opts([['', 'Não mudar'], ['empresa', 'Empresa (administrativo)'], ['obra', 'Obra']], r.ambito || '')}</select></label>
      <div class="g2"><label>Destino<select name="destino"><option value="">Não mudar</option>${optsDestino(r.destino || '')}</select></label><label>Fornecedor<input name="forn" value="${esc(r.forn || '')}"></label></div>
      <div class="g2">${variasEmpresas() ? `<label>CNPJ<select name="emp">${opts(empOpts(), r.emp || '', 'Não mudar')}</select></label>` : ''}<label>Etapa da obra<select name="etapa">${opts(etapasOrc(), r.etapa || '', 'Não mudar')}</select></label></div>
      <div class="acts">${i !== undefined ? '<button type="button" class="btn danger ghost" id="rm">Remover</button>' : ''}<button type="button" class="btn" data-close>Cancelar</button><button class="btn primary">Salvar</button></div></form>`);
    const F = f.el.querySelector('#fx'); if (!F.destino.value) F.destino.value = '';
    F.onsubmit = e => { e.preventDefault(); const n = { id: r.id || uid(), txt: F.txt.value.trim(), so: F.so.value, ambito: F.ambito.value, destino: F.destino.value, forn: F.forn.value.trim(), emp: F.emp ? F.emp.value : '', etapa: F.etapa.value }; const l = rs.slice(); if (i === undefined) l.push(n); else l[i] = n; salvar(l); f.fechar(); toast('Regra salva'); };
    const rm = f.el.querySelector('#rm'); if (rm) rm.onclick = () => { salvar(rs.filter((_, j) => j !== i)); f.fechar(); };
  };
  const nv = $('#rgNova', box); if (nv) nv.onclick = () => editar();
  $$('[data-ed]', box).forEach(b => b.onclick = () => editar(+b.dataset.ed));
  const sg = $('#rgSug', box); if (sg) sg.onclick = () => { salvar(['cim', 'strada', 'container', 'claude', 'contador', 'internet', 'aluguel escritorio', 'ipva', 'abastecimento', 'tarifa'].map(txt => ({ id: uid(), txt, so: 'saida', ambito: 'empresa', destino: E().destinos.nonop?.includes('Administrativo') ? 'Administrativo' : '' }))); toast('Regras criadas: confira e ajuste'); };
  $$('[data-apl]', box).forEach(b => b.onclick = async () => {
    const r = rs[+b.dataset.apl]; const xs = alvo(r);
    const ex = xs.slice(0, 6).map(t => `<li>${esc(t.item)} · ${brl2(t.valor)}</li>`).join('');
    if (!(await confirmar(`Aplicar a regra “${esc(r.txt)}” em <b>${xs.length}</b> lançamento(s)?${r.ambito === 'empresa' ? ' As partes de um mesmo pagamento divididas entre lotes serão juntadas em um único lançamento da empresa.' : ''}<ul class="small">${ex}${xs.length > 6 ? '<li>…</li>' : ''}</ul>`, 'Aplicar', false))) return;
    const extra = {}; if (r.destino) extra.destino = r.destino; if (r.forn) extra.forn = r.forn; if (r.emp) extra.emp = r.emp;
    if (r.ambito === 'empresa') { const lotes = xs.filter(t => t.quadra && t.quadra !== '-'); const resto = xs.filter(t => !lotes.includes(t)); if (lotes.length) juntarEmEmpresa(lotes, extra); if (resto.length) salvarTitulos(resto.map(t => ({ ...t, ...extra, forn: r.forn && !t.forn ? r.forn : t.forn, ambito: 'empresa', quadra: '-', lote: '-', local: '', emp: r.emp || (empDef(t) ? t.emp : matrizId()) })), { acao: 'regra ' + r.txt }); }
    else salvarTitulos(xs.map(t => { const n = { ...t, ...extra }; if (r.forn && t.forn) n.forn = t.forn; if (r.ambito === 'obra') delete n.ambito; if (r.etapa && !t.etapaObra) n.etapaObra = r.etapa; return n; }), { acao: 'regra ' + r.txt });
    toast('Regra aplicada');
  });
}
function revRateio(box) {
  const R = rateioAdm(); const ult = R.porMes.filter(x => x.adm > 0).slice(-12).reverse();
  const top = Object.entries(R.porLote).sort((a, b) => b[1] - a[1]).slice(0, 12);
  box.innerHTML = `<div class="card"><h2>Rateio gerencial da administração</h2>
    <p class="muted small">As despesas da empresa continuam registradas na empresa (é assim que vão para a contabilidade). Para saber quanto cada lote realmente custou, o app distribui essas despesas <b>só entre os lotes que estavam em obra no mês</b>: lotes com gasto de construção pago naquele mês (material, mão de obra, pavimentação ou operacional; terreno, ITBI e cartório não contam). Lotes ainda não iniciados não recebem nada, e lotes com repasse recebido deixam de receber a partir do mês seguinte. Esse rateio é só uma visão: nenhum lançamento é alterado.</p>
    <div class="g2"><label>Critério<select id="rtCrit">${opts([['igual', 'Partes iguais entre os lotes ativos no mês'], ['custo', 'Proporcional ao custo de obra do lote no mês']], R.crit)}</select></label><div></div></div>
    <div class="grid kpis">${kpi('Administração paga', brl(R.total), 'todas as datas', 'neg')}${kpi('Rateada entre lotes', brl(R.total - R.semLotes), '', '')}${kpi('Sem lote ativo no mês', brl(R.semLotes), 'fica só na empresa', R.semLotes > 1 ? 'neg' : '')}</div>
    <div class="grid two"><div><div class="lbl">Últimos meses</div><div class="tbl-wrap"><table class="tbl"><thead><tr><th class="l">Mês</th><th>Administração</th><th>Lotes ativos</th><th>Por lote</th></tr></thead><tbody>${ult.map(x => `<tr><td class="l">${mesNome(x.m)}</td><td class="neg">${brl(x.adm)}</td><td>${x.n}</td><td>${x.n ? brl(x.rateado / x.n) : '<span class="muted">sem lote ativo</span>'}</td></tr>`).join('') || '<tr><td colspan="4" class="l muted">Nenhuma despesa da empresa paga.</td></tr>'}</tbody></table></div></div>
    <div><div class="lbl">Lotes que mais receberam</div><div class="tbl-wrap"><table class="tbl"><thead><tr><th class="l">Lote</th><th>Administração rateada</th></tr></thead><tbody>${top.map(([k, v]) => { const [q, l] = k.split('|'); return `<tr class="click" data-act="ir" data-href="#/obras/${encodeURIComponent(q)}/${encodeURIComponent(l)}"><td class="l">${esc(q)} ${esc(l)}</td><td class="neg">${brl(v)}</td></tr>`; }).join('') || '<tr><td colspan="2" class="l muted">—</td></tr>'}</tbody></table></div></div></div></div>`;
  $('#rtCrit', box).onchange = e => { gravar(doc(db, 'config', 'rateio'), { criterio: e.target.value }, { merge: true }).catch(onErroEscrita); S.rateio = { ...(S.rateio || {}), criterio: e.target.value }; revRateio(box); };
}
