/* =========================================================
   Dados cadastrais de obras e lotes: localização, matrícula,
   construtor; resumo da venda (corretor, correspondente, agência).
   ========================================================= */
usarConfig('lotesInfo');
const loteInfo = (q, l) => ((S.lotesInfo || {})[q + '|' + l]) || {};
const obraLocal = o => { const i = E().obraInfo[o] || {}; return [i.cidade && i.uf ? i.cidade + ' - ' + i.uf : i.cidade || i.uf, i.construtor ? 'construtor ' + i.construtor : ''].filter(Boolean).join(' · '); };
const construtores = () => [...new Set([...Object.values(E().obraInfo).map(i => i.construtor), ...Object.values(S.lotesInfo || {}).map(i => i.construtor), ...(S.contratos || []).map(c => c.forn)].filter(Boolean))].sort();
const construtorLote = (q, l) => loteInfo(q, l).construtor || (E().obraInfo[obraDe(q)] || {}).construtor || '';
const CAMPOS_LOTE = [['endereco', 'Endereço (como na certidão)'], ['matricula', 'Matrícula'], ['cartorio', 'Cartório'], ['inscricao', 'Inscrição imobiliária (IPTU)'], ['area', 'Área do terreno (m²)'], ['construida', 'Área construída (m²)'], ['frente', 'Medidas e confrontações'], ['coord', 'Coordenadas / mapa'], ['construtor', 'Construtor da casa'], ['obs', 'Observação']];
function cardDadosLote(q, l) {
  const i = loteInfo(q, l), o = E().obraInfo[obraDe(q)] || {}, v = ((S.vendas && S.vendas.lotes) || {})[q + '|' + l] || {};
  const linha = (k, val) => val ? `<dt>${esc(k)}</dt><dd>${esc(val)}</dd>` : '';
  const mapa = i.coord || o.mapa; const link = mapa ? (/^https?:/.test(mapa) ? mapa : 'https://www.google.com/maps?q=' + encodeURIComponent(mapa)) : '';
  return `<div class="grid two"><div class="card"><h2>Dados do lote <button class="btn sm rw" data-act="loteInfo" data-q="${esc(q)}" data-l="${esc(l)}">Editar</button></h2>
    <dl class="dl">${linha('Obra', obraDe(q) + (o.cidade ? ` · ${o.cidade}${o.uf ? ' - ' + o.uf : ''}` : ''))}${CAMPOS_LOTE.filter(([k]) => k !== 'coord' && k !== 'construtor').map(([k, n]) => linha(n, i[k])).join('')}${linha('Construtor', construtorLote(q, l) + (!i.construtor && o.construtor ? ' (da obra)' : ''))}</dl>
    ${link ? `<a class="btn sm" target="_blank" rel="noopener" href="${esc(link)}">Abrir no mapa</a>` : ''}${!Object.keys(i).length ? '<p class="muted small">Sem dados cadastrados. Toque em Editar e copie da certidão de inteiro teor: endereço, matrícula, área e confrontações.</p>' : ''}</div>
    <div class="card"><h2>Venda <button class="btn sm rw" data-act="venda" data-k="${esc(q + '|' + l)}">Editar</button></h2>
    <dl class="dl">${linha('Etapa', etapaNome(infoVendaLote(q, l).etapa))}${linha('Cliente', v.cliente)}${linha('Corretor', v.corretor ? v.corretor + (v.imob ? ' · ' + v.imob : '') : (v.cliente ? 'Venda direta (sem corretor)' : ''))}${linha('Correspondente bancário', v.corresp)}${linha('Agência', v.agencia)}</dl>
    ${!v.cliente ? '<p class="muted small">Lote sem venda registrada.</p>' : ''}</div></div>`;
}
ACOES.loteInfo = d => {
  const q = d.q, l = d.l, i = loteInfo(q, l);
  const f = folha(`Dados do lote · ${q} ${l}`, `<form class="form" id="fx"><p class="muted small">Copie da certidão de inteiro teor da matrícula.</p>
    <label>${CAMPOS_LOTE[0][1]}<input name="endereco" value="${esc(i.endereco || '')}"></label>
    <div class="g3"><label>Matrícula<input name="matricula" value="${esc(i.matricula || '')}"></label><label>Cartório<input name="cartorio" value="${esc(i.cartorio || (E().obraInfo[obraDe(q)] || {}).cartorio || '')}"></label><label>Inscrição imobiliária<input name="inscricao" value="${esc(i.inscricao || '')}"></label></div>
    <div class="g2"><label>Área do terreno (m²)<input name="area" inputmode="decimal" value="${esc(i.area || '')}"></label><label>Área construída (m²)<input name="construida" inputmode="decimal" value="${esc(i.construida || '')}"></label></div>
    <label>Medidas e confrontações<textarea name="frente" rows="2" placeholder="Ex.: frente 8,00 m para a Rua A; fundos 8,00 m; lados 20,00 m">${esc(i.frente || '')}</textarea></label>
    <div class="g2"><label>Coordenadas ou link do mapa<input name="coord" value="${esc(i.coord || '')}" placeholder="-9.3912, -40.5021"></label><label>Construtor da casa<input name="construtor" list="dlConstr3" value="${esc(i.construtor || '')}" placeholder="${esc((E().obraInfo[obraDe(q)] || {}).construtor || 'Empreiteiro')}"></label></div>
    <datalist id="dlConstr3">${construtores().map(x => `<option value="${esc(x)}">`).join('')}</datalist>
    <label>Observação<input name="obs" value="${esc(i.obs || '')}"></label>
    <div class="acts"><button type="button" class="btn" data-close>Cancelar</button><button class="btn primary">Salvar</button></div></form>`);
  const F = f.el.querySelector('#fx');
  F.onsubmit = e => { e.preventDefault(); const n = {}; for (const [k] of CAMPOS_LOTE) { const v = F[k].value.trim(); if (v) n[k] = v; }
    S.lotesInfo = { ...(S.lotesInfo || {}), [q + '|' + l]: n }; gravar(doc(db, 'config', 'lotesInfo'), { [q + '|' + l]: n }, { merge: true }).catch(onErroEscrita); f.fechar(); toast('Dados do lote salvos'); refreshView(); };
};
BUSCA_EXTRA.push(q => Object.entries(S.lotesInfo || {}).filter(([, i]) => normTxt(Object.values(i).join(' ')).includes(q)).slice(0, 8).map(([k, i]) => { const [qq, l] = k.split('|'); return { titulo: `Lote ${qq} ${l}`, sub: [i.endereco, i.matricula ? 'matrícula ' + i.matricula : ''].filter(Boolean).join(' · '), href: `#/obras/${encodeURIComponent(qq)}/${encodeURIComponent(l)}` }; }));

/* Abas de cadastro na exportação em Excel (o backup .json já guarda tudo). */
function abasCadastro(wb) {
  const aba = (nome, linhas, larg) => { const ws = XLSX.utils.aoa_to_sheet(linhas); if (larg) ws['!cols'] = larg.map(w => ({ wch: w })); XLSX.utils.book_append_sheet(wb, ws, nome); };
  const e = E();
  aba('Obras', [['Obra', 'Empresa (CNPJ)', 'Cidade', 'UF', 'CEP', 'Endereço', 'Construtor', 'Matrícula', 'Cartório', 'Mapa', 'Quadras', 'Lotes'],
    ...e.obras.map(o => { const i = e.obraInfo[o] || {}; const em = empPorId(empDaObra(o)) || {}; return [o, [em.nome, em.cnpj ? fCnpj(em.cnpj) : ''].filter(Boolean).join(' · '), i.cidade || '', i.uf || '', i.cep || '', i.endereco || '', i.construtor || '', i.matricula || '', i.cartorio || '', i.mapa || '', quadrasDaObra(o).join(', '), sum(quadrasDaObra(o), q => lotesDe(q).length)]; })], [24, 34, 16, 4, 10, 36, 24, 12, 18, 20, 20, 6]);
  aba('Empresas', [['Nome', 'Tipo', 'CNPJ', 'Optante RET', 'Alíquota RET %', 'Chave PIX', 'Nome no PIX', 'Cidade no PIX', 'Obras'],
    ...e.empresas.map(x => [x.nome, x.tipo === 'matriz' ? 'Matriz' : 'Filial', x.cnpj ? fCnpj(x.cnpj) : '', x.ret ? 'Sim' : 'Não', x.ret ? x.retAliq || 4 : '', x.pix || '', x.pixNome || '', x.pixCidade || '', e.obras.filter(o => empDaObra(o) === x.id).join(', ')])], [30, 8, 20, 10, 12, 26, 20, 16, 30]);
  const sd = saldos().contas;
  aba('Contas bancárias', [['Conta', 'Banco', 'Agência', 'Número', 'Tipo', 'Empresa (CNPJ)', 'Saldo inicial', 'Saldo atual'],
    ...e.contas.map(c => { const em = empPorId(empConta(c.nome)) || {}; return [c.nome, c.banco || '', c.ag || '', c.num || '', ({ corrente: 'Conta corrente', aplicacao: 'Aplicação', caixa: 'Caixa' })[c.tipoC] || '', [em.nome, em.cnpj ? fCnpj(em.cnpj) : ''].filter(Boolean).join(' · '), +c.inicial || 0, (sd.find(x => x.nome === c.nome) || {}).saldo || 0]; })], [16, 22, 8, 14, 14, 34, 14, 14]);
  const vl = (S.vendas && S.vendas.lotes) || {};
  aba('Lotes', [['Obra', 'Quadra', 'Lote', 'Endereço (certidão)', 'Matrícula', 'Cartório', 'Inscrição', 'Área terreno m²', 'Área construída m²', 'Medidas e confrontações', 'Coordenadas', 'Construtor', 'Etapa da venda', 'Cliente', 'Telefone', 'Corretor', 'Imobiliária', 'Correspondente', 'Agência', 'Recebido', 'A receber'],
    ...e.quadras.flatMap(q => lotesDe(q).map(l => { const i = loteInfo(q, l), v = vl[q + '|' + l] || {}, iv = infoVendaLote(q, l); return [obraDe(q), q, isNaN(+l) ? l : +l, i.endereco || '', i.matricula || '', i.cartorio || '', i.inscricao || '', i.area || '', i.construida || '', i.frente || '', i.coord || '', construtorLote(q, l), etapaNome(iv.etapa), v.cliente || '', v.tel || '', v.corretor || '', v.imob || '', v.corresp || '', v.agencia || '', iv.rec || 0, iv.aberto || 0]; }))], [20, 7, 6, 34, 12, 18, 14, 10, 10, 30, 18, 22, 18, 24, 14, 20, 18, 22, 22, 12, 12]);
  const fl = (S.fornecedores && S.fornecedores.lista) || {}; const pagoF = {}; S.list.forEach(t => { if (t.forn && t.valor < 0 && pago(t)) pagoF[t.forn] = (pagoF[t.forn] || 0) - t.valor; });
  const nomesF = [...new Set([...Object.keys(fl), ...Object.keys(pagoF)])].sort((a, b) => a.localeCompare(b, 'pt-BR'));
  aba('Fornecedores', [['Fornecedor', 'CNPJ/CPF', 'Telefone', 'Categoria', 'Observação', 'Total pago'], ...nomesF.map(n => { const x = fl[n] || {}; return [n, x.cnpj || '', x.tel || '', x.cat || '', x.obs || '', r2(pagoF[n] || 0)]; })], [32, 20, 14, 16, 30, 14]);
  const cs = S.contratos || [];
  if (cs.length) {
    aba('Contratos', [['Empreiteiro', 'Objeto', 'Obra', 'Quadras', 'Lotes', 'Valor', 'Aditivos', 'Retenção %', 'Medido', 'Pago', 'Saldo', 'Início', 'Término', 'Situação', 'Banco', 'PIX'],
      ...cs.map(c => { const r = resumoContrato(c); return [c.forn, c.objeto || '', c.obra || '', (c.quadras || []).join(', '), (c.lotes || []).map(k => k.replace('|', ' ')).join(', '), +c.valor || 0, sum(c.aditivos || [], a => +a.valor || 0), c.retPct || 0, r.medido, r.pago, r.saldo, c.inicio || '', c.fim || '', c.status || '', c.banco || '', c.pix || '']; })], [28, 30, 18, 12, 24, 14, 12, 10, 14, 14, 14, 11, 11, 10, 24, 22]);
    aba('Medições', [['Empreiteiro', 'Nº', 'Data', 'Referência', 'Bruto', 'Retenção', 'Líquido', 'Físico acumulado %', 'Etapas (%)'],
      ...cs.flatMap(c => (c.medicoes || []).slice().sort((a, b) => a.n - b.n).map(m => [c.forn, m.n, m.data || '', m.ref || '', m.valor, m.ret || 0, r2(m.valor - (m.ret || 0)), m.fisico || '', m.etapas ? Object.entries(m.etapas).map(([k, p]) => `${k}: ${r2(p * 100)}%`).join('; ') : '']))], [28, 4, 11, 28, 14, 12, 14, 12, 40]);
    const pl = cs.filter(c => c.planilha);
    if (pl.length) aba('Planilhas de contrato', [['Empreiteiro', 'Item', 'Discriminação', 'Unid.', 'Quant.', 'Unitário', 'Total'], ...pl.flatMap(c => c.planilha.etapas.flatMap(e => [[c.forn, e.n, e.nome.toUpperCase(), '', '', '', e.valor], ...e.itens.map(x => [c.forn, x.item, x.desc, x.un, x.qt, x.unit, x.total])]))], [26, 6, 50, 6, 10, 12, 14]);
  }
  const rg = (S.regras && S.regras.lista) || [];
  if (rg.length) aba('Regras de classificação', [['Contém', 'Vale para', 'Pertence a', 'Destino', 'Fornecedor', 'CNPJ', 'Etapa'], ...rg.map(r => [r.txt, r.so || 'todos', r.ambito || '', r.destino || '', r.forn || '', (empPorId(r.emp) || {}).nome || '', r.etapa || ''])]);
}
