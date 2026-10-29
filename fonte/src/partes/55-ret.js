/* =========================================================
   Imposto do RET por filial: alíquota sobre as receitas recebidas no mês, DARF até o dia 20 do mês seguinte
   ========================================================= */
const RET_PARTES_4 = [['IRPJ', 1.26], ['CSLL', 0.66], ['PIS', 0.37], ['COFINS', 1.71]];
const empresasRET = () => E().empresas.filter(x => x.ret);
const idGuiaRET = (emp, m) => 'ret_' + emp + '_' + m.replace('-', '');
const vencRET = m => proximoDiaUtil(addMonths(m + '-01', 1).slice(0, 8) + '20');
function baseRET(emp, m) {
  const ts = S.list.filter(t => pago(t) && t.valor > 0 && empT(t) === emp && (t.pag || t.venc || '').slice(0, 7) === m && secaoDe(t.destino) === 'rec');
  return { ts, base: sum(ts, t => t.valor) };
}
function linhasRET(emp, meses) {
  const x = empPorId(emp); const aliq = (+x?.retAliq || 4) / 100;
  return meses.map(m => { const { ts, base } = baseRET(emp, m); const imposto = r2(base * aliq); const g = S.byId.get(idGuiaRET(emp, m)); return { m, base, n: ts.length, imposto, g, dif: g ? r2(imposto + g.valor) : imposto, aliq }; });
}
const mesesRET = () => { const h = hoje().slice(0, 7); return mesesEntre(addMonths(h + '-01', -12).slice(0, 7), h).reverse(); };
function guiaRET(emp, l) {
  const x = empPorId(emp); const o = E().obras.find(ob => empDaObra(ob) === emp) || E().obras[0];
  const obs = l.aliq === 0.04 ? 'RET 4% = ' + RET_PARTES_4.map(([n, p]) => `${n} ${brl2(r2(l.base * p / 100))}`).join(', ') : `RET ${l.aliq * 100}%`;
  const t = { ...(l.g || { juros: 0, nf: 0, nfNum: '', status: 'aberto', criado: Date.now() }), id: idGuiaRET(emp, l.m), item: `RET ${MES3[+l.m.slice(5, 7) - 1]}/${l.m.slice(2, 4)} - ${x?.nome || ''}`, valor: -l.imposto, local: o, quadra: '-', lote: '-', destino: todosDestinos().includes('Impostos') ? 'Impostos' : 'Administrativo', venc: vencRET(l.m), forn: 'Receita Federal (DARF)', obs: `${obs}. Base: ${brl2(l.base)} recebidos em ${mesNome(l.m)}.`, emp };
  return t;
}
const pendentesRET = () => { const h = hoje().slice(0, 7); const out = []; for (const x of empresasRET()) for (const l of linhasRET(x.id, mesesRET())) if (l.m < h && l.imposto >= 0.01 && (!l.g || (Math.abs(l.dif) >= 0.01 && !pago(l.g)))) out.push([x.id, l]); return out; };
VIEWS.ret = {
  titulo: 'Impostos (RET)',
  refresh(el) {
    const es = empresasRET(); const h = hoje().slice(0, 7);
    if (!es.length) { el.innerHTML = `<div class="card"><h2>Imposto do RET</h2><p class="muted">Nenhuma empresa está marcada como optante pelo RET. Em Cadastros, Empresas (CNPJ), edite a filial, marque "Optante pelo RET" e escolha a alíquota (4% na regra geral, 1% no MCMV Faixa 1).</p><a class="btn" href="#/cadastros">Ir para Cadastros</a></div>`; return; }
    const pend = pendentesRET();
    el.innerHTML = `${pend.length ? `<div class="alerta-caixa"><span>${pend.length} guia(s) do RET para lançar ou atualizar, no total de ${brl2(sum(pend, ([, l]) => l.dif))}.</span><button class="btn sm primary rw" data-act="retTodas">Lançar todas</button></div>` : ''}
      ${es.map(x => { const ls = linhasRET(x.id, mesesRET()); return `<div class="card"><h2>${esc(x.nome)} <span class="muted small">${esc(fCnpj(x.cnpj))} · alíquota ${(+x.retAliq || 4)}%</span></h2>
        <div class="tbl-wrap"><table class="tbl"><thead><tr><th class="l">Mês da receita</th><th>Receitas recebidas</th><th>Imposto</th><th class="l">Vencimento</th><th class="l">Guia</th><th></th></tr></thead><tbody>
        ${ls.map(l => `<tr><td class="l">${mesNome(l.m)}${l.m === h ? ' <span class="muted small">(em andamento)</span>' : ''}</td><td>${brl2(l.base)}<br><span class="muted small">${l.n} recebimento(s)</span></td><td><b>${brl2(l.imposto)}</b></td><td class="l">${dBR(vencRET(l.m))}</td><td class="l">${l.g ? `<span class="chip ${pago(l.g) ? 'ok' : ''}">${pago(l.g) ? 'paga' : 'lançada'}</span>${Math.abs(l.dif) >= 0.01 ? ` <span class="neg small">diferença ${brl2(l.dif)}</span>` : ''}` : l.imposto ? '<span class="chip warn">não lançada</span>' : '—'}</td><td>${l.imposto >= 0.01 && l.m < h && (!l.g || (Math.abs(l.dif) >= 0.01 && !pago(l.g))) ? `<button class="btn sm rw" data-act="retLanc" data-e="${esc(x.id)}" data-m="${l.m}">${l.g ? 'Atualizar' : 'Lançar guia'}</button>` : l.g ? `<button class="btn sm ghost" data-act="abrir" data-id="${esc(l.g.id)}">Ver</button>` : ''}</td></tr>`).join('')}</tbody></table></div></div>`; }).join('')}
      <p class="muted small">Base: receitas recebidas no mês pelo CNPJ (vendas, repasses e rendimentos), sem aportes. Na alíquota de 4%: IRPJ 1,26%, CSLL 0,66%, PIS 0,37% e COFINS 1,71%. A guia (DARF) vence no dia 20 do mês seguinte; o app passa para o próximo dia útil. Confira com a contabilidade, que pode deduzir devoluções e descontos.</p>`;
  }
};
ACOES.retLanc = d => { const l = linhasRET(d.e, [d.m])[0]; salvarTitulos([guiaRET(d.e, l)]); toast('Guia do RET lançada em Contas a pagar'); refreshView(); };
ACOES.retTodas = () => { const p = pendentesRET(); salvarTitulos(p.map(([e, l]) => guiaRET(e, l))); toast(p.length + ' guia(s) do RET lançada(s)'); refreshView(); };
