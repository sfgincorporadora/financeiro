/* =========================================================
   Arrumar lançamentos: classificação em lote com sugestões
   e painel de alertas do Início
   ========================================================= */
const chaveGrupo = t => normTxt(baseItem(t)).replace(/\(\d+\/\d+\)/g, '').replace(/\b\d+[a-z]?\b/g, ' ').replace(/\s+/g, ' ').trim() || '(sem descrição)';
const TIPOS_ARRUMAR = [['forn', 'Sem fornecedor'], ['etapa', 'Sem etapa da obra'], ['cnpj', 'Sem CNPJ'], ['anexo', 'Pagos sem nota ou anexo']];
let arrTipo = 'empresa';
function candidatosArrumar(tipo) {
  if (tipo === 'forn') return S.list.filter(t => t.valor < 0 && !t.forn);
  if (tipo === 'etapa') return S.list.filter(t => t.valor < 0 && !t.etapaObra && secaoDe(t.destino) === 'op' && t.quadra && t.quadra !== '-');
  if (tipo === 'cnpj') return variasEmpresas() ? S.list.filter(t => !empDef(t)) : [];
  return S.list.filter(t => pago(t) && t.valor < 0 && !t.nf && !anexosDe(t).length && ['op', 'nonop'].includes(secaoDe(t.destino)) && t.destino !== 'Impostos' && (t.pag || t.venc) >= addDays(hoje(), -120));
}
function sugestao(tipo, chave, ts) {
  if (tipo === 'cnpj') { const c = {}; ts.forEach(t => { const e = empDaObra(obraT(t)); c[e] = (c[e] || 0) + 1; }); return Object.entries(c).sort((a, b) => b[1] - a[1])[0]?.[0] || ''; }
  if (tipo === 'anexo') return '';
  const campo = tipo === 'forn' ? 'forn' : 'etapaObra';
  const freq = {}; for (const t of S.list) if (t[campo] && chaveGrupo(t) === chave) freq[t[campo]] = (freq[t[campo]] || 0) + 1;
  const top = Object.entries(freq).sort((a, b) => b[1] - a[1])[0]; if (top) return top[0];
  if (tipo === 'forn') { const nomes = [...new Set([...S.list.map(t => t.forn).filter(Boolean), ...Object.keys((S.fornecedores && S.fornecedores.lista) || {})])].sort((a, b) => b.length - a.length); const ck = ' ' + chave + ' '; for (const n of nomes) { const k = normTxt(n); if (k.length > 3 && ck.includes(' ' + k + ' ')) return n; const p = k.split(' ')[0]; if (p.length > 4 && ck.includes(' ' + p + ' ')) return n; } }
  return '';
}
const VIEWS_ARRUMAR_ANTIGO = {
  refresh(el, abas) {
    const tipos = TIPOS_ARRUMAR.filter(([k]) => k !== 'cnpj' || variasEmpresas());
    const cont = Object.fromEntries(tipos.map(([k]) => [k, candidatosArrumar(k).length]));
    const ts = candidatosArrumar(arrTipo);
    const gs = {}; for (const t of ts) (gs[chaveGrupo(t)] = gs[chaveGrupo(t)] || []).push(t);
    const lista = Object.entries(gs).map(([k, xs]) => ({ k, xs, n: xs.length, tot: sum(xs, t => t.valor), ex: xs[0], sug: sugestao(arrTipo, k, xs) })).sort((a, b) => b.n - a.n || a.tot - b.tot).slice(0, 150);
    const opcoes = arrTipo === 'forn' ? [...new Set([...S.list.map(t => t.forn).filter(Boolean), ...Object.keys((S.fornecedores && S.fornecedores.lista) || {})])].sort() : arrTipo === 'etapa' ? etapasOrc() : [];
    el.innerHTML = `${abas}
      <div class="card"><h2>${esc(TIPOS_ARRUMAR.find(x => x[0] === arrTipo)[1])} <span class="muted small">${ts.length.toLocaleString('pt-BR')} lançamento(s) em ${Object.keys(gs).length} grupo(s) parecido(s)</span>${arrTipo !== 'anexo' && lista.some(g => g.sug) ? '<button class="btn sm primary rw" id="aTodas">Aplicar todas as sugestões</button>' : ''}</h2>
      <p class="muted small">${arrTipo === 'anexo' ? 'Pagamentos dos últimos 4 meses sem número de NF e sem anexo. Abra para anexar a nota ou o comprovante.' : 'Lançamentos com descrição parecida ficam juntos. Confira a sugestão (ou escolha outra) e aplique ao grupo inteiro.'}</p>
      ${opcoes.length ? `<datalist id="dlArr">${opcoes.map(o => `<option value="${esc(o)}">`).join('')}</datalist>` : ''}
      <div class="arr-lista">${lista.map((g, i) => arrTipo === 'anexo' ? g.xs.slice(0, 3).map(t => linhaTitulo(t, { mostrarPago: true })).join('') : `<div class="arr-g"><div class="arr-txt"><b>${esc(g.ex.item)}</b><small>${g.n} lançamento(s) · ${brl2(g.tot)} · ${esc([...new Set(g.xs.map(t => t.quadra))].slice(0, 4).join(', '))}</small></div>
        ${arrTipo === 'cnpj' ? `<select data-g="${i}">${opts(empOpts(), g.sug, 'Escolha…')}</select>` : arrTipo === 'etapa' ? `<select data-g="${i}">${opts(etapasOrc(), g.sug, 'Escolha…')}</select>` : `<input data-g="${i}" list="dlArr" value="${esc(g.sug)}" placeholder="Fornecedor">`}<button class="btn sm rw" data-ap="${i}">Aplicar</button></div>`).join('') || '<p class="muted">Nada para arrumar aqui.</p>'}</div></div>`;
    const campo = arrTipo === 'forn' ? 'forn' : arrTipo === 'etapa' ? 'etapaObra' : 'emp';
    const aplicar = (g, v) => { if (!v) return 0; salvarTitulos(g.xs.map(t => ({ ...t, [campo]: v }))); return g.xs.length; };
    $$('[data-ap]', el).forEach(b => b.onclick = () => { const g = lista[+b.dataset.ap]; const v = $(`[data-g="${b.dataset.ap}"]`, el).value.trim(); const n = aplicar(g, v); if (n) toast(n + ' lançamento(s) atualizados'); });
    const at = $('#aTodas', el); if (at) at.onclick = async () => { const gsS = lista.filter(g => g.sug); if (!(await confirmar(`Aplicar a sugestão em ${gsS.length} grupo(s), ${sum(gsS, g => g.n)} lançamento(s)?`, 'Aplicar', false))) return; let n = 0; gsS.forEach(g => { n += aplicar(g, g.sug); }); toast(n + ' lançamento(s) atualizados'); };
  }
};
ACOES.arrTipo = d => { arrTipo = d.v; refreshView(); };

/* ---------- Painel de alertas (Início) ---------- */
function listaAlertas() {
  const h = hoje(), A = [];
  const add = (nivel, txt, href, acao) => A.push({ nivel, txt, href, acao });
  const venc = S.list.filter(t => !pago(t) && t.valor < 0 && t.venc < h);
  if (venc.length) add('alto', `${venc.length} conta(s) a pagar vencida(s), ${brl(-sum(venc, t => t.valor))}`, '#/pagar');
  const pj = projecao(); const neg = pj.dias.find(x => x.v < 0);
  if (neg) add('alto', `O caixa fica negativo em ${dBR(neg.d)} (menor saldo ${brl(pj.min.v)} em ${dBR(pj.min.d)})`, '#/fluxo');
  const recV = S.list.filter(t => !pago(t) && t.valor > 0 && t.venc < addDays(h, -5));
  if (recV.length) add('medio', `${recV.length} recebimento(s) atrasado(s) há mais de 5 dias, ${brl(sum(recV, t => t.valor))}`, '#/receber');
  if (typeof pendentesRET === 'function') { const p = pendentesRET(); if (p.length) add('alto', `${p.length} guia(s) do RET para lançar`, '#/ret'); }
  if (typeof backupAtrasado === 'function' && backupAtrasado()) add('medio', 'Backup automático atrasado: conecte o Google e faça um backup', '#/dados');
  if (agendaAtiva() && (!S.agenda.ultima || Date.now() - S.agenda.ultima > 864e5)) add('baixo', 'Avisos da Google Agenda desatualizados', null, 'agSync');
  let docsV = 0; const qs = E().quadras.filter(q => lotesDe(q).length); const nd = docsCk().length;
  if (((S.checklist && S.checklist.validade) || null)) for (const q of qs) for (const l of lotesDe(q)) for (let i = 0; i < nd; i++) if (estadoDoc(q, l, i).venc) docsV++;
  if (docsV) add('medio', `${docsV} documento(s) do checklist com validade vencida`, '#/checklist');
  if (S.fluxoLotes && Object.keys(S.fluxoLotes).length) { let atr = 0; for (const k of Object.keys(S.fluxoLotes)) { const [q, l] = k.split('|'); const e = estadoFx(q, l); for (const id in e) if (e[id].atraso && !e[id].ok) atr++; } if (atr) add('medio', `${atr} etapa(s) de lotes acima do prazo`, '#/etapas'); }
  const vp = qs.flatMap(q => lotesDe(q).map(l => infoVendaLote(q, l))).filter(x => x.v && x.v.data && !['disp', 'recebido'].includes(x.etapa) && x.v.data < addDays(h, -30));
  if (vp.length) add('baixo', `${vp.length} venda(s) parada(s) na mesma etapa há mais de 30 dias`, '#/vendas');
  if (typeof comprasAprovacao === 'function' && comprasAprovacao().length) add('medio', `${comprasAprovacao().length} compra(s) aguardando sua aprovação`, '#/compras');
  if (S.contratos) { const c = S.contratos.filter(c => c.status !== 'encerrado').map(c => [c, resumoContrato(c)]).filter(([, r]) => r.fisico !== null && r.pct !== null && r.pct * 100 > r.fisico + 10); if (c.length) add('medio', `${c.length} contrato(s) com mais pago do que o avanço físico medido`, '#/contratos'); }
  if (variasEmpresas()) { const d = S.list.filter(empDiverge).length; if (d) add('baixo', `${d} lançamento(s) com CNPJ diferente da obra`, '#/empresas'); }
  const semNota = candidatosArrumar('anexo').filter(t => t.criadoPor && (t.pag || t.venc) >= addDays(h, -45)).length;
  const emLotes = typeof candidatosRevisar === 'function' ? gruposJuntar(candidatosRevisar('empresa')).length : 0;
  if (emLotes) add('medio', `${emLotes} despesa(s) da empresa lançada(s) dividida(s) em lotes: revise e junte`, '#/arrumar', 'arrEmpresa');
  if (semNota) add('baixo', `${semNota} pagamento(s) lançado(s) no app nos últimos 45 dias sem nota fiscal ou anexo`, '#/arrumar', 'arrAnexo');
  return A;
}
ACOES.arrAnexo = () => { arrTipo = 'anexo'; location.hash = '#/arrumar'; };
ACOES.arrEmpresa = () => { arrTipo = 'empresa'; location.hash = '#/arrumar'; };
function painelAlertas() {
  const A = listaAlertas(); if (!A.length) return `<div class="alertas ok"><span class="dot-ok"></span> Tudo em dia: nenhuma pendência importante.</div>`;
  const ord = { alto: 0, medio: 1, baixo: 2 }; A.sort((a, b) => ord[a.nivel] - ord[b.nivel]);
  const vis = A.slice(0, 4), resto = A.slice(4);
  const item = a => a.acao ? `<button class="al al-${a.nivel}" data-act="${a.acao}"><i></i><span>${esc(a.txt)}</span>${ic('voltar', 14)}</button>` : `<a class="al al-${a.nivel}" href="${a.href}"><i></i><span>${esc(a.txt)}</span>${ic('voltar', 14)}</a>`;
  return `<section class="alertas"><h2>Atenção <span class="muted small">${A.length} ponto(s)</span></h2>${vis.map(item).join('')}${resto.length ? `<details><summary>Ver mais ${resto.length}</summary>${resto.map(item).join('')}</details>` : ''}</section>`;
}
