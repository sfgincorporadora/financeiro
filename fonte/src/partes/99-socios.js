/* =========================================================
   Sócios e participações: aportes e retiradas por sócio,
   acordos de divisão do resultado por quadra e posição de cada um.
   config/socios = { lista: [{id, nome, apelidos}], acordos: [{id, nome, escopo: 'quadras'|'demais', quadras: [], partes: {socioId: %}}], adm: bool }
   ========================================================= */
usarConfig('socios');
const TIPOS_FORA_SOCIOS = ['intragrupo', 'transferencia', 'terceiro', 'conversivel'];
const cfgSocios = () => S.socios || {};
const listaSocios = () => cfgSocios().lista || [];
const acordos = () => cfgSocios().acordos || [];
function salvarSocios(n) { S.socios = { ...cfgSocios(), ...n }; gravar(doc(db, 'config', 'socios'), S.socios).catch(onErroEscrita); refreshView(); }
const ehAporteSocio = t => secaoDe(t.destino) === 'aporte' && !TIPOS_FORA_SOCIOS.includes(t.aporteTipo);
/* Nome que aparece depois de "Aporte" na descrição (ex.: "Aporte Eduardo (3/10)"). */
const nomeNoAporte = t => { const m = normTxt(t.item).match(/^(?:aporte|retirada|devolucao|distribuicao) (?:de |do |da )?([a-z]+)/); return m ? m[1] : ''; };
function socioDe(t) {
  if (t.socio) return listaSocios().find(s => s.id === t.socio) ? t.socio : '';
  const n = nomeNoAporte(t), f = normTxt(t.forn);
  for (const s of listaSocios()) { const nomes = [s.nome, ...(s.apelidos || '').split(',')].map(x => normTxt(x)).filter(Boolean);
    if (n && nomes.some(x => x.split(' ')[0] === n)) return s.id; if (f && nomes.includes(f)) return s.id; }
  return '';
}
function detectarSocios() { const c = {}; for (const t of S.list) if (ehAporteSocio(t)) { const n = nomeNoAporte(t); if (n) c[n] = (c[n] || 0) + 1; } return Object.entries(c).sort((a, b) => b[1] - a[1]).map(([n]) => n.charAt(0).toUpperCase() + n.slice(1)); }
/* Resultado por lote: lançamentos do lote + parte igual do custo geral da quadra. */
function resultadoLotes(adm) {
  const o = {}; const ger = {};
  for (const t of S.list) {
    if (!t.quadra || t.quadra === '-' || !E().quadras.includes(t.quadra) || secaoDe(t.destino) === 'aporte') continue;
    const p = pago(t), v = t.valor + (p ? (t.juros || 0) : 0);
    const alvo = t.lote && t.lote !== '-' ? (o[t.quadra + '|' + t.lote] = o[t.quadra + '|' + t.lote] || { rec: 0, custo: 0, real: 0, proj: 0 }) : (ger[t.quadra] = ger[t.quadra] || { rec: 0, custo: 0, real: 0, proj: 0 });
    if (p) { alvo.real += v; if (v > 0) alvo.rec += v; else alvo.custo += v; } alvo.proj += v;
  }
  const out = {};
  for (const q of E().quadras) { const ls = lotesDe(q); const g = ger[q] || { rec: 0, custo: 0, real: 0, proj: 0 };
    for (const l of ls) { const x = o[q + '|' + l] || { rec: 0, custo: 0, real: 0, proj: 0 }; const ad = adm && typeof admLote === 'function' ? admLote(q, l) : 0;
      out[q + '|' + l] = { q, l, rec: x.rec + g.rec / ls.length, custo: x.custo + g.custo / ls.length - ad, real: x.real + g.real / ls.length - ad, proj: x.proj + g.proj / ls.length - ad }; } }
  return out;
}
/* Lotes de cada acordo: lotes escolhidos valem mais que quadras; "demais" pega o que sobrou. */
function lotesDosAcordos() {
  const dono = {}; const as = acordos();
  for (const a of as) if (a.escopo === 'quadras') for (const q of a.quadras || []) for (const l of lotesDe(q)) dono[q + '|' + l] = a.id;
  for (const a of as) if (a.escopo === 'lotes') for (const k of a.lotes || []) dono[k] = a.id;
  const dem = as.find(a => a.escopo === 'demais');
  if (dem) for (const q of E().quadras) for (const l of lotesDe(q)) if (!dono[q + '|' + l]) dono[q + '|' + l] = dem.id;
  return dono;
}
/* Investidor com cota no acordo: a cota dele sai do resultado antes da divisão principal
   e é repartida pela divisão própria (ex.: Antônio 60% / Luis 40%); o restante segue a divisão do acordo. */
const invDe = a => a.inv && a.inv.socio && (a.inv.modo === 'liquidado' || +a.inv.cota) ? a.inv : null;
/* Investidor liquidado: a cota é um valor fixo = o que ele retirou além do aporte ÷ a % dele na cota. */
/* Data da liquidação: a informada, ou a da última retirada do investidor. Aportes depois dela são investimento novo. */
function dataLiquidacao(inv) { if (inv.ate) return inv.ate; const rs = S.list.filter(t => ehAporteSocio(t) && socioDe(t) === inv.socio && pago(t) && t.valor < 0).map(t => t.pag || t.venc).sort(); return rs.pop() || hoje(); }
function valorCotaLiquidada(inv) { const ate = dataLiquidacao(inv); const ts = S.list.filter(t => ehAporteSocio(t) && socioDe(t) === inv.socio && pago(t) && (t.pag || t.venc) <= ate); const ganho = -sum(ts, t => t.valor); const pi = (+((inv.div || {})[inv.socio]) || 0) / 100; return pi ? ganho / pi : 0; }
/* Parte do sócio no resultado de um acordo. */
function parteResultado(a, sid, res) { const inv = invDe(a); const p = (+(a.partes || {})[sid] || 0) / 100, d = inv ? (+(inv.div || {})[sid] || 0) / 100 : 0;
  if (inv && inv.modo === 'liquidado') { const V = valorCotaLiquidada(inv); return (res - V) * p + V * d; }
  return res * pctEfetivo(a, sid); }
function pctEfetivo(a, sid) { const inv = invDe(a); const c = inv ? +inv.cota / 100 : 0;
  return (1 - c) * ((+(a.partes || {})[sid] || 0) / 100) + c * ((+((inv && inv.div) || {})[sid] || 0) / 100); }
function calcSocios() {
  const adm = !!cfgSocios().adm; const RL = resultadoLotes(adm); const dono = lotesDosAcordos();
  const res = acordos().map(a => { const ks = Object.keys(dono).filter(k => dono[k] === a.id); const qs = [...new Set(ks.map(k => k.split('|')[0]))];
    const rs = qs.map(q => { const xs = ks.filter(k => k.startsWith(q + '|')).map(k => RL[k]).filter(Boolean); const fin = xs.filter(x => infoVendaLote(x.q, x.l).etapa === 'recebido'); return { q, n: xs.length, total: lotesDe(q).length, nFin: fin.length, rec: sum(xs, x => x.rec), custo: sum(xs, x => x.custo), real: sum(xs, x => x.real), proj: sum(xs, x => x.proj), conc: sum(fin, x => x.real) }; });
    return { a, qs, ks, rs, real: sum(rs, r => r.real), proj: sum(rs, r => r.proj), conc: sum(rs, r => r.conc) }; });
  const so = listaSocios().map(s => {
    const ts = S.list.filter(t => ehAporteSocio(t) && socioDe(t) === s.id);
    const ent = sum(ts.filter(t => pago(t) && t.valor > 0), t => t.valor), sai = sum(ts.filter(t => pago(t) && t.valor < 0), t => t.valor);
    const prev = sum(ts.filter(t => !pago(t)), t => t.valor);
    const conc = sum(res, r => parteResultado(r.a, s.id, r.conc));
    const real = sum(res, r => parteResultado(r.a, s.id, r.real)), proj = sum(res, r => parteResultado(r.a, s.id, r.proj));
    return { s, ts, ent, sai, liq: ent + sai, prev, real, proj, conc, pos: ent + sai + real, posHoje: ent + sai + conc, posProj: ent + sai + prev + proj };
  });
  const totLiq = sum(so.filter(x => x.liq > 0), x => x.liq);
  so.forEach(x => { x.pctCap = totLiq && x.liq > 0 ? x.liq / totLiq : 0; });
  /* Participação no patrimônio: saldo de cada sócio sobre a soma dos saldos positivos (quem está negativo deve à sociedade e fica com 0%). */
  const totPos = sum(so.filter(x => x.posHoje > 0), x => x.posHoje), totProj = sum(so.filter(x => x.posProj > 0), x => x.posProj);
  so.forEach(x => { x.partHoje = totPos && x.posHoje > 0 ? x.posHoje / totPos : 0; x.partProj = totProj && x.posProj > 0 ? x.posProj / totProj : 0; });
  const semSocio = S.list.filter(t => ehAporteSocio(t) && !socioDe(t));
  return { res, so, semSocio, adm };
}
VIEWS.socios = {
  titulo: 'Sócios e participações',
  refresh(el) {
    if (!listaSocios().length) { const det = detectarSocios();
      el.innerHTML = `<div class="card"><h2>Sócios e participações</h2><p>Esta tela soma os aportes e retiradas de cada sócio e divide o resultado das quadras conforme os acordos entre vocês.</p>
        <p class="muted small">${det.length ? `Encontrei aportes com estes nomes na descrição: <b>${det.map(esc).join(', ')}</b>.` : 'Não encontrei aportes com nome de sócio na descrição (ex.: "Aporte Fulano").'}</p>
        <div class="acts left"><button class="btn primary rw" id="sCriar">${det.length ? 'Cadastrar estes sócios' : 'Cadastrar sócio'}</button></div></div>`;
      $('#sCriar', el).onclick = () => det.length ? salvarSocios({ lista: det.map(n => ({ id: uid(), nome: n, apelidos: '' })), acordos: [] }) : ACOES.socioEd({}); return; }
    const C = calcSocios(); const so = C.so;
    const cor = i => ['#1F3A6B', '#2E8B6A', '#C08A2E', '#8A5BB5', '#C2554A'][i % 5];
    const barra = (k, tit, sub) => `<div class="part-linha"><div class="part-tit"><b>${tit}</b><span class="muted small">${sub}</span></div><div class="part-barra">${so.map((x, i) => x[k] > 0.0005 ? `<span style="width:${(x[k] * 100).toFixed(2)}%;background:${cor(i)}" title="${esc(x.s.nome)} ${fPct.format(x[k])}">${x[k] > 0.08 ? `${esc(x.s.nome)} ${fPct.format(x[k])}` : ''}</span>` : '').join('')}</div></div>`;
    const quadroPart = `<div class="card"><h2>Participação de cada sócio <span class="hint">sobre o patrimônio de cada um</span></h2>
      ${barra('pctCap', 'Pelo capital', 'só o que cada um colocou menos o que retirou')}
      ${barra('partHoje', 'Hoje', 'capital + lucro dos lotes já concluídos (obras em andamento contam pelo custo)')}
      ${barra('partProj', 'Projetada', 'capital + resultado quando as obras terminarem (a receber e a pagar)')}
      <div class="tbl-wrap"><table class="tbl"><thead><tr><th class="l">Sócio</th><th>Pelo capital</th><th>Hoje</th><th>Projetada</th><th>Variação hoje → projetada</th></tr></thead><tbody>
      ${so.map((x, i) => `<tr><td class="l"><span class="dot-s" style="background:${cor(i)}"></span><b>${esc(x.s.nome)}</b></td><td>${fPct.format(x.pctCap)}</td><td><b>${fPct.format(x.partHoje)}</b> <span class="muted small">${brl(x.posHoje)}</span></td><td><b>${fPct.format(x.partProj)}</b> <span class="muted small">${brl(x.posProj)}</span></td><td class="${cls(x.partProj - x.partHoje)}">${((x.partProj - x.partHoje) * 100).toLocaleString('pt-BR', { maximumFractionDigits: 1, signDisplay: 'exceptZero' })} p.p.</td></tr>`).join('')}</tbody></table></div>
      <p class="muted small">Participação = patrimônio do sócio ÷ soma dos patrimônios positivos. <b>Hoje</b>: o que ele colocou − o que retirou + a parte dele no lucro dos lotes com repasse recebido; o dinheiro gasto em obras em andamento é investimento (vale o custo), não prejuízo. <b>Projetada</b>: soma também o resultado previsto de todos os lotes (o que falta receber e pagar). Patrimônio negativo (retirou mais do que a parte dele) conta como 0%. A participação muda porque cada acordo divide o lucro numa proporção diferente do capital de cada um.</p></div>`;
    el.innerHTML = `<div class="grid kpis k4">${so.slice(0, 4).map(x => kpi(esc(x.s.nome), fPct.format(x.partHoje) + ' → ' + fPct.format(x.partProj), `hoje ${brl(x.posHoje)} · projetado ${brl(x.posProj)}`, '')).join('')}</div>
      ${!acordos().length ? `<div class="alert">Nenhum acordo de divisão cadastrado: por isso o resultado das obras ainda não entra na conta e a participação mostra só o capital. Toque em <b>Novo acordo</b> (mais abaixo) para dizer como o resultado de cada grupo de quadras é dividido. <button class="btn sm primary rw" data-act="acordoEd">Criar primeiro acordo</button></div>` : ''}
      ${quadroPart}
      <div class="card"><h2>Posição de cada sócio <span class="hint">capital líquido + parte do resultado</span><button class="btn sm rw" data-act="socioEd">Adicionar sócio</button></h2>
      <div class="tbl-wrap"><table class="tbl"><thead><tr><th class="l">Sócio</th><th>Aportou</th><th>Retirou</th><th>Capital líquido</th><th>% do capital</th><th>Resultado realizado</th><th>Saldo realizado</th><th>Resultado projetado</th><th>Saldo projetado</th><th></th></tr></thead><tbody>
      ${so.map((x, i) => `<tr><td class="l"><span class="dot-s" style="background:${cor(i)}"></span><b>${esc(x.s.nome)}</b></td><td class="pos">${brl(x.ent)}</td><td class="neg">${brl(x.sai)}</td><td class="${cls(x.liq)}"><b>${brl(x.liq)}</b></td><td>${x.pctCap ? fPct.format(x.pctCap) : '—'}</td><td class="${cls(x.real)}">${brl(x.real)}</td><td class="${cls(x.pos)}"><b>${brl(x.pos)}</b></td><td class="${cls(x.proj)}">${brl(x.proj)}</td><td class="${cls(x.posProj)}">${brl(x.posProj)}</td><td><button class="btn sm" data-act="socioVer" data-id="${esc(x.s.id)}">Movimentos</button> <button class="btn sm link rw" data-act="socioEd" data-id="${esc(x.s.id)}">Editar</button></td></tr>`).join('')}
      <tr class="tot"><td class="l">Total</td><td>${brl(sum(so, x => x.ent))}</td><td>${brl(sum(so, x => x.sai))}</td><td>${brl(sum(so, x => x.liq))}</td><td></td><td>${brl(sum(so, x => x.real))}</td><td>${brl(sum(so, x => x.pos))}</td><td>${brl(sum(so, x => x.proj))}</td><td>${brl(sum(so, x => x.posProj))}</td><td></td></tr></tbody></table></div>
      <p class="muted small">Obras em andamento têm resultado realizado negativo (o custo já foi pago e a venda ainda não entrou); por isso o saldo projetado é o que mostra a divisão final. <b>Saldo realizado</b> = o que o sócio colocou − o que retirou + a parte dele no resultado já realizado (recebido − pago) das quadras dos acordos. Saldo positivo: a empresa deve ao sócio; negativo: o sócio retirou mais do que a parte dele. <b>Projetado</b> inclui o que ainda está a receber e a pagar.</p>
      <label class="chk"><input type="checkbox" id="sAdm" ${C.adm ? 'checked' : ''}> Descontar do resultado das quadras a administração rateada (despesas da empresa)</label></div>
      <div class="card"><h2>Acordos de divisão do resultado <button class="btn sm primary rw" data-act="acordoEd">Novo acordo</button></h2>
      ${C.res.length ? `<div class="tbl-wrap"><table class="tbl"><thead><tr><th class="l">Acordo</th><th class="l">Quadras</th><th class="l">Divisão</th><th>Resultado realizado</th><th>Resultado projetado</th><th></th></tr></thead><tbody>
      ${C.res.map(r => `<tr><td class="l"><b>${esc(r.a.nome || 'Acordo')}</b></td><td class="l wrap">${r.rs.length ? esc(r.rs.map(x => x.n < x.total ? `${x.q} (${x.n} de ${x.total} lotes)` : x.q).join(', ')) : '<span style="color:var(--warn)">defina as quadras ou lotes</span>'}${r.a.escopo === 'demais' ? ' <span class="muted small">(demais)</span>' : ''}</td><td class="l wrap">${(() => { const nm = id => esc((listaSocios().find(s => s.id === id) || {}).nome || '?'); const fx = o => Object.entries(o || {}).filter(([, p]) => +p).map(([id, p]) => `${nm(id)} ${p}%`).join(' · '); const inv = invDe(r.a);
        return !inv ? fx(r.a.partes) : inv.modo === 'liquidado' ? `<b>Cota de ${nm(inv.socio)} (liquidada em ${dBR(dataLiquidacao(inv))}): ${brl(valorCotaLiquidada(inv))}</b> → ${fx(inv.div)}<br>Restante do resultado → ${fx(r.a.partes)}` : `<b>Cota de ${nm(inv.socio)}: ${String(inv.cota).replace('.', ',')}%</b> do resultado → ${fx(inv.div)}<br>Restante (${(100 - inv.cota).toLocaleString('pt-BR', { maximumFractionDigits: 2 })}%) → ${fx(r.a.partes)}`; })()}</td><td class="${cls(r.real)}">${brl(r.real)}</td><td class="${cls(r.proj)}">${brl(r.proj)}</td><td><button class="btn sm link rw" data-act="acordoEd" data-id="${esc(r.a.id)}">Editar</button></td></tr>`).join('')}</tbody></table></div>
      <details class="mt"><summary class="small">Resultado por quadra</summary><div class="tbl-wrap"><table class="tbl"><thead><tr><th class="l">Quadra</th><th class="l">Acordo</th><th>Lotes</th><th>Recebido</th><th>Pago</th><th>Realizado</th><th>Projetado</th></tr></thead><tbody>${C.res.flatMap(r => r.rs.map(q => `<tr><td class="l">${esc(q.q)}</td><td class="l">${esc(r.a.nome || '')}</td><td>${q.n < q.total ? q.n + ' de ' + q.total : q.n}</td><td class="pos">${brl(q.rec)}</td><td class="neg">${brl(q.custo)}</td><td class="${cls(q.real)}">${brl(q.real)}</td><td class="${cls(q.proj)}">${brl(q.proj)}</td></tr>`)).join('')}</tbody></table></div></details>`
      : '<p class="muted">Nenhum acordo ainda. Crie um acordo por grupo de quadras com a divisão entre os sócios (ex.: quadras K1, Q1, U1 e Y1: 30% / 70%; demais quadras: 50% / 50%).</p>'}</div>
      ${C.semSocio.length ? `<div class="card"><h2>Aportes sem sócio identificado <span class="muted small">${C.semSocio.length} lançamento(s), ${brl(sum(C.semSocio, t => t.valor))}</span></h2><p class="muted small">Escolha de qual sócio é cada grupo. O app reconhece sozinho descrições como "Aporte Eduardo" ou o nome do sócio no campo fornecedor.</p>
        <div class="arr-lista">${Object.entries(C.semSocio.reduce((o, t) => { (o[chaveGrupo(t)] = o[chaveGrupo(t)] || []).push(t); return o; }, {})).slice(0, 60).map(([k, xs], i) => `<div class="arr-g"><div class="arr-txt"><b>${esc(xs[0].item)}</b><small>${xs.length} lançamento(s) · ${brl2(sum(xs, t => t.valor))} · ${esc(xs[0].forn || '')}</small></div><select data-sg="${i}">${opts(listaSocios().map(s => [s.id, s.nome]), '', 'Sócio…')}</select><button class="btn sm rw" data-sa="${i}" data-k="${esc(k)}">Aplicar</button></div>`).join('')}</div></div>` : ''}`;
    $('#sAdm', el).onchange = e => salvarSocios({ adm: e.target.checked });
    const grupos = Object.values(C.semSocio.reduce((o, t) => { (o[chaveGrupo(t)] = o[chaveGrupo(t)] || []).push(t); return o; }, {}));
    $$('[data-sa]', el).forEach(b => b.onclick = () => { const v = $(`[data-sg="${b.dataset.sa}"]`, el).value; if (!v) return; salvarTitulos(grupos[+b.dataset.sa].map(t => ({ ...t, socio: v }))); toast('Sócio definido'); });
  }
};
ACOES.socioEd = d => {
  const s = listaSocios().find(x => x.id === d.id);
  const f = folha(s ? 'Sócio ' + s.nome : 'Novo sócio', `<form class="form" id="fx"><label>Nome<input name="nome" required value="${esc(s?.nome || '')}"></label>
    <label>Outros nomes <span class="muted">(separados por vírgula; ajudam a reconhecer os aportes)</span><input name="ap" value="${esc(s?.apelidos || '')}" placeholder="Ex.: Antonio Lacerda, Lacerda"></label>
    <div class="acts">${s ? '<button type="button" class="btn danger ghost" id="rm">Remover</button>' : ''}<button type="button" class="btn" data-close>Cancelar</button><button class="btn primary">Salvar</button></div></form>`);
  const F = f.el.querySelector('#fx');
  F.onsubmit = e => { e.preventDefault(); const n = { id: s?.id || uid(), nome: F.nome.value.trim(), apelidos: F.ap.value.trim() }; salvarSocios({ lista: s ? listaSocios().map(x => x.id === s.id ? n : x) : [...listaSocios(), n] }); f.fechar(); };
  const rm = f.el.querySelector('#rm'); if (rm) rm.onclick = () => { salvarSocios({ lista: listaSocios().filter(x => x.id !== s.id) }); f.fechar(); };
};
ACOES.acordoEd = d => {
  const a = acordos().find(x => x.id === d.id) || { escopo: 'quadras', quadras: [], partes: {} };
  const outras = new Set(acordos().filter(x => x.id !== a.id && x.escopo !== 'demais').flatMap(x => x.quadras || []));
  const f = folha(d.id ? 'Editar acordo' : 'Novo acordo', `<form class="form" id="fx"><label>Nome do acordo<input name="nome" value="${esc(a.nome || '')}" placeholder="Ex.: Primeiras quadras"></label>
    <label>Vale para<select name="esc">${opts([['quadras', 'As quadras marcadas abaixo'], ['lotes', 'Lotes específicos'], ['demais', 'Todos os demais lotes (os que não estão em outro acordo)']], a.escopo)}</select></label>
    <div id="lBox">${E().quadras.filter(q => lotesDe(q).length).map(q => `<div class="lbl">Quadra ${esc(q)}</div><div class="chips">${lotesDe(q).map(l => `<label class="pick"><input type="checkbox" name="ls" value="${esc(q + '|' + l)}" ${(a.lotes || []).includes(q + '|' + l) ? 'checked' : ''}><span>${esc(l)}</span></label>`).join('')}</div>`).join('')}</div>
    <div id="qBox"><div class="lbl">Quadras</div><div class="chips">${E().quadras.map(q => `<label class="pick"><input type="checkbox" name="qs" value="${esc(q)}" ${(a.quadras || []).includes(q) ? 'checked' : ''}><span>${esc(q)}${outras.has(q) ? '*' : ''}</span></label>`).join('')}</div>${outras.size ? '<p class="muted small">* já está em outro acordo</p>' : ''}</div>
    <div class="lbl">Divisão do resultado (%)</div>
    <div class="g3">${listaSocios().map(s => `<label>${esc(s.nome)}<input name="p_${esc(s.id)}" inputmode="decimal" value="${esc(String((a.partes || {})[s.id] || ''))}" placeholder="0"></label>`).join('')}</div>
    <p class="hint-data" id="pSum"></p>
    <details class="inv-box" ${invDe(a) ? 'open' : ''}><summary><b>Investidor com cota neste acordo</b> <span class="muted small">(opcional)</span></summary>
      <p class="muted small">Use quando um sócio investiu numa operação e tem direito a uma fatia do resultado dela. Essa fatia (a cota) sai do resultado antes da divisão acima e é repartida pela divisão própria dele; o restante segue a divisão acima.</p>
      <label>Como calcular a cota<select name="invM">${opts([['liquidado', 'Já liquidado: a cota é exatamente o que ele retirou além do aporte'], ['pct', 'Percentual do resultado da operação']], (a.inv || {}).modo || 'liquidado')}</select></label>
      <div class="g2"><label>Investidor<select name="invS">${opts(listaSocios().map(s => [s.id, s.nome]), (a.inv || {}).socio || '', 'Nenhum')}</select></label><label id="rowCota">Cota do resultado (%)<input name="invC" inputmode="decimal" value="${esc(String((a.inv || {}).cota || '').replace('.', ','))}" placeholder="Ex.: 7,2"></label></div>
      <label id="rowAte">Liquidado em <span class="muted">(aportes e retiradas até esta data; o que vier depois é investimento novo)</span><input type="date" name="invAte" value="${esc((a.inv || {}).ate || '')}"></label>
      <div class="acts left" id="rowSug"><button type="button" class="btn sm" id="sugCap">Calcular pelo capital investido</button><button type="button" class="btn sm" id="sugRet">Calcular pelo que ele já retirou</button></div>
      <p class="hint-data info" id="invHint"></p>
      <div class="lbl">Divisão da cota (%)</div>
      <div class="g3">${listaSocios().map(s => `<label>${esc(s.nome)}<input name="iv_${esc(s.id)}" inputmode="decimal" value="${esc(String(((a.inv || {}).div || {})[s.id] || ''))}" placeholder="0"></label>`).join('')}</div>
    </details>
    <div class="acts">${d.id ? '<button type="button" class="btn danger ghost" id="rm">Remover</button>' : ''}<button type="button" class="btn" data-close>Cancelar</button><button class="btn primary">Salvar</button></div></form>`);
  const F = f.el.querySelector('#fx');
  const soma = () => { const t = sum(listaSocios(), s => parseBR(F['p_' + s.id].value) || 0); $('#pSum', F).textContent = `Total ${t.toLocaleString('pt-BR')}%${Math.abs(t - 100) > 0.01 ? ' (precisa somar 100%)' : ''}`; return t; };
  const esc_ = () => { $('#qBox', F).classList.toggle('hidden', F.esc.value !== 'quadras'); $('#lBox', F).classList.toggle('hidden', F.esc.value !== 'lotes'); };
  F.addEventListener('input', soma); F.esc.onchange = esc_; soma(); esc_();
  /* Resultado e custo das quadras/lotes marcados no formulário, para sugerir a cota. */
  const escopoForm = () => { const RL = resultadoLotes(!!cfgSocios().adm); const ks = F.esc.value === 'quadras' ? $$('[name=qs]:checked', F).flatMap(x => lotesDe(x.value).map(l => x.value + '|' + l)) : F.esc.value === 'lotes' ? $$('[name=ls]:checked', F).map(x => x.value) : []; const xs = ks.map(k => RL[k]).filter(Boolean); return { custo: -sum(xs, x => x.custo), proj: sum(xs, x => x.proj) }; };
  const movInv = () => { const sid = F.invS.value; const ate = F.invM.value === 'liquidado' ? (F.invAte.value || dataLiquidacao({ socio: sid })) : '9999'; const ts = S.list.filter(t => ehAporteSocio(t) && socioDe(t) === sid && pago(t) && (t.pag || t.venc) <= ate); return { ent: sum(ts.filter(t => t.valor > 0), t => t.valor), sai: -sum(ts.filter(t => t.valor < 0), t => t.valor) }; };
  const pctNum = v => String(Math.round(v * 1e4) / 1e4).replace('.', ',');
  const modoInv = () => { const liq = F.invM.value === 'liquidado'; $('#rowCota', F).classList.toggle('hidden', liq); $('#rowSug', F).classList.toggle('hidden', liq); $('#rowAte', F).classList.toggle('hidden', !liq);
    if (liq && F.invS.value && !F.invAte.value) F.invAte.value = dataLiquidacao({ socio: F.invS.value });
    if (liq && F.invS.value) { const m = movInv(); const pi = (parseBR(F['iv_' + F.invS.value].value) || 0) / 100; $('#invHint', F).textContent = pi ? `Aportou ${brl(m.ent)} e retirou ${brl(m.sai)}: a parte dele no resultado é ${brl(m.sai - m.ent)} (${Math.round(pi * 100)}% da cota). Cota total ${brl((m.sai - m.ent) / pi)}, que sai do resultado antes da divisão principal. O saldo dele fica zero.` : 'Preencha a % do investidor na divisão da cota.'; } else if (liq) $('#invHint', F).textContent = ''; };
  F.invM.onchange = modoInv; F.invS.addEventListener('change', modoInv); F.addEventListener('input', e => { if ((e.target.name || '').startsWith('iv_') || e.target.name === 'invAte') modoInv(); }); modoInv();
  $('#sugCap', F).onclick = () => { if (!F.invS.value) { toast('Escolha o investidor.', 'err'); return; } const m = movInv(), e = escopoForm(); if (!e.custo) { toast('Marque as quadras do acordo.', 'err'); return; } const c = m.ent / e.custo * 100; F.invC.value = pctNum(c); $('#invHint', F).textContent = `Aportou ${brl(m.ent)} de um custo total de ${brl(e.custo)} nessas quadras: cota de ${pctNum(c)}%.`; };
  $('#sugRet', F).onclick = () => { if (!F.invS.value) { toast('Escolha o investidor.', 'err'); return; } const m = movInv(), e = escopoForm(); const pi = (parseBR(F['iv_' + F.invS.value].value) || 0) / 100; if (!pi) { toast('Preencha primeiro a % do investidor na divisão da cota.', 'err'); return; } if (!e.proj) { toast('Marque as quadras do acordo.', 'err'); return; }
    const ganho = m.sai - m.ent; const c = ganho / pi / e.proj * 100; F.invC.value = pctNum(c); $('#invHint', F).textContent = `Retirou ${brl(m.sai)} e aportou ${brl(m.ent)}: recebeu ${brl(ganho)} de resultado, que são ${Math.round(pi * 100)}% da cota. Cota = ${brl(ganho / pi)} ÷ resultado projetado de ${brl(e.proj)} = ${pctNum(c)}%.`; };
  F.onsubmit = e => { e.preventDefault(); if (Math.abs(soma() - 100) > 0.01) { toast('A divisão precisa somar 100%.', 'err'); return; }
    let inv = null; if (F.invS.value && (F.invM.value === 'liquidado' || parseBR(F.invC.value) > 0)) { const div = {}; listaSocios().forEach(s => { const v = parseBR(F['iv_' + s.id].value) || 0; if (v) div[s.id] = v; }); if (Math.abs(sum(Object.values(div), v => v) - 100) > 0.01) { toast('A divisão da cota precisa somar 100%.', 'err'); return; } inv = { socio: F.invS.value, modo: F.invM.value, cota: F.invM.value === 'liquidado' ? 0 : parseBR(F.invC.value), div, ...(F.invM.value === 'liquidado' ? { ate: F.invAte.value || dataLiquidacao({ socio: F.invS.value }) } : {}) }; }
    const partes = {}; listaSocios().forEach(s => { const v = parseBR(F['p_' + s.id].value) || 0; if (v) partes[s.id] = v; });
    const n = { id: a.id || uid(), nome: F.nome.value.trim() || (F.esc.value === 'demais' ? 'Demais quadras' : 'Acordo'), escopo: F.esc.value, quadras: F.esc.value === 'quadras' ? $$('[name=qs]:checked', F).map(x => x.value) : [], lotes: F.esc.value === 'lotes' ? $$('[name=ls]:checked', F).map(x => x.value) : [], partes, inv };
    salvarSocios({ acordos: d.id ? acordos().map(x => x.id === a.id ? n : x) : [...acordos(), n] }); f.fechar(); };
  const rm = f.el.querySelector('#rm'); if (rm) rm.onclick = () => { salvarSocios({ acordos: acordos().filter(x => x.id !== a.id) }); f.fechar(); };
};
ACOES.socioVer = d => {
  const x = calcSocios().so.find(y => y.s.id === d.id); if (!x) return;
  const ts = x.ts.slice().sort((a, b) => (dataRef(b) || '').localeCompare(dataRef(a) || ''));
  const porAno = {}; ts.filter(pago).forEach(t => { const a = (t.pag || t.venc).slice(0, 4); porAno[a] = porAno[a] || [0, 0]; porAno[a][t.valor > 0 ? 0 : 1] += t.valor; });
  folha(`${x.s.nome} · aportes e retiradas`, `<div class="tbl-wrap"><table class="tbl"><thead><tr><th class="l">Ano</th><th>Aportou</th><th>Retirou</th><th>Líquido</th></tr></thead><tbody>${Object.keys(porAno).sort().map(a => `<tr><td class="l">${a}</td><td class="pos">${brl(porAno[a][0])}</td><td class="neg">${brl(porAno[a][1])}</td><td>${brl(porAno[a][0] + porAno[a][1])}</td></tr>`).join('')}</tbody></table></div>
    <div class="lbl">Lançamentos (${ts.length})</div><div class="tlist">${ts.slice(0, 300).map(t => linhaTitulo(t, { mostrarPago: true })).join('')}</div>`, { largo: true });
};
