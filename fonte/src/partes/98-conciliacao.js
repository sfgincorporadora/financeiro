/* =========================================================
   Conciliação: recusar sugestão, escolher o lançamento à mão
   e limpar o extrato importado para recomeçar.
   ========================================================= */
const botoesConcExtra = l => ` <button class="btn sm" data-act="cEsc" data-id="${l.id}">Escolher outro</button> <button class="btn sm link" data-act="cRej" data-id="${l.id}">Não é este</button> <button class="btn sm link" data-act="cIgn" data-id="${l.id}">Ignorar</button>`;
function gravarLinhaConc(id, fn) {
  const ex = S.extratos && S.extratos[contaId(concConta)]; if (!ex || !ex.linhas[id]) return;
  const l = JSON.parse(ex.linhas[id]); fn(l); ex.linhas[id] = JSON.stringify(l);
  gravar(doc(db, 'extratos', contaId(concConta)), { linhas: { [id]: ex.linhas[id] } }, { merge: true }).catch(onErroEscrita); refreshView();
}
/* "Não é este": a sugestão some e o app procura outro lançamento (ou deixa sem lançamento). */
ACOES.cRej = d => { const l = linhaConc(d.id); if (!l) return; gravarLinhaConc(d.id, x => { x.rej = [...new Set([...(x.rej || []), ...l.ts.map(t => t.id)])]; }); toast('Sugestão descartada'); };
/* Escolher à mão o lançamento que corresponde à linha do banco. */
ACOES.cEsc = d => {
  const l = linhaConc(d.id); if (!l) return;
  const usados = new Set(S.list.filter(t => t.conc).map(t => t.id));
  const cand = S.list.filter(t => !usados.has(t.id) && Math.sign(t.valor) === Math.sign(l.v) && Math.abs(new Date(dataRef(t) || t.venc) - new Date(l.d)) <= 120 * 864e5);
  const f = folha('Escolher lançamento', `<p class="muted small">No banco: ${dBR(l.d)} · ${esc(l.desc)} · <b class="${cls(l.v)}">${brl2(l.v)}</b>. Marque um ou mais lançamentos (por exemplo as partes de um pagamento dividido entre lotes). Os que estiverem em aberto recebem a baixa na data do banco.</p>
    <input type="search" id="eq" placeholder="Buscar descrição, fornecedor, valor…" style="width:100%;margin-bottom:8px">
    <div class="vinc-lista" id="el"></div><p class="small" id="esum"></p>
    <div class="acts"><button class="btn" data-close>Cancelar</button><button class="btn primary" id="eOk">Conciliar</button></div>`, { largo: true });
  const ord = cand.sort((a, b) => Math.abs(a.valor - l.v) - Math.abs(b.valor - l.v) || Math.abs(new Date(dataRef(a)) - new Date(l.d)) - Math.abs(new Date(dataRef(b)) - new Date(l.d)));
  const pinta = () => { const q = normTxt($('#eq', f.el).value); const xs = ord.filter(t => !q || normTxt(t.item + ' ' + (t.forn || '') + ' ' + fNum2.format(Math.abs(t.valor))).includes(q)).slice(0, 200);
    $('#el', f.el).innerHTML = xs.map(t => `<label class="vinc"><input type="checkbox" value="${esc(t.id)}"><span>${dBR(dataRef(t))}</span><span class="wrap">${esc(t.item)}<br><small class="muted">${pago(t) ? 'pago · ' + esc(t.conta || '') : 'em aberto'}${t.forn ? ' · ' + esc(t.forn) : ''}</small></span><b class="${cls(t.valor)}">${brl2(t.valor)}</b></label>`).join('') || '<p class="muted">Nenhum lançamento parecido.</p>'; soma(); };
  const soma = () => { const ids = $$('#el input:checked', f.el).map(x => x.value); const s = sum(ids.map(id => S.byId.get(id)).filter(Boolean), t => t.valor + (pago(t) ? (t.juros || 0) : 0)); $('#esum', f.el).innerHTML = ids.length ? `Selecionado ${brl2(s)} · diferença ${brl2(l.v - s)}` : ''; };
  $('#eq', f.el).oninput = pinta; $('#el', f.el).onchange = soma; pinta();
  $('#eOk', f.el).onclick = () => { const ts = $$('#el input:checked', f.el).map(x => S.byId.get(x.value)).filter(Boolean); if (!ts.length) { toast('Marque pelo menos um lançamento.', 'err'); return; }
    salvarTitulos(ts.map(t => pago(t) ? { ...t, conc: l.id } : { ...t, status: 'pago', pag: l.d, conta: concConta, conc: l.id })); f.fechar(); toast('Conciliado'); };
};
/* Limpar o extrato importado da conta. */
ACOES.cLimpar = () => {
  const a = concCache; if (!a) return; const ls = a.linhas; const ok = ls.filter(l => l.st === 'ok'); const pend = ls.length - ok.length;
  const f = folha('Limpar extrato · ' + concConta, `<p>Este extrato tem <b>${ls.length}</b> linha(s): ${ok.length} conciliada(s) e ${pend} pendente(s) (sugestões, sem lançamento ou ignoradas).</p>
    <div class="opcoes-limpar">
      <label class="chk"><input type="radio" name="m" value="pend" checked> <span><b>Apagar só as pendentes</b><br><span class="muted small">As conciliadas continuam. Depois importe o extrato novo: as linhas voltam limpas, sem sugestões recusadas.</span></span></label>
      <label class="chk"><input type="radio" name="m" value="tudo"> <span><b>Recomeçar do zero</b><br><span class="muted small">Apaga todas as linhas importadas desta conta e desfaz as conciliações (os lançamentos e as baixas continuam; só perdem o vínculo com o extrato).</span></span></label>
    </div>
    <div class="acts"><button class="btn" data-close>Cancelar</button><button class="btn danger" id="lOk">Limpar</button></div>`);
  $('#lOk', f.el).onclick = () => {
    const m = f.el.querySelector('[name=m]:checked').value; const ex = S.extratos[contaId(concConta)] || { linhas: {} };
    const manter = {}; if (m === 'pend') for (const l of ok) if (ex.linhas[l.id]) manter[l.id] = ex.linhas[l.id];
    if (m === 'tudo') { const ids = new Set(ls.map(l => l.id)); const des = S.list.filter(t => t.conc && ids.has(t.conc)).map(t => { const n = { ...t }; delete n.conc; return n; }); if (des.length) salvarTitulos(des, { sistema: true }); }
    const novo = { conta: concConta, linhas: manter, importado: Date.now() }; if (m === 'pend' && ex.saldo != null) { novo.saldo = ex.saldo; novo.saldoData = ex.saldoData; }
    S.extratos[contaId(concConta)] = novo; gravar(doc(db, 'extratos', contaId(concConta)), novo).catch(onErroEscrita);
    f.fechar(); toast(m === 'tudo' ? 'Extrato apagado. Importe o arquivo novo.' : `${pend} linha(s) pendente(s) apagada(s)`); refreshView();
  };
};
