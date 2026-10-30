/* =========================================================
   Alterar a divisão de um lançamento entre quadras e lotes
   (ex.: esqueceu de marcar um lote). Refaz as partes mantendo
   valor total, datas, conta, conciliação, anexos e CNPJ.
   ========================================================= */
/* Partes do mesmo pagamento: mesmo grupo e mesma data (ou as partes parecidas sem grupo). */
function partesDe(t) {
  const d = x => (pago(x) ? (x.pag || x.venc) : x.venc) + '|' + x.status;
  if (t.grupo) return S.list.filter(x => x.grupo === t.grupo && d(x) === d(t));
  const k = chaveJuntar(t); return S.list.filter(x => chaveJuntar(x) === k);
}
ACOES.redividir = d => {
  const t0 = S.byId.get(d.id); if (!t0) return;
  const partes = partesDe(t0); const todas = t0.grupo ? S.list.filter(x => x.grupo === t0.grupo) : partes;
  const datas = [...new Set(todas.map(x => (pago(x) ? (x.pag || x.venc) : x.venc) + '|' + x.status))];
  const sel = new Set(partes.map(x => x.quadra + '|' + (x.lote || '-')));
  const qsSel = new Set(partes.map(x => x.quadra).filter(q => q && q !== '-'));
  const base = baseItem(t0).replace(/\s*[-–]\s*[A-Z]\w*\s*$/, '');
  const tot = sum(partes, x => x.valor);
  const qs = quadrasDaObra(variasObras() ? obraT(t0) : '');
  const f = folha('Alterar quadras e lotes', `<form class="form" id="fx"><p><b>${esc(base)}</b> · ${partes.length} parte(s) · total <b class="${cls(tot)}">${brl2(tot)}</b></p>
    <p class="muted small">Marque todas as quadras e lotes que este gasto deve cobrir. O valor total é redividido em partes iguais; data, conta, conciliação, nota, anexos e CNPJ continuam os mesmos.</p>
    <div class="lbl">Quadras</div><div class="chips" id="rdQ">${qs.map(q => `<label class="pick"><input type="checkbox" name="rq" value="${esc(q)}" ${qsSel.has(q) ? 'checked' : ''}><span>${esc(q)}</span></label>`).join('')}</div>
    <div id="rdL"></div>
    ${datas.length > 1 ? `<label class="chk"><input type="checkbox" name="todas" checked> Aplicar às ${datas.length} parcelas deste lançamento (cada parcela mantém o próprio valor e data)</label>` : ''}
    <p class="hint-data info" id="rdPrev"></p>
    <div class="acts"><button type="button" class="btn" data-close>Cancelar</button><button class="btn primary">Salvar nova divisão</button></div></form>`);
  const F = f.el.querySelector('#fx');
  const lotesBox = () => { const ant = new Set([...$$('[name=rl]:checked', F).map(c => c.value), ...(lotesBox.ini ? [] : sel)]); lotesBox.ini = true;
    $('#rdL', F).innerHTML = $$('[name=rq]:checked', F).map(c => c.value).filter(q => lotesDe(q).length).map(q => `<div class="lbl">Lotes da quadra ${esc(q)} <span class="muted">(nenhum = custo geral da quadra)</span> <button type="button" class="btn link sm" data-lt="all" data-q="${esc(q)}">Todos</button><button type="button" class="btn link sm" data-lt="none" data-q="${esc(q)}">Limpar</button></div><div class="chips">${lotesDe(q).map(l => `<label class="pick"><input type="checkbox" name="rl" value="${esc(q + '|' + l)}" ${ant.has(q + '|' + l) ? 'checked' : ''}><span>${esc(l)}</span></label>`).join('')}</div>`).join('');
    $$('[data-lt]', F).forEach(b => b.onclick = () => { $$('[name=rl]', F).filter(c => c.value.startsWith(b.dataset.q + '|')).forEach(c => c.checked = b.dataset.lt === 'all'); prev(); }); prev(); };
  const pares = () => { const qsM = $$('[name=rq]:checked', F).map(c => c.value); const ls = $$('[name=rl]:checked', F).map(c => c.value.split('|')); if (!qsM.length) return [['-', '-']]; return qsM.flatMap(q => { const x = ls.filter(p => p[0] === q); return x.length ? x : [[q, '-']]; }); };
  const prev = () => { const P = pares(); $('#rdPrev', F).textContent = `${P.length} parte(s) de ${brl2(tot / P.length)}: ${P.map(p => p[1] !== '-' ? p[0] + ' ' + p[1] : p[0] !== '-' ? p[0] + ' (geral)' : 'geral da obra').join(', ')}`; };
  $('#rdQ', F).onchange = lotesBox; F.addEventListener('change', e => { if (e.target.name === 'rl') prev(); }); lotesBox();
  F.onsubmit = e => { e.preventDefault(); const P = pares();
    const grupos = F.todas && F.todas.checked ? datas.map(dk => todas.filter(x => (pago(x) ? (x.pag || x.venc) : x.venc) + '|' + x.status === dk)) : [partes];
    const novos = [], apagar = []; const g = t0.grupo || uid();
    for (const ps of grupos) {
      const ref = ps[0]; const totV = sum(ps, x => x.valor), totJ = sum(ps, x => x.juros || 0), totQ = sum(ps, x => x.qtd || 0);
      const cents = Math.round(Math.abs(totV) * 100), sg = totV < 0 ? -1 : 1, b = Math.floor(cents / P.length); let r = cents - b * P.length;
      const jc = Math.round(Math.abs(totJ) * 100), jb = Math.floor(jc / P.length); let jr = jc - jb * P.length;
      const nomeBase = baseItem(ref).replace(new RegExp('\\s*[-–]\\s*' + reEsc(ref.quadra || '§') + '\\s*$'), '');
      const livres = ps.slice();
      P.forEach(([q, l]) => {
        const v = sg * (b + (r-- > 0 ? 1 : 0)) / 100, j = (totJ < 0 ? -1 : 1) * (jb + (jr-- > 0 ? 1 : 0)) / 100;
        const igual = livres.findIndex(x => x.quadra === q && String(x.lote || '-') === String(l)); const old = igual >= 0 ? livres.splice(igual, 1)[0] : null;
        const n = { ...ref, ...(old ? { id: old.id, hist: old.hist } : { id: uid(), hist: [], criado: Date.now() }), item: nomeBase + (l !== '-' ? ` - ${q} ${l}` : q !== '-' ? ` - ${q}` : ''), valor: v, juros: j, quadra: q, lote: l, local: q !== '-' && E().quadras.includes(q) ? obraDe(q) : ref.local, grupo: P.length > 1 || grupos.length > 1 ? g : (ref.grupo || '') };
        if (totQ) n.qtd = Math.round(totQ / P.length * 1000) / 1000;
        novos.push(n);
      });
      apagar.push(...livres.map(x => x.id));
    }
    salvarTitulos(novos, { acao: 'redividiu entre lotes' }); if (apagar.length) excluirTitulos(apagar, { semLixeira: true });
    f.fechar(); $$('#sheets .sheet-wrap').forEach(w => { w.classList.add('out'); setTimeout(() => w.remove(), 180); }); refreshView(); toast(`Divisão atualizada: ${P.length} parte(s)${grupos.length > 1 ? ` em ${grupos.length} parcelas` : ''}`);
  };
};
