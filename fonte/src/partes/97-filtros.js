/* =========================================================
   Filtro de seleção múltipla (lista suspensa com caixas de marcar)
   grupos: [[rótulo do grupo, [[valor, rótulo], …]], …]
   ========================================================= */
function multiSel(id, vazio, grupos, sel, extra = '') {
  const s = new Set(sel || []); const todos = grupos.flatMap(g => g[1]);
  const nomes = todos.filter(([v]) => s.has(v)).map(x => x[1]);
  const resumo = !nomes.length ? vazio : nomes.length <= 2 ? nomes.join(', ') : `${nomes.length} selecionados`;
  return `<details class="msel${nomes.length ? ' ativo' : ''}" id="${id}" data-vazio="${esc(vazio)}"><summary>${esc(resumo)}</summary><div class="msel-pop">
    ${todos.length > 12 ? '<input type="search" class="msel-q" placeholder="Filtrar…">' : ''}${extra}
    <div class="msel-acts"><button type="button" class="btn link sm" data-ms="limpar">Limpar</button><button type="button" class="btn sm primary" data-ms="ok">Aplicar</button></div>
    <div class="msel-lista">${grupos.map(([g, xs]) => `<div class="msel-g">${g ? `<div class="msel-gt">${esc(g)}</div>` : ''}${xs.map(([v, l, cls]) => `<label class="msel-i ${cls || ''}" data-txt="${esc(normTxt(g + ' ' + l))}"><input type="checkbox" value="${esc(v)}" ${s.has(v) ? 'checked' : ''}><span>${esc(l)}</span></label>`).join('')}</div>`).join('')}</div></div></details>`;
}
/* Quadras inteiras e lotes de várias quadras. Valores: "Q:K3" (quadra inteira) e "L:K3|12". */
function multiSelLocais(id, sel, obra) {
  const qs = quadrasDaObra(obra && obra !== '__emp' ? obra : '');
  const grupos = [['', [['G:-', 'Geral (sem quadra)']]], ...qs.map(q => [`Quadra ${q}`, [['Q:' + q, `Quadra ${q} inteira`, 'msel-q-int'], ...lotesDe(q).map(l => ['L:' + q + '|' + l, `${q} ${l}`, 'msel-lote'])]])];
  return multiSel(id, 'Todas as quadras e lotes', grupos, sel);
}
function casaLocais(t, locs) {
  for (const x of locs) {
    if (x === 'G:-') { if (!t.quadra || t.quadra === '-') return true; continue; }
    if (x.startsWith('Q:')) { if (t.quadra === x.slice(2)) return true; continue; }
    const [q, l] = x.slice(2).split('|'); if (t.quadra === q && String(t.lote) === l) return true;
  }
  return false;
}
function ligarMultiSel(el, id, aoMudar) {
  const d = $('#' + id, el); if (!d) return;
  const valores = () => $$('.msel-lista input:checked', d).map(c => c.value);
  const resumo = () => { const xs = $$('.msel-lista input:checked', d).map(c => c.closest('label').textContent.trim()); d.querySelector('summary').textContent = !xs.length ? d.dataset.vazio : xs.length <= 2 ? xs.join(', ') : xs.length + ' selecionados'; d.classList.toggle('ativo', !!xs.length); };
  const aplicar = () => { d.open = false; resumo(); aoMudar(valores()); };
  d.querySelector('[data-ms=ok]').onclick = aplicar;
  d.querySelector('[data-ms=limpar]').onclick = () => { $$('.msel-lista input', d).forEach(c => c.checked = false); aplicar(); };
  const q = d.querySelector('.msel-q'); if (q) q.oninput = () => { const k = normTxt(q.value); $$('.msel-i', d).forEach(i => i.classList.toggle('hidden', !!k && !i.dataset.txt.includes(k))); };
  // quadra inteira marcada: desmarca os lotes dela (redundantes)
  d.addEventListener('change', e => { const v = e.target.value || ''; if (v.startsWith('Q:') && e.target.checked) $$('.msel-lista input', d).filter(c => c.value.startsWith('L:' + v.slice(2) + '|')).forEach(c => c.checked = false); if (v.startsWith('L:') && e.target.checked) { const qi = $$('.msel-lista input', d).find(c => c.value === 'Q:' + v.slice(2).split('|')[0]); if (qi) qi.checked = false; } });
  d.addEventListener('toggle', () => { if (d.open) { $$('details.msel[open]').forEach(o => { if (o !== d) o.open = false; }); if (q) setTimeout(() => q.focus(), 30); } });
}
document.addEventListener('click', e => { if (!e.target.closest('details.msel')) $$('details.msel[open]').forEach(o => { o.open = false; }); });

/* Números grandes com centavos: reduz a fonte do valor até caber no cartão (nunca corta o número). */
function encaixarValores() {
  for (const el of document.querySelectorAll('.kpi .val, .eh-val')) {
    el.style.fontSize = ''; let tam = parseFloat(getComputedStyle(el).fontSize); let n = 0;
    while (el.scrollWidth > el.clientWidth + 1 && tam > 12 && n++ < 20) { tam -= 1; el.style.fontSize = tam + 'px'; }
  }
}
{ let tm = 0; const agenda = () => { clearTimeout(tm); tm = setTimeout(encaixarValores, 40); };
  new MutationObserver(agenda).observe(document.documentElement, { childList: true, subtree: true }); addEventListener('resize', agenda); }
