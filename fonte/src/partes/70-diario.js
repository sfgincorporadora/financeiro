/* =========================================================
   Diário de obra: clima, equipe, serviços, ocorrências e fotos (as fotos vão para o Drive)
   Guardado em diario/{AAAA-MM}.
   ========================================================= */
usarBlocos('diario', o => { const l = []; for (const b in o) for (const k in (o[b].itens || {})) { try { const x = JSON.parse(o[b].itens[k]); x._b = b; l.push(x); } catch (e) { /* ignora */ } } S.diario = l; });
const CLIMAS = [['sol', 'Sol'], ['nublado', 'Nublado'], ['chuva', 'Chuva'], ['forte', 'Chuva forte, sem trabalho']];
const climaNome = k => (CLIMAS.find(x => x[0] === k) || ['', '—'])[1];
function salvarDiario(x) {
  const b = (x.data || hoje()).slice(0, 7); const o = { ...x }; const velho = o._b; delete o._b;
  S.diario = [...(S.diario || []).filter(y => y.id !== o.id), { ...o, _b: b }];
  if (velho && velho !== b) gravar(doc(db, 'diario', velho), { itens: { [o.id]: deleteField() } }, { merge: true }).catch(onErroEscrita);
  gravar(doc(db, 'diario', b), { itens: { [o.id]: JSON.stringify(o) } }, { merge: true }).catch(onErroEscrita);
}
const equipeTotal = txt => sum(String(txt || '').split('\n'), l => +((l.match(/(\d+)\s*$/) || [])[1] || 0));
let diaObra = '', diaQ = '', diaMes = '';
VIEWS.diario = {
  titulo: 'Diário de obra',
  refresh(el) {
    if (!diaMes) diaMes = hoje().slice(0, 7);
    const ds = (S.diario || []).filter(x => (!diaObra || x.obra === diaObra) && (!diaQ || (x.quadras || []).includes(diaQ)) && (x.data || '').slice(0, 7) === diaMes).sort((a, b) => b.data.localeCompare(a.data));
    const meses = [...new Set([hoje().slice(0, 7), ...(S.diario || []).map(x => (x.data || '').slice(0, 7))])].filter(Boolean).sort().reverse();
    const chuva = ds.filter(x => ['chuva', 'forte'].includes(x.clima?.m) || ['chuva', 'forte'].includes(x.clima?.t)).length;
    el.innerHTML = `<div class="filtros card">${variasObras() ? `<label class="tgl">Obra <select id="dO">${opts(E().obras, diaObra, 'Todas')}</select></label>` : ''}<label class="tgl">Quadra <select id="dQ">${opts(E().quadras, diaQ, 'Todas')}</select></label><label class="tgl">Mês <select id="dMs">${opts(meses.map(m => [m, mesNome(m)]), diaMes)}</select></label><button class="btn primary rw" data-act="diaNovo">Registrar dia</button><button class="btn" data-act="diaPdf">PDF do mês</button></div>
      <div class="grid kpis k4">${kpi('Dias registrados', ds.length, mesNome(diaMes), '')}${kpi('Dias com chuva', chuva, '', chuva ? 'neg' : '')}${kpi('Média de pessoas', ds.length ? fNum2.format(sum(ds, x => equipeTotal(x.equipe)) / ds.length).replace(',00', '') : '—', 'por dia', '')}${kpi('Fotos', sum(ds, x => (x.fotos || []).length), sum(ds, x => (x.fotos || []).filter(f => !f.link).length) ? sum(ds, x => (x.fotos || []).filter(f => !f.link).length) + ' aguardando envio ao Drive' : 'no Google Drive', '')}</div>
      ${ds.map(x => `<div class="card diario-dia"><h2><span>${new Date(x.data + 'T12:00').toLocaleDateString('pt-BR', { weekday: 'long', day: 'numeric', month: 'long' })}</span><button class="btn sm ghost rw" data-act="diaEd" data-id="${esc(x.id)}">Editar</button></h2>
        <p class="muted small">${esc([x.obra, (x.quadras || []).length ? 'Quadra ' + x.quadras.join(', ') : ''].filter(Boolean).join(' · '))} · manhã: ${climaNome(x.clima?.m)} · tarde: ${climaNome(x.clima?.t)} · ${equipeTotal(x.equipe)} pessoa(s) · por ${quem(x.autor)}</p>
        ${x.servicos ? `<div class="lbl">Serviços executados</div><p class="pre">${esc(x.servicos)}</p>` : ''}${x.equipe ? `<div class="lbl">Equipe</div><p class="pre small">${esc(x.equipe)}</p>` : ''}${x.ocorrencias ? `<div class="lbl">Ocorrências</div><p class="pre">${esc(x.ocorrencias)}</p>` : ''}
        ${(x.fotos || []).length ? `<div class="diario-fotos">${x.fotos.map((f, i) => `<button class="foto-chip" data-act="diaFoto" data-id="${esc(x.id)}" data-i="${i}">${ic('clip', 14)} Foto ${i + 1}${f.link ? '' : ' (aguardando Drive)'}</button>`).join('')}</div>` : ''}</div>`).join('') || `<div class="card"><p class="muted">Nenhum dia registrado em ${mesNome(diaMes)}. Toque em "Registrar dia" ao fim de cada dia de obra: leva um minuto e vira o histórico para medições, vistorias da Caixa e conversas com empreiteiros.</p></div>`}`;
    const on = (id, fn) => { const s = $('#' + id, el); if (s) s.onchange = e => { fn(e.target.value); this.refresh(el); }; };
    on('dO', v => diaObra = v); on('dQ', v => diaQ = v); on('dMs', v => diaMes = v);
  }
};
ACOES.diaNovo = () => formDiario(null);
ACOES.diaEd = d => formDiario((S.diario || []).find(x => x.id === d.id));
function formDiario(x) {
  const ult = (S.diario || []).slice().sort((a, b) => b.data.localeCompare(a.data))[0];
  const f = folha(x ? 'Editar dia' : 'Registrar dia de obra', `<form class="form" id="fx"><div class="g2"><label>Data<input type="date" name="data" required value="${esc(x?.data || hoje())}"></label>${variasObras() ? `<label>Obra<select name="obra">${opts(E().obras, x?.obra || diaObra || E().obras[0])}</select></label>` : ''}</div>
    <div class="lbl">Quadras com serviço</div><div class="chips">${E().quadras.map(q => `<label class="chip"><input type="checkbox" name="qs" value="${esc(q)}" ${(x?.quadras || (diaQ ? [diaQ] : [])).includes(q) ? 'checked' : ''}> ${esc(q)}</label>`).join('')}</div>
    <div class="g2"><label>Clima de manhã<select name="cm">${opts(CLIMAS, x?.clima?.m || 'sol')}</select></label><label>Clima à tarde<select name="ct">${opts(CLIMAS, x?.clima?.t || 'sol')}</select></label></div>
    <label>Equipe <span class="muted">(uma função por linha com a quantidade)</span>${!x && ult?.equipe ? ' <button type="button" class="btn link sm" id="cpEq">copiar do último dia</button>' : ''}<textarea name="equipe" rows="3" placeholder="Pedreiro 4&#10;Servente 6&#10;Eletricista 1">${esc(x?.equipe || '')}</textarea></label>
    <label>Serviços executados<textarea name="servicos" rows="3" placeholder="Ex.: alvenaria lotes 19 a 21, concretagem da laje do 22">${esc(x?.servicos || '')}</textarea></label>
    <label>Ocorrências <span class="muted">(atrasos, faltas de material, visitas, acidentes)</span><textarea name="ocorrencias" rows="2">${esc(x?.ocorrencias || '')}</textarea></label>
    <label>Fotos<input type="file" name="fotos" accept="image/*" multiple></label>${(x?.fotos || []).length ? `<p class="muted small">${x.fotos.length} foto(s) já registradas.</p>` : ''}
    <div class="acts">${x ? '<button type="button" class="btn danger ghost" id="dDel">Excluir</button>' : ''}<button type="button" class="btn" data-close>Cancelar</button><button class="btn primary">Salvar</button></div></form>`, { largo: true });
  const F = f.el.querySelector('#fx');
  const cp = F.querySelector('#cpEq'); if (cp) cp.onclick = () => { F.equipe.value = ult.equipe; };
  F.onsubmit = async e => {
    e.preventDefault(); const b = F.querySelector('button.primary'); b.disabled = true; b.textContent = 'Salvando…';
    const fotos = [...(x?.fotos || [])];
    try { for (const fl of F.fotos.files) fotos.push({ id: await salvarAnexo(fl) }); } catch (err) { toast(err.message, 'err'); b.disabled = false; b.textContent = 'Salvar'; return; }
    const n = { ...(x || { id: 'dia' + uid(), criado: Date.now(), autor: S.user?.email || '' }), data: F.data.value, obra: F.obra ? F.obra.value : E().obras[0], quadras: $$('[name=qs]:checked', F).map(c => c.value), clima: { m: F.cm.value, t: F.ct.value }, equipe: F.equipe.value.trim(), servicos: F.servicos.value.trim(), ocorrencias: F.ocorrencias.value.trim(), fotos };
    salvarDiario(n); diaMes = n.data.slice(0, 7); f.fechar(); toast('Dia registrado'); refreshView();
  };
  const del = F.querySelector('#dDel'); if (del) del.onclick = async () => { if (await confirmar('Excluir este registro do diário? As fotos já enviadas continuam no Drive.', 'Excluir')) { gravar(doc(db, 'diario', x._b), { itens: { [x.id]: deleteField() } }, { merge: true }).catch(onErroEscrita); S.diario = S.diario.filter(y => y.id !== x.id); f.fechar(); refreshView(); } };
}
ACOES.diaFoto = d => { const x = (S.diario || []).find(y => y.id === d.id); const fo = x && x.fotos[+d.i]; if (fo) verAnexo(fo.id, fo.link); };
/* Fotos pendentes vão para SFG Financeiro / Diário de obra / AAAA-MM, junto com o envio dos anexos. */
async function pastaDiario(ym) {
  const c = clone(S.drive || {}); c.diario = c.diario || {};
  if (!c.diario.raiz || !(await pastaExiste(c.diario.raiz))) { c.diario = { raiz: (await novaPasta('Diário de obra', await raizSFG())).id, meses: {} }; }
  if (!c.diario.meses[ym]) c.diario.meses[ym] = (await novaPasta(ym, c.diario.raiz)).id;
  S.drive = { ...(S.drive || {}), diario: c.diario }; gravar(doc(db, 'config', 'drive'), S.drive).catch(onErroEscrita);
  return c.diario.meses[ym];
}
DRIVE_EXTRA.push({
  pendentes: () => sum(S.diario || [], x => (x.fotos || []).filter(f => f.id && !f.link).length),
  enviar: async () => {
    let ok = 0;
    for (const x0 of (S.diario || []).slice()) {
      if (!(x0.fotos || []).some(f => f.id && !f.link)) continue;
      const x = { ...x0, fotos: x0.fotos.map(f => ({ ...f })) };
      for (let i = 0; i < x.fotos.length; i++) {
        const fo = x.fotos[i]; if (!fo.id || fo.link) continue;
        const s = await getDoc(doc(db, 'anexos', fo.id)); if (!s.exists()) continue; const a = s.data();
        if (a.link) { fo.link = a.link; fo.drive = a.drive; ok++; continue; }
        if (!a.b64) continue;
        const bin = atob(a.b64); const u8 = new Uint8Array(bin.length); for (let j = 0; j < bin.length; j++) u8[j] = bin.charCodeAt(j);
        const nome = `${x.data} ${x.obra || ''} ${(x.quadras || []).join('-')} foto ${i + 1}.jpg`.replace(/[\/\\:*?"<>|#%]+/g, '-').replace(/\s+/g, ' ');
        const r = await enviarArquivoDrive(new Blob([u8], { type: a.mime }), nome, await pastaDiario(x.data.slice(0, 7)));
        fo.link = r.webViewLink; fo.drive = r.id; ok++;
        gravar(doc(db, 'anexos', fo.id), { drive: r.id, link: r.webViewLink, mime: a.mime, nome, enviado: Date.now() }).catch(onErroEscrita);
      }
      salvarDiario(x);
    }
    return ok;
  }
});
ACOES.diaPdf = async () => {
  try {
    await carregarLib('pdf'); const ds = (S.diario || []).filter(x => (!diaObra || x.obra === diaObra) && (!diaQ || (x.quadras || []).includes(diaQ)) && (x.data || '').slice(0, 7) === diaMes).sort((a, b) => a.data.localeCompare(b.data));
    if (!ds.length) { toast('Nenhum dia registrado neste mês.', 'err'); return; }
    const pdf = new jsPDF({ unit: 'pt', format: 'a4' }); const W = pdf.internal.pageSize.getWidth(); const P = s => String(s || '').replace(/[  ]/g, ' ');
    try { const lg = await imgParaPNG(LOGO_COMPLETO, 480); pdf.addImage(lg.data, 'PNG', 36, 16, 96, 96 * lg.r); } catch (e) { /* sem logo */ }
    pdf.setTextColor(11, 32, 70); pdf.setFont('helvetica', 'bold'); pdf.setFontSize(16); pdf.text('Diário de obra', W - 36, 40, { align: 'right' }); pdf.setFont('helvetica', 'normal'); pdf.setFontSize(10.5); pdf.text(P(`${mesNome(diaMes)}${diaObra ? ' · ' + diaObra : ''}${diaQ ? ' · quadra ' + diaQ : ''}`), W - 36, 58, { align: 'right' });
    autoTable(pdf, { startY: 104, head: [['Data', 'Clima (manhã / tarde)', 'Equipe', 'Serviços executados', 'Ocorrências', 'Fotos']], body: ds.map(x => [dBR(x.data), `${climaNome(x.clima?.m)} / ${climaNome(x.clima?.t)}`, P(x.equipe || '').replace(/\n/g, ', ') + ` (${equipeTotal(x.equipe)})`, P(x.servicos), P(x.ocorrencias), String((x.fotos || []).length)]), margin: { left: 30, right: 30 }, styles: { fontSize: 8, cellPadding: 3, valign: 'top' }, headStyles: { fillColor: [11, 32, 70] }, columnStyles: { 0: { cellWidth: 52 }, 1: { cellWidth: 70 }, 5: { cellWidth: 32, halign: 'center' } } });
    const lk = S.drive?.diario?.raiz; if (lk) { pdf.setFontSize(8.5); pdf.setTextColor(91, 107, 130); pdf.text('Fotos na pasta "SFG Financeiro / Diário de obra" do Google Drive.', 30, pdf.lastAutoTable.finalY + 18); }
    await baixarArquivo(pdf.output('blob'), `SFG Diário de obra ${diaMes}.pdf`);
  } catch (e) { toast('Erro ao gerar PDF: ' + e.message, 'err'); }
};
