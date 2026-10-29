/* =========================================================
   Vendas: plano de pagamento do cliente e cobrança por PIX
   ========================================================= */
const TIPOS_PLANO = [['entrada', 'Entrada'], ['direta', 'Parcelas diretas'], ['financiamento', 'Financiamento Caixa'], ['fgts', 'FGTS'], ['subsidio', 'Subsídio'], ['outro', 'Outro']];
const tipoPlanoNome = k => (TIPOS_PLANO.find(x => x[0] === k) || ['', 'Parcela'])[1];
const vendaLote = k => ((S.vendas && S.vendas.lotes) || {})[k] || {};
function salvarVendaLote(k, v) { S.vendas = S.vendas || { lotes: {} }; S.vendas.lotes = { ...(S.vendas.lotes || {}), [k]: v }; gravar(doc(db, 'config', 'vendas'), { lotes: { [k]: v } }, { merge: true }).catch(onErroEscrita); }
const idPlano = (k, itemId, i) => 'pl' + hashTxt(k) + itemId + '_' + i;
/* Transforma o plano em parcelas: [{id, tipo, i, n, venc, vencOrig, valor}] */
function parcelasPlano(k, p) {
  const out = [];
  for (const it of p.itens || []) {
    const n = Math.max(1, Math.min(360, +it.n || 1)), passo = Math.max(1, +it.passo || 1);
    const tot = Math.round((+it.valor || 0) * 100), b = Math.floor(tot / n); let r = tot - b * n;
    for (let i = 0; i < n; i++) {
      let v = (b + (r-- > 0 ? 1 : 0)) / 100;
      if (it.tipo === 'direta' && +p.reajuste) v = r2(v * Math.pow(1 + p.reajuste / 100, Math.floor(i * passo / 12)));
      const base = addMonths(it.venc, i * passo); const venc = p.util === false ? base : proximoDiaUtil(base);
      out.push({ id: idPlano(k, it.id, i), tipo: it.tipo, desc: it.desc || '', i, n, venc, vencOrig: venc !== base ? base : '', valor: v });
    }
  }
  return out;
}
function formPlano(k) {
  const [q, l] = k.split('|'); const v = vendaLote(k); const p = v.plano || { valor: 0, itens: [], util: true, reajuste: 0 };
  const jaLanc = S.list.filter(t => t.quadra === q && String(t.lote) === String(l) && t.destino === 'Venda' && !t.plano);
  const itens = (p.itens || []).map(x => ({ ...x }));
  if (!itens.length) itens.push({ id: uid().slice(-6), tipo: 'entrada', valor: 0, n: 1, passo: 1, venc: hoje() }, { id: uid().slice(-6), tipo: 'financiamento', valor: 0, n: 1, passo: 1, venc: addMonths(hoje(), 2) });
  const f = folha(`Plano de pagamento · quadra ${q}, lote ${l}`, `<form class="form" id="fx">
    <div class="g3"><label>Cliente<input name="cliente" value="${esc(v.cliente || '')}"></label><label>Preço de venda (R$)<input name="valor" inputmode="decimal" value="${p.valor ? fNum2.format(p.valor) : ''}"></label><label>Reajuste anual das parcelas diretas (%)<input name="reaj" inputmode="decimal" value="${p.reajuste ? String(p.reajuste).replace('.', ',') : ''}" placeholder="Ex.: INCC estimado"></label></div>
    ${variasEmpresas() ? `<label>CNPJ que recebe<select name="emp" required>${opts(empOpts(), p.emp || empDaObra(obraDe(q)), 'Escolha o CNPJ…')}</select></label>` : ''}
    <div class="lbl">Composição</div><div id="pItens"></div>
    <button type="button" class="btn sm" id="pAdd">+ Adicionar parte</button>
    <label class="chk"><input type="checkbox" name="util" ${p.util === false ? '' : 'checked'}> Vencimentos em fim de semana ou feriado vão para o próximo dia útil</label>
    ${jaLanc.filter(t => !pago(t)).length ? `<label class="chk"><input type="checkbox" name="subst"> Substituir as ${jaLanc.filter(t => !pago(t)).length} conta(s) a receber em aberto deste lote lançadas antes do plano (${brl2(sum(jaLanc.filter(t => !pago(t)), t => t.valor))})</label>` : ''}
    ${jaLanc.length ? `<p class="muted small">Este lote já tem ${brl2(sum(jaLanc.filter(pago), t => t.valor))} recebidos fora do plano.</p>` : ''}
    <div class="preview" id="pPrev"></div>
    <div class="acts">${p.itens && p.itens.length ? '<button type="button" class="btn danger ghost" id="pDel">Excluir plano</button>' : ''}<button type="button" class="btn" data-close>Cancelar</button><button class="btn primary">Salvar e gerar contas a receber</button></div></form>`, { largo: true });
  const F = f.el.querySelector('#fx'); const box = F.querySelector('#pItens');
  const desenhar = () => {
    box.innerHTML = itens.map((it, i) => `<div class="plano-row"><select data-i="${i}" data-k="tipo">${opts(TIPOS_PLANO, it.tipo)}</select><input data-i="${i}" data-k="valor" inputmode="decimal" placeholder="Valor total" value="${it.valor ? fNum2.format(it.valor) : ''}"><input data-i="${i}" data-k="n" type="number" min="1" inputmode="numeric" value="${it.n || 1}" title="Parcelas"><select data-i="${i}" data-k="passo" title="Intervalo">${opts([['1', 'mensal'], ['2', 'bimestral'], ['3', 'trimestral'], ['6', 'semestral'], ['12', 'anual']], String(it.passo || 1))}</select><input data-i="${i}" data-k="venc" type="date" value="${esc(it.venc || hoje())}" title="Primeiro vencimento"><button type="button" class="x" data-rm="${i}" aria-label="Remover">×</button></div>`).join('');
    $$('[data-k]', box).forEach(x => x.oninput = x.onchange = () => { const it = itens[+x.dataset.i]; const kk = x.dataset.k; it[kk] = kk === 'valor' ? parseBR(x.value) || 0 : kk === 'n' || kk === 'passo' ? +x.value || 1 : x.value; prev(); });
    $$('[data-rm]', box).forEach(b => b.onclick = () => { itens.splice(+b.dataset.rm, 1); desenhar(); prev(); });
    prev();
  };
  const plano = () => ({ valor: parseBR(F.valor.value) || 0, itens: itens.filter(x => x.valor > 0), util: F.util.checked, reajuste: parseBR(F.reaj.value) || 0, emp: F.emp ? F.emp.value : '' });
  const prev = () => { const pl = plano(); const soma = sum(pl.itens, x => x.valor); const ps = parcelasPlano(k, pl); const dif = r2(pl.valor - soma);
    $('#pPrev', F).innerHTML = `<b>${ps.length} parcela(s)</b>, total ${brl2(sum(ps, x => x.valor))}${pl.valor ? (Math.abs(dif) >= 0.01 ? ` · <span class="neg">${dif > 0 ? 'faltam' : 'sobram'} ${brl2(Math.abs(dif))} para o preço de venda</span>` : ' · fecha com o preço de venda') : ''}${ps.length ? `<br><span class="muted">Primeira ${dBR(ps.map(x => x.venc).sort()[0])}, última ${dBR(ps.map(x => x.venc).sort().pop())}</span>` : ''}`; };
  F.querySelector('#pAdd').onclick = () => { itens.push({ id: uid().slice(-6), tipo: 'direta', valor: 0, n: 12, passo: 1, venc: addMonths(hoje(), 1) }); desenhar(); };
  F.addEventListener('input', e => { if (!e.target.dataset.k) prev(); });
  desenhar();
  F.onsubmit = e => {
    e.preventDefault(); const pl = plano(); if (!pl.itens.length) { toast('Informe ao menos uma parte com valor.', 'err'); return; }
    if (F.emp && !F.emp.value) { toast('Escolha o CNPJ que recebe.', 'err'); return; }
    const cliente = F.cliente.value.trim(); const ps = parcelasPlano(k, pl); const novosIds = new Set(ps.map(x => x.id));
    const existentes = S.list.filter(t => t.plano === k);
    const apagar = existentes.filter(t => !pago(t) && !novosIds.has(t.id)).map(t => t.id);
    if (F.subst && F.subst.checked) apagar.push(...jaLanc.filter(t => !pago(t)).map(t => t.id));
    const out = ps.filter(x => !(S.byId.get(x.id) && pago(S.byId.get(x.id)))).map(x => {
      const t = { ...(S.byId.get(x.id) || { juros: 0, nf: 0, nfNum: '', status: 'aberto', criado: Date.now() }), id: x.id, item: `${tipoPlanoNome(x.tipo)}${x.n > 1 ? ` (${x.i + 1}/${x.n})` : ''} - ${cliente || 'cliente'} - ${q} ${l}`, valor: x.valor, local: obraDe(q), quadra: q, lote: l, destino: 'Venda', venc: x.venc, forn: cliente, grupo: 'pl' + hashTxt(k), plano: k, planoTipo: x.tipo };
      if (x.vencOrig) t.vencOrig = x.vencOrig; else delete t.vencOrig; if (pl.emp) t.emp = pl.emp; return t; });
    if (apagar.length) excluirTitulos(apagar);
    salvarTitulos(out);
    salvarVendaLote(k, { ...v, cliente: cliente || v.cliente || '', plano: { ...pl, atualizado: Date.now() }, etapa: !v.etapa || v.etapa === 'disp' ? 'contrato' : v.etapa, data: !v.etapa || v.etapa === 'disp' ? hoje() : (v.data || hoje()) });
    f.fechar(); toast(`Plano salvo: ${out.length} conta(s) a receber`); refreshView();
  };
  const del = F.querySelector('#pDel'); if (del) del.onclick = async () => { const ab = S.list.filter(t => t.plano === k && !pago(t)); if (!(await confirmar(`Excluir o plano e as ${ab.length} parcela(s) em aberto? As já recebidas ficam.`, 'Excluir'))) return; excluirTitulos(ab.map(t => t.id)); const nv = { ...v }; delete nv.plano; salvarVendaLote(k, nv); f.fechar(); refreshView(); };
}

/* ---------- PIX: código copia e cola (BR Code estático) e QR ---------- */
const semAcento = s => String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^A-Za-z0-9 .\-]/g, '').toUpperCase();
function crc16(s) { let c = 0xFFFF; for (let i = 0; i < s.length; i++) { c ^= s.charCodeAt(i) << 8; for (let j = 0; j < 8; j++) c = c & 0x8000 ? ((c << 1) ^ 0x1021) & 0xFFFF : (c << 1) & 0xFFFF; } return c.toString(16).toUpperCase().padStart(4, '0'); }
function payloadPix({ chave, nome, cidade, valor, txid, desc }) {
  const f = (id, v) => id + String(v.length).padStart(2, '0') + v;
  const conta = f('00', 'br.gov.bcb.pix') + f('01', chave) + (desc ? f('02', desc.slice(0, 40)) : '');
  let p = f('00', '01') + f('26', conta) + f('52', '0000') + f('53', '986') + (valor ? f('54', valor.toFixed(2)) : '') + f('58', 'BR') + f('59', semAcento(nome).slice(0, 25) || 'SFG') + f('60', semAcento(cidade).slice(0, 15) || 'BRASIL') + f('62', f('05', (txid || '***').replace(/[^A-Za-z0-9]/g, '').slice(0, 25) || '***')) + '6304';
  return p + crc16(p);
}
function svgQR(texto) { const qr = qrcode(0, 'M'); qr.addData(texto); qr.make(); return qr.createSvgTag({ cellSize: 5, margin: 3, scalable: true }); }
function telCliente(t) { const v = vendaLote(t.quadra + '|' + t.lote); return (v.tel || '').replace(/\D/g, ''); }
ACOES.pix = d => {
  const t = S.byId.get(d.id); if (!t) return;
  const e = empPorId(empT(t)) || {}; const chave = (e.pix || '').trim();
  if (!chave) { toast('Cadastre a chave PIX da empresa em Cadastros, Empresas (CNPJ).', 'err'); return; }
  const txid = ('SFG' + t.id).replace(/[^A-Za-z0-9]/g, '').slice(0, 25);
  const code = payloadPix({ chave, nome: e.pixNome || e.nome, cidade: e.pixCidade || 'Juazeiro', valor: Math.abs(t.valor), txid });
  const tel = telCliente(t); const cliente = t.forn || vendaLote(t.quadra + '|' + t.lote).cliente || '';
  const msg = `Olá${cliente ? ' ' + cliente.split(' ')[0] : ''}! Segue a cobrança de ${t.item.replace(/\s*-\s*[^-]+$/, '')}: ${brl2(Math.abs(t.valor))}, vencimento ${dBR(t.venc)}.\n\nPIX copia e cola:\n${code}\n\nSFG Incorporadora`;
  const f = folha('Cobrar por PIX', `<div class="pix"><div class="pix-qr">${svgQR(code)}</div>
    <p><b>${esc(t.item)}</b><br>${brl2(Math.abs(t.valor))} · vence ${dBR(t.venc)}<br><span class="muted small">Recebedor: ${esc(e.pixNome || e.nome)} · chave ${esc(chave)}</span></p>
    <label>PIX copia e cola<textarea readonly rows="3" id="pixCode">${esc(code)}</textarea></label>
    <div class="acts left"><button class="btn primary" id="pCopy">Copiar código</button>${tel ? `<a class="btn" target="_blank" rel="noopener" href="https://wa.me/55${tel.replace(/^55/, '')}?text=${encodeURIComponent(msg)}">Enviar pelo WhatsApp</a>` : `<a class="btn" target="_blank" rel="noopener" href="https://wa.me/?text=${encodeURIComponent(msg)}">Enviar pelo WhatsApp</a>`}</div>
    <p class="muted small">O QR code é um PIX estático com o valor e a identificação deste lançamento. Quando o dinheiro cair, registre o recebimento (ou concilie pelo extrato). ${tel ? '' : 'Cadastre o telefone do cliente em Vendas e repasses para abrir a conversa direto.'}</p></div>`);
  f.el.querySelector('#pCopy').onclick = async () => { try { await navigator.clipboard.writeText(code); toast('Código PIX copiado'); } catch (err) { const x = f.el.querySelector('#pixCode'); x.select(); document.execCommand('copy'); toast('Código PIX copiado'); } };
};
BUSCA_EXTRA.push(q => Object.entries((S.vendas && S.vendas.lotes) || {}).filter(([, v]) => v.cliente && normTxt(v.cliente).includes(q)).slice(0, 5).map(([k, v]) => ({ titulo: 'Cliente · ' + v.cliente, sub: 'Quadra ' + k.replace('|', ', lote '), href: '#/vendas' })));
