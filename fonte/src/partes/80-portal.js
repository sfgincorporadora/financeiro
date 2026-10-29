/* =========================================================
   Link de acompanhamento para o cliente (página pública por lote, só leitura)
   O app publica um resumo em portal/{código}; a página cliente.html lê só esse documento.
   ========================================================= */
usarConfig('portais');
const urlPortal = tk => new URL('cliente.html?c=' + tk, location.href.split('#')[0]).href;
function dadosPortal(k) {
  const [q, l] = k.split('|'); const v = vendaLote(k); const est = estadoFx(q, l); const es = ordemFx(etapasFx());
  const grupo = g => es.filter(e => e.grupo === g).map(e => ({ n: e.nome, s: est[e.id].ok ? 'ok' : est[e.id].st === 'andamento' ? 'andamento' : 'pendente', d: est[e.id].f || est[e.id].i || '' }));
  const obra = es.filter(e => e.grupo === 'Obra'); const okObra = obra.filter(e => est[e.id].ok).length;
  const docs = docsCk().map((d, i) => ({ n: d, ok: estadoDoc(q, l, i).feito }));
  const ts = S.list.filter(t => t.quadra === q && String(t.lote) === String(l) && t.destino === 'Venda' && t.valor > 0);
  const emp = empPorId((ts[0] && empT(ts[0])) || empDaObra(obraDe(q))) || {};
  const abertas = ts.filter(t => !pago(t)).sort((a, b) => a.venc.localeCompare(b.venc)).slice(0, 8).map(t => ({ d: t.venc, v: t.valor, desc: t.item.replace(/\s*-\s*[^-]*-\s*[^-]+$/, ''), pix: emp.pix && !['financiamento', 'fgts', 'subsidio'].includes(t.planoTipo) ? payloadPix({ chave: emp.pix, nome: emp.pixNome || emp.nome, cidade: emp.pixCidade || 'Juazeiro', valor: t.valor, txid: ('SFG' + t.id).replace(/[^A-Za-z0-9]/g, '').slice(0, 25) }) : '' }));
  return { lote: l, quadra: q, obra: obraDe(q), cliente: (v.cliente || '').split(' ')[0], etapa: etapaNome(v.etapa || infoVendaLote(q, l).etapa), etapaData: v.data || '', venda: grupo('Venda e repasse'), obraPct: obra.length ? okObra / obra.length : null, obraAtual: (obra.find(e => !est[e.id].ok) || {}).nome || '', docs, parcelas: abertas, recebido: sum(ts.filter(pago), t => t.valor), empresa: emp.nome || 'SFG Incorporadora' };
}
function publicarPortal(k, forcar) {
  const v = vendaLote(k); if (!v.portal || somenteLeitura()) return false;
  const d = dadosPortal(k); const json = JSON.stringify(d); const h = hashTxt(json);
  if (!forcar && ((S.portais || {})[k] === h)) return false;
  gravar(doc(db, 'portal', v.portal), { json, atualizado: new Date().toISOString() }).catch(onErroEscrita);
  S.portais = { ...(S.portais || {}), [k]: h }; gravar(doc(db, 'config', 'portais'), { [k]: h }, { merge: true }).catch(onErroEscrita);
  return true;
}
/* Mantém os links atualizados: confere a cada mudança (com folga de 20 s). */
let _portT = 0;
function agendarPortais() { clearTimeout(_portT); _portT = setTimeout(() => { if (!S.loaded.titulos || S.portais === undefined || somenteLeitura()) return; for (const k in ((S.vendas && S.vendas.lotes) || {})) if (S.vendas.lotes[k].portal) publicarPortal(k); }, 20000); }
setInterval(() => { if (S.user) agendarPortais(); }, 5 * 60000);
ACOES.portalLote = d => {
  const k = d.k; const [q, l] = k.split('|'); const v = vendaLote(k);
  if (!v.portal) {
    const f = folha(`Link do cliente · quadra ${q}, lote ${l}`, `<p>Gera uma página só de leitura para o comprador acompanhar o lote pelo celular, sem login: etapa do repasse, andamento da obra, documentos e próximas parcelas com PIX.</p><p class="muted small">Quem tiver o link vê apenas este lote. Valores da obra, outros clientes e o restante do sistema não aparecem. Você pode desativar o link quando quiser.</p>
      <div class="acts"><button class="btn" data-close>Cancelar</button><button class="btn primary" id="pGer">Gerar link</button></div>`);
    f.el.querySelector('#pGer').onclick = () => { const tk = Array.from(crypto.getRandomValues(new Uint8Array(16)), b => 'abcdefghjkmnpqrstuvwxyz23456789'[b % 31]).join(''); salvarVendaLote(k, { ...v, portal: tk }); publicarPortal(k, true); f.fechar(); ACOES.portalLote(d); };
    return;
  }
  const url = urlPortal(v.portal); const tel = (v.tel || '').replace(/\D/g, '');
  const msg = `Olá${v.cliente ? ' ' + v.cliente.split(' ')[0] : ''}! Por este link você acompanha o andamento do seu imóvel (quadra ${q}, lote ${l}): ${url}\n\nSFG Incorporadora`;
  const f = folha(`Link do cliente · quadra ${q}, lote ${l}`, `<label>Link<input readonly value="${esc(url)}" id="pUrl"></label>
    <div class="acts left"><button class="btn primary" id="pCp">Copiar link</button><a class="btn" target="_blank" rel="noopener" href="https://wa.me/${tel ? '55' + tel.replace(/^55/, '') : ''}?text=${encodeURIComponent(msg)}">Enviar pelo WhatsApp</a><a class="btn" target="_blank" rel="noopener" href="${esc(url)}">Ver como o cliente vê</a></div>
    <p class="muted small">A página se atualiza sozinha quando você muda a etapa, os documentos, o fluxo do lote ou as parcelas.</p>
    <div class="acts"><button class="btn danger ghost" id="pOff">Desativar link</button><button class="btn" id="pUp">Atualizar agora</button></div>`);
  f.el.querySelector('#pCp').onclick = async () => { try { await navigator.clipboard.writeText(url); } catch (e) { f.el.querySelector('#pUrl').select(); document.execCommand('copy'); } toast('Link copiado'); };
  f.el.querySelector('#pUp').onclick = () => { publicarPortal(k, true); toast('Página do cliente atualizada'); };
  f.el.querySelector('#pOff').onclick = async () => { if (!(await confirmar('Desativar este link? Quem abrir verá "link desativado".', 'Desativar'))) return; const b = writeBatch(db); b.delete(doc(db, 'portal', v.portal)); b.commit().catch(onErroEscrita); const nv = { ...v }; delete nv.portal; salvarVendaLote(k, nv); f.fechar(); toast('Link desativado'); };
};
