import { initializeApp } from 'firebase/app';
import { initializeAuth, indexedDBLocalPersistence, browserLocalPersistence, onAuthStateChanged, signInWithEmailAndPassword, signOut, sendPasswordResetEmail, connectAuthEmulator } from 'firebase/auth';
import { initializeFirestore, persistentLocalCache, persistentMultipleTabManager, connectFirestoreEmulator, collection, doc, onSnapshot, setDoc, getDoc, deleteField, writeBatch } from 'firebase/firestore';
import * as XLSX from 'xlsx';
import Chart from 'chart.js/auto';
import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable';
import { firebaseConfig } from './config.js';

const VERSAO = '__VERSAO__';

/* =========================================================
   Utilidades
   ========================================================= */
const $ = (s, el = document) => el.querySelector(s);
const $$ = (s, el = document) => [...el.querySelectorAll(s)];
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const fBRL = new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL', maximumFractionDigits: 0 });
const fBRL2 = new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL', minimumFractionDigits: 2, maximumFractionDigits: 2 });
const fPct = new Intl.NumberFormat('pt-BR', { style: 'percent', maximumFractionDigits: 1 });
const fNum2 = new Intl.NumberFormat('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const brl = v => fBRL.format(Math.abs(v) < 0.5 ? 0 : v);
const brl2 = v => fBRL2.format(Math.abs(v) < 0.005 ? 0 : v);
const pct = v => v === null || v === undefined || !isFinite(v) ? '—' : fPct.format(v);
function short(v) { const a = Math.abs(v), s = v < 0 ? '−' : ''; if (a >= 1e6) return s + 'R$ ' + (a / 1e6).toLocaleString('pt-BR', { maximumFractionDigits: 2 }) + ' mi'; if (a >= 1e3) return s + 'R$ ' + (a / 1e3).toLocaleString('pt-BR', { maximumFractionDigits: 1 }) + ' mil'; return s + 'R$ ' + a.toLocaleString('pt-BR', { maximumFractionDigits: 0 }); }
const cls = v => v > 0.005 ? 'pos' : v < -0.005 ? 'neg' : '';
const sum = (a, f) => a.reduce((s, x) => s + f(x), 0);
const isoLocal = d => d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
const hoje = () => isoLocal(new Date());
function addDays(iso, n) { const [y, m, d] = iso.split('-').map(Number); return isoLocal(new Date(y, m - 1, d + n)); }
function addMonths(iso, n) { const [y, m, d] = iso.split('-').map(Number); const t = new Date(y, m - 1 + n, 1); const last = new Date(t.getFullYear(), t.getMonth() + 1, 0).getDate(); t.setDate(Math.min(d, last)); return isoLocal(t); }
const dBR = iso => iso ? iso.slice(8, 10) + '/' + iso.slice(5, 7) + '/' + iso.slice(0, 4) : '—';
const MESES = ['janeiro', 'fevereiro', 'março', 'abril', 'maio', 'junho', 'julho', 'agosto', 'setembro', 'outubro', 'novembro', 'dezembro'];
const MES3 = ['jan', 'fev', 'mar', 'abr', 'mai', 'jun', 'jul', 'ago', 'set', 'out', 'nov', 'dez'];
const mesNome = k => k && k.length >= 7 ? MESES[+k.slice(5, 7) - 1] + ' ' + k.slice(0, 4) : 'Sem data';
const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 9);
const lotSort = (a, b) => { const x = parseFloat(a), y = parseFloat(b); if (isNaN(x) || isNaN(y) || x === y) return String(a).localeCompare(String(b), 'pt-BR', { numeric: true }); return x - y; };
const r2 = v => Math.round(v * 100) / 100;
function parseBR(s) {
  s = String(s ?? '').trim().replace(/[R$\s]/g, '');
  if (!s) return NaN;
  if (s.includes(',')) s = s.replace(/\./g, '').replace(',', '.');
  else if ((s.match(/\./g) || []).length > 1 || /\.\d{3}$/.test(s)) s = s.replace(/\./g, '');
  return parseFloat(s);
}
const css = n => getComputedStyle(document.documentElement).getPropertyValue(n).trim();
function toast(msg, tipo) { const t = $('#toast'); t.textContent = msg; t.className = 'on ' + (tipo || ''); clearTimeout(toast._t); toast._t = setTimeout(() => t.className = '', 3500); }

/* =========================================================
   Firebase
   ========================================================= */
const EMU = ['localhost', '127.0.0.1'].includes(location.hostname);
const CONFIGURADO = EMU || (firebaseConfig.apiKey && !/COLE/.test(firebaseConfig.apiKey));
let auth, db;
if (CONFIGURADO) {
  const app = initializeApp(EMU ? { apiKey: 'demo-key', projectId: 'demo-sfg', authDomain: 'demo-sfg.firebaseapp.com', appId: 'demo' } : firebaseConfig);
  auth = initializeAuth(app, { persistence: [indexedDBLocalPersistence, browserLocalPersistence] });
  db = initializeFirestore(app, { localCache: persistentLocalCache({ tabManager: persistentMultipleTabManager() }) });
  if (EMU) { connectAuthEmulator(auth, 'http://127.0.0.1:9099', { disableWarnings: true }); connectFirestoreEmulator(db, '127.0.0.1', 8080); }
}

/* =========================================================
   Estado e modelo
   ========================================================= */
const DEST_PADRAO = {
  op: ['Mão de Obra', 'Material', 'Notas', 'Operacional', 'Pavimentação', 'Prefeitura', 'Registro', 'Terreno'],
  nonop: ['Aditivo', 'Administrativo', 'Caixa', 'Corretagem', 'Impostos'],
  rec: ['Rendimento', 'Venda'],
  aporte: ['Aporte']
};
const SECAO_NOME = { op: 'Despesas operacionais', nonop: 'Despesas não operacionais', rec: 'Receitas', aporte: 'Aportes' };
const DOCS_PADRAO = ['CND IPTU', 'ITBI', 'Termo de Entrega', 'Taxa Vistoria', 'Seguro', 'Certidão Venda', 'Averbação Habite-se', 'Habite-se', 'Escritura', 'Registro Escritura', 'Autorização Escritura', 'Licença de Construção', 'Averbação Construção', 'Procuração', 'ISS', 'ART', 'Projetos', 'SCPO', 'Firma', 'Buscas', 'Síntese', 'Memorial Descritivo'];

const S = {
  user: null, buckets: {}, list: [], byId: new Map(),
  estrutura: null, orcamento: null, checklist: null,
  loaded: { titulos: false, estrutura: false }, pendente: false, online: navigator.onLine, unsubs: []
};
const E = () => {
  const e = S.estrutura || {};
  return {
    local: e.local || 'Nova Juazeiro',
    quadras: e.quadras || [],
    lotes: e.lotes || {},
    destinos: e.destinos || DEST_PADRAO,
    contas: e.contas && e.contas.length ? e.contas : [{ nome: 'BTG', inicial: 0 }, { nome: 'Caixa', inicial: 0 }]
  };
};
function secaoDe(dest) { const d = E().destinos; for (const s of ['op', 'nonop', 'rec', 'aporte']) if ((d[s] || []).includes(dest)) return s; return dest === 'Aporte' ? 'aporte' : 'nonop'; }
const todosDestinos = () => { const d = E().destinos; return [...(d.op || []), ...(d.nonop || []), ...(d.rec || []), ...(d.aporte || [])]; };
const lotesDe = q => (E().lotes[q] || []).slice().sort(lotSort);
const pago = t => t.status === 'pago';
const dataRef = t => pago(t) ? (t.pag || t.venc) : t.venc;

function parseBucket(data) {
  const out = [];
  const items = (data && data.items) || {};
  for (const k in items) { try { const t = JSON.parse(items[k]); if (t && t.id) out.push(t); } catch (e) { /* ignora */ } }
  return out;
}
function rebuildList() {
  const all = [];
  for (const k in S.buckets) for (const t of S.buckets[k]) { t.mes = k; all.push(t); }
  S.list = all;
  S.byId = new Map(all.map(t => [t.id, t]));
}
const bucketDe = t => (t.venc && /^\d{4}-\d{2}/.test(t.venc)) ? t.venc.slice(0, 7) : 'sem-data';
function limpar(t) { const o = { ...t }; delete o.mes; return o; }
function onErroEscrita(e) { console.error(e); toast(e.code === 'permission-denied' ? 'Sem permissão para salvar. Confira o usuário logado.' : 'Erro ao salvar: ' + (e.message || e), 'err'); }

/* Grava vários títulos agrupando por mês (um documento por mês). Não aguarda o servidor: funciona offline. */
function salvarTitulos(lista) {
  if (somenteLeitura()) { toast('Seu acesso é somente leitura.', 'err'); return; }
  const porMes = {}, apagar = {};
  for (const t of lista) {
    t.alterado = Date.now();
    const novo = bucketDe(t);
    const antigo = S.byId.get(t.id)?.mes;
    if (antigo && antigo !== novo) (apagar[antigo] = apagar[antigo] || {})[t.id] = deleteField();
    (porMes[novo] = porMes[novo] || {})[t.id] = JSON.stringify(limpar(t));
  }
  for (const m in apagar) gravar(doc(db, 'titulos', m), { items: apagar[m] }, { merge: true }).catch(onErroEscrita);
  for (const m in porMes) gravar(doc(db, 'titulos', m), { items: porMes[m] }, { merge: true }).catch(onErroEscrita);
}
function excluirTitulos(ids) {
  if (somenteLeitura()) { toast('Seu acesso é somente leitura.', 'err'); return; }
  const porMes = {};
  for (const id of ids) { const t = S.byId.get(id); if (!t) continue; (porMes[t.mes] = porMes[t.mes] || {})[id] = deleteField(); }
  for (const m in porMes) gravar(doc(db, 'titulos', m), { items: porMes[m] }, { merge: true }).catch(onErroEscrita);
}
const DONOS = ['luismatagalvao@gmail.com'];
function papel() {
  const em = (S.user?.email || '').toLowerCase(); const a = S.acesso || {};
  if (DONOS.includes(em)) return 'dono';
  if ((a.editores || []).map(x => x.toLowerCase()).includes(em)) return 'editor';
  if ((a.leitores || []).map(x => x.toLowerCase()).includes(em)) return 'leitor';
  return 'editor';
}
const somenteLeitura = () => papel() === 'leitor';
function gravar(ref, data, opt) {
  if (somenteLeitura()) { toast('Seu acesso é somente leitura.', 'err'); return Promise.resolve(); }
  return opt ? setDoc(ref, data, opt) : setDoc(ref, data);
}
function salvarEstrutura(e) { S.estrutura = e; gravar(doc(db, 'config', 'estrutura'), e).catch(onErroEscrita); refreshView(); }

function saldos() {
  const contas = E().contas.map(c => ({ nome: c.nome, inicial: +c.inicial || 0, saldo: +c.inicial || 0 }));
  const idx = Object.fromEntries(contas.map((c, i) => [c.nome, i]));
  for (const t of S.list) if (pago(t)) {
    let i = idx[t.conta || ''];
    if (i === undefined) { idx[t.conta || ''] = i = contas.length; contas.push({ nome: t.conta || 'Sem conta', inicial: 0, saldo: 0 }); }
    contas[i].saldo += (t.valor || 0) + (t.juros || 0);
  }
  return { contas, total: sum(contas, c => c.saldo) };
}
function projecao() {
  const h = hoje(), base = saldos().total;
  const ev = S.list.filter(t => !pago(t)).map(t => ({ d: t.venc && t.venc > h ? t.venc : h, v: t.valor, t }));
  ev.sort((a, b) => a.d.localeCompare(b.d) || (b.v - a.v));
  const dias = [{ d: h, v: base }]; let bal = base;
  for (const e of ev) { bal += e.v; e.bal = bal; const last = dias[dias.length - 1]; if (last.d === e.d) last.v = bal; else dias.push({ d: e.d, v: bal }); }
  let min = dias[0]; for (const x of dias) if (x.v < min.v) min = x;
  return { dias, ev, min, base };
}
function byDest(rows) { const o = {}; for (const r of rows) { o[r.destino] = (o[r.destino] || 0) + r.valor; if (r.juros) o['Juros e Multa'] = (o['Juros e Multa'] || 0) + r.juros; } return o; }
function secoesViab() { const d = E().destinos; return { op: ['Juros e Multa', ...(d.op || [])], nonop: d.nonop || [], rec: d.rec || [] }; }
function indicadores(t) {
  const sec = secoesViab(); const S_ = k => sum(sec[k], d => t[d] || 0);
  const op = S_('op'), non = S_('nonop'), rec = S_('rec');
  const base = -op - (t['Aditivo'] || 0) - (t['Administrativo'] || 0) - (t['Caixa'] || 0);
  return { op, non, rec, roi: base ? (rec + non + op) / base : null, lop: rec + op, mop: rec ? (rec + op) / rec : null, ll: rec + op + non, ml: rec ? (rec + op + non) / rec : null };
}
function previstoLote(q) { const o = {}; const p = (S.orcamento && S.orcamento.quadras && S.orcamento.quadras[q]) || {}; for (const k in p) o[k] = (secaoDe(k) === 'rec' ? 1 : -1) * (+p[k] || 0); return o; }
function ckPct(q, l) {
  const c = S.checklist; if (!c || !c.docs || !c.docs.length) return null;
  const lots = l === undefined ? lotesDe(q) : [l]; if (!lots.length) return null;
  let ok = 0; for (const x of lots) { const s = (c.lotes || {})[q + '|' + x] || ''; for (let i = 0; i < c.docs.length; i++) if (s[i] === '1') ok++; }
  return ok / (lots.length * c.docs.length);
}
function resumoQuadra(q) {
  const ts = S.list.filter(t => t.quadra === q && t.destino !== 'Aporte');
  const real = ts.filter(pago), ab = ts.filter(t => !pago(t));
  const n = lotesDe(q).length;
  const custo = sum(real, t => Math.min(t.valor, 0) + Math.min(t.juros || 0, 0));
  const receita = sum(real, t => Math.max(t.valor, 0));
  const pagar = sum(ab.filter(t => t.valor < 0), t => t.valor), receber = sum(ab.filter(t => t.valor > 0), t => t.valor);
  const pi = indicadores(previstoLote(q));
  return { q, n, custo, receita, pagar, receber, proj: custo + receita + pagar + receber, prevCusto: (pi.op + pi.non) * n, prevRes: pi.ll * n, cl: ckPct(q) };
}

/* =========================================================
   Sincronização
   ========================================================= */
function iniciarDados() {
  pararDados();
  S.unsubs.push(onSnapshot(collection(db, 'titulos'), { includeMetadataChanges: true }, snap => {
    const ch = snap.docChanges();
    for (const c of ch) { if (c.type === 'removed') delete S.buckets[c.doc.id]; else S.buckets[c.doc.id] = parseBucket(c.doc.data()); }
    const first = !S.loaded.titulos; S.loaded.titulos = true;
    S.pendente = snap.metadata.hasPendingWrites; atualizarSync();
    if (ch.length || first) { rebuildList(); agendarRefresh(); }
  }, e => {
    console.error(e);
    if (e.code === 'permission-denied') { atual.view = null; const v = $('#view'); if (v) v.innerHTML = `<div class="empty card"><h2>Sem acesso</h2><p>O usuário <b>${esc(S.user?.email)}</b> não tem permissão para ver os dados da SFG.</p><button class="btn primary" id="bSairX">Sair e entrar com outro usuário</button></div>`; const b = $('#bSairX'); if (b) b.onclick = () => signOut(auth); }
    else toast('Erro de acesso aos dados: ' + (e.code || e.message), 'err');
  }));
  const cfg = (id, key) => S.unsubs.push(onSnapshot(doc(db, 'config', id), snap => {
    S[key] = snap.exists() ? snap.data() : null; if (key === 'estrutura') S.loaded.estrutura = true; agendarRefresh();
  }, e => console.error(e)));
  cfg('estrutura', 'estrutura'); cfg('orcamento', 'orcamento'); cfg('checklist', 'checklist');
  cfg('vendas', 'vendas'); cfg('obras', 'obras'); cfg('fornecedores', 'fornecedores'); cfg('simulacoes', 'simulacoes');
  S.unsubs.push(onSnapshot(doc(db, 'config', 'acesso'), s => { S.acesso = s.exists() ? s.data() : null; document.body.classList.toggle('ro', somenteLeitura()); agendarRefresh(); }, e => console.warn(e)));
  S.unsubs.push(onSnapshot(collection(db, 'extratos'), s => { S.extratos = {}; s.forEach(d => S.extratos[d.id] = d.data()); agendarRefresh(); }, e => console.warn(e)));
}
function pararDados() { S.unsubs.forEach(u => u()); S.unsubs = []; S.buckets = {}; S.list = []; S.byId = new Map(); S.loaded = { titulos: false, estrutura: false }; }
function atualizarSync() {
  const el = $('#sync'); if (!el) return;
  S.online = navigator.onLine;
  let txt, c;
  if (!S.online) { txt = S.pendente ? 'Offline · alterações guardadas no aparelho' : 'Offline'; c = 'off'; }
  else if (S.pendente) { txt = 'Sincronizando…'; c = 'pend'; }
  else { txt = 'Sincronizado'; c = 'ok'; }
  el.className = 'sync ' + c; el.innerHTML = `<i></i><span>${txt}</span>`;
}
addEventListener('online', atualizarSync); addEventListener('offline', atualizarSync);
let _raf = 0;
function agendarRefresh() { if (_raf) return; _raf = requestAnimationFrame(() => { _raf = 0; refreshView(); }); }

/* =========================================================
   Ícones
   ========================================================= */
const IC = {
  inicio: '<path d="M3 11l9-7 9 7v9a1 1 0 0 1-1 1h-5v-6H9v6H4a1 1 0 0 1-1-1z"/>',
  pagar: '<path d="M12 19V5M5 12l7-7 7 7"/>',
  receber: '<path d="M12 5v14M5 12l7 7 7-7"/>',
  extrato: '<path d="M6 3h12v18l-3-2-3 2-3-2-3 2zM9 8h6M9 12h6"/>',
  obras: '<path d="M3 21h18M5 21V9l7-5 7 5v12M9 21v-6h6v6"/>',
  fluxo: '<path d="M3 17l6-6 4 4 8-8M15 7h6v6"/>',
  checklist: '<path d="M9 11l3 3 8-8M20 12v7a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h9"/>',
  orcamento: '<path d="M4 4h16v16H4zM4 10h16M10 4v16"/>',
  cadastros: '<circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.7 1.7 0 0 0 .3 1.8l.1.1a2 2 0 1 1-2.8 2.8l-.1-.1a1.7 1.7 0 0 0-1.8-.3 1.7 1.7 0 0 0-1 1.5V21a2 2 0 1 1-4 0v-.1a1.7 1.7 0 0 0-1.1-1.5 1.7 1.7 0 0 0-1.8.3l-.1.1a2 2 0 1 1-2.8-2.8l.1-.1a1.7 1.7 0 0 0 .3-1.8 1.7 1.7 0 0 0-1.5-1H3a2 2 0 1 1 0-4h.1a1.7 1.7 0 0 0 1.5-1.1 1.7 1.7 0 0 0-.3-1.8l-.1-.1a2 2 0 1 1 2.8-2.8l.1.1a1.7 1.7 0 0 0 1.8.3H9a1.7 1.7 0 0 0 1-1.5V3a2 2 0 1 1 4 0v.1a1.7 1.7 0 0 0 1 1.5 1.7 1.7 0 0 0 1.8-.3l.1-.1a2 2 0 1 1 2.8 2.8l-.1.1a1.7 1.7 0 0 0-.3 1.8V9a1.7 1.7 0 0 0 1.5 1H21a2 2 0 1 1 0 4h-.1a1.7 1.7 0 0 0-1.5 1z"/>',
  dados: '<ellipse cx="12" cy="5" rx="8" ry="3"/><path d="M4 5v14c0 1.7 3.6 3 8 3s8-1.3 8-3V5M4 12c0 1.7 3.6 3 8 3s8-1.3 8-3"/>',
  mais: '<circle cx="5" cy="12" r="1.5"/><circle cx="12" cy="12" r="1.5"/><circle cx="19" cy="12" r="1.5"/>',
  plus: '<path d="M12 5v14M5 12h14"/>',
  conciliacao: '<path d="M7 7h11l-3-3M17 17H6l3 3"/>',
  relatorios: '<path d="M14 3H6a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V9zM14 3v6h6M8 13h8M8 17h5"/>',
  cronograma: '<path d="M4 20V10M10 20V4M16 20v-7M22 20H2"/>',
  vendas: '<path d="M16 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2M9 11a4 4 0 1 0 0-8 4 4 0 0 0 0 8zM22 21v-2a4 4 0 0 0-3-3.9M16 3.1a4 4 0 0 1 0 7.8"/>',
  simulador: '<path d="M5 3h14v18H5zM8 7h8M8 11h2M12 11h2M8 15h2M12 15h2M8 18h2M12 18h2M16 11v7"/>',
  fornecedores: '<path d="M3 7h13v10H3zM16 10h3l2 3v4h-5M7 20a2 2 0 1 0 0-4 2 2 0 0 0 0 4zM18 20a2 2 0 1 0 0-4 2 2 0 0 0 0 4z"/>',
  usuarios: '<path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2M12 11a4 4 0 1 0 0-8 4 4 0 0 0 0 8z"/>',
  calendario: '<path d="M4 5h16v16H4zM4 9h16M8 3v4M16 3v4"/>',
  x: '<path d="M6 6l12 12M18 6L6 18"/>',
  voltar: '<path d="M15 6l-6 6 6 6"/>',
  clip: '<path d="M21 11l-8.5 8.5a5 5 0 0 1-7-7L14 4a3.5 3.5 0 0 1 5 5l-8.5 8.5a2 2 0 0 1-3-3L15 7"/>'
};
const ic = (n, s = 20) => `<svg width="${s}" height="${s}" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${IC[n]}</svg>`;

const NAV = [
  ['#', 'Financeiro'], ['inicio', 'Início'], ['pagar', 'Contas a pagar'], ['receber', 'Contas a receber'], ['extrato', 'Extrato'], ['fluxo', 'Fluxo de caixa'], ['conciliacao', 'Conciliação'], ['relatorios', 'Relatórios'],
  ['#', 'Obras'], ['obras', 'Quadras e lotes'], ['cronograma', 'Cronograma'], ['vendas', 'Vendas e repasses'], ['checklist', 'Checklist'], ['simulador', 'Simulador'],
  ['#', 'Gestão'], ['fornecedores', 'Fornecedores'], ['cadastros', 'Cadastros'], ['usuarios', 'Usuários'], ['dados', 'Dados e backup']
];
const LOGO_BRANCO = 'icons/logo-branco.svg', LOGO_AZUL = 'icons/logo-azul.svg';
const NAV_MOBILE = [['inicio', 'Início'], ['pagar', 'Pagar'], ['receber', 'Receber'], ['obras', 'Obras'], ['mais', 'Mais']];

/* =========================================================
   Telas de entrada
   ========================================================= */
function telaNaoConfigurado() {
  document.body.innerHTML = `<div class="center"><div class="login card"><img src="${LOGO_AZUL}" alt="SFG Incorporadora" class="login-logo"><h1>Sistema ainda não configurado</h1><p class="muted">Falta colar a configuração do Firebase em <code>src/config.js</code> e gerar o app novamente.</p></div></div>`;
}
function telaLogin(msg) {
  document.body.innerHTML = `<div class="center"><form class="login card" id="fLogin" autocomplete="on">
    <img src="${LOGO_AZUL}" alt="SFG Incorporadora" class="login-logo"><h1>Financeiro</h1><p class="muted">Acesso restrito</p>
    <label>E-mail<input name="email" type="email" autocomplete="username" required></label>
    <label>Senha<input name="senha" type="password" autocomplete="current-password" required></label>
    <p class="erro" id="loginErr">${esc(msg || '')}</p>
    <button class="btn primary full" type="submit">Entrar</button>
    <button class="btn link" type="button" id="esqueci">Esqueci a senha</button>
  </form></div><div id="toast"></div>`;
  $('#fLogin').onsubmit = async e => {
    e.preventDefault(); const f = new FormData(e.target); const b = e.target.querySelector('button[type=submit]'); b.disabled = true; b.textContent = 'Entrando…';
    try { await signInWithEmailAndPassword(auth, f.get('email').trim(), f.get('senha')); }
    catch (err) {
      const m = { 'auth/invalid-credential': 'E-mail ou senha incorretos.', 'auth/wrong-password': 'E-mail ou senha incorretos.', 'auth/user-not-found': 'E-mail ou senha incorretos.', 'auth/too-many-requests': 'Muitas tentativas. Aguarde alguns minutos.', 'auth/network-request-failed': 'Sem internet. O primeiro login neste aparelho precisa de conexão.' }[err.code] || ('Não foi possível entrar (' + err.code + ').');
      $('#loginErr').textContent = m; b.disabled = false; b.textContent = 'Entrar';
    }
  };
  $('#esqueci').onclick = async () => {
    const email = $('#fLogin [name=email]').value.trim(); if (!email) { $('#loginErr').textContent = 'Digite seu e-mail acima e toque de novo em "Esqueci a senha".'; return; }
    try { await sendPasswordResetEmail(auth, email); toast('Enviamos um link de redefinição para ' + email); } catch (e) { $('#loginErr').textContent = 'Não foi possível enviar o e-mail (' + e.code + ').'; }
  };
}

/* =========================================================
   Estrutura da aplicação
   ========================================================= */
function montarShell() {
  document.body.innerHTML = `<div class="app">
    <aside class="side"><div class="brand"><img class="logo-img" src="${LOGO_BRANCO}" alt=""><div><b>SFG</b><small>INCORPORADORA</small></div></div>
      <nav>${NAV.map(([k, n]) => k === '#' ? `<div class="sec">${n.toUpperCase()}</div>` : `<a href="#/${k}" data-nav="${k}">${ic(k, 18)}<span>${n}</span></a>`).join('')}</nav>
      <div class="side-foot"><span id="sideUser" class="muted"></span></div></aside>
    <div class="main">
      <header class="top"><button class="icon-btn back hidden" id="btnBack" aria-label="Voltar">${ic('voltar')}</button><img class="tlogo" src="${LOGO_BRANCO}" alt="SFG"><h1 id="title"></h1><span id="sync" class="sync"></span><button class="btn primary desk rw" data-act="novo">${ic('plus', 18)} Novo lançamento</button></header>
      <div id="view" class="view"></div>
    </div>
    <nav class="bottom">${NAV_MOBILE.map(([k, n]) => `<a href="#/${k}" data-nav="${k}">${ic(k, 22)}<span>${n}</span></a>`).join('')}</nav>
    <button class="fab rw" data-act="novo" aria-label="Novo lançamento">${ic('plus', 26)}</button>
  </div><div id="sheets"></div><div id="toast"></div>`;
  $('#sideUser').textContent = S.user?.email || ''; document.body.classList.toggle('ro', somenteLeitura());
  $('#btnBack').onclick = () => history.back();
  atualizarSync();
  document.addEventListener('click', onClick);
  roteador();
}
addEventListener('hashchange', () => { if (S.user) roteador(); });

let atual = { nome: null, args: [], view: null };
function roteador() {
  const parts = (location.hash.replace(/^#\/?/, '') || 'inicio').split('/').map(decodeURIComponent);
  const nome = VIEWS[parts[0]] ? parts[0] : 'inicio';
  const args = parts.slice(1);
  $$('[data-nav]').forEach(a => {
    const k = a.dataset.nav;
    const on = k === nome || (k === 'mais' && !['inicio', 'pagar', 'receber', 'obras'].includes(nome));
    a.classList.toggle('on', on);
  });
  const v = VIEWS[nome];
  $('#title').textContent = typeof v.titulo === 'function' ? v.titulo(args) : v.titulo;
  $('#btnBack').classList.toggle('hidden', !args.length);
  if (atual.view && atual.view.sair) atual.view.sair();
  atual = { nome, args, view: v };
  const el = $('#view'); el.innerHTML = ''; el.scrollTop = 0; window.scrollTo(0, 0);
  if (v.montar) v.montar(el, args);
  if (S.loaded.titulos) v.refresh(el, args); else el.insertAdjacentHTML('beforeend', '<div class="loading" id="carregando">Carregando dados…</div>');
}
function refreshView() {
  if (!atual.view) return;
  const c = $('#carregando'); if (c && S.loaded.titulos) c.remove();
  if (S.loaded.titulos) atual.view.refresh($('#view'), atual.args);
}

/* =========================================================
   Ações (cliques)
   ========================================================= */
const ACOES = {};
function onClick(e) {
  const a = e.target.closest('[data-act]'); if (!a) return;
  const f = ACOES[a.dataset.act]; if (!f) return;
  e.preventDefault(); f(a.dataset, a, e);
}
ACOES.novo = () => {
  const n = atual.nome;
  formNovo({ tipo: n === 'receber' ? 'entrada' : 'saida', realizado: n === 'extrato', quadra: n === 'obras' && atual.args[0] ? atual.args[0] : '', lote: n === 'obras' && atual.args[1] ? atual.args[1] : '' });
};
ACOES.abrir = d => { const t = S.byId.get(d.id); if (t) detalhe(t); };
ACOES.baixar = d => formBaixa([d.id]);
ACOES.ir = d => { location.hash = d.href; };

/* =========================================================
   Folhas (modais)
   ========================================================= */
function folha(titulo, html, opts = {}) {
  const w = document.createElement('div'); w.className = 'sheet-wrap';
  w.innerHTML = `<div class="sheet ${opts.largo ? 'wide' : ''}" role="dialog" aria-modal="true"><header><h2>${esc(titulo)}</h2><button class="icon-btn" data-close aria-label="Fechar">${ic('x')}</button></header><div class="sheet-body">${html}</div></div>`;
  const fechar = () => { w.classList.add('out'); setTimeout(() => w.remove(), 180); document.removeEventListener('keydown', esc_); if (opts.aoFechar) opts.aoFechar(); };
  const esc_ = ev => { if (ev.key === 'Escape' && w === $('#sheets').lastElementChild) fechar(); };
  w.addEventListener('click', ev => { if (ev.target === w || ev.target.closest('[data-close]')) fechar(); });
  document.addEventListener('keydown', esc_);
  $('#sheets').appendChild(w);
  requestAnimationFrame(() => w.classList.add('in'));
  return { el: w, fechar };
}
function confirmar(msg, ok = 'Confirmar', perigo = true) {
  return new Promise(res => {
    let r = false;
    const f = folha('Confirmação', `<p>${msg}</p><div class="acts"><button class="btn" data-close>Cancelar</button><button class="btn ${perigo ? 'danger' : 'primary'}" id="cfOk">${esc(ok)}</button></div>`, { aoFechar: () => res(r) });
    f.el.querySelector('#cfOk').onclick = () => { r = true; f.fechar(); };
  });
}
const opts = (arr, sel, vazio) => (vazio !== undefined ? `<option value="">${esc(vazio)}</option>` : '') + arr.map(x => { const [v, l] = Array.isArray(x) ? x : [x, x]; return `<option value="${esc(v)}" ${String(v) === String(sel) ? 'selected' : ''}>${esc(l)}</option>`; }).join('');
const optsDestino = sel => ['op', 'nonop', 'rec', 'aporte'].map(s => { const ds = E().destinos[s] || []; return ds.length ? `<optgroup label="${SECAO_NOME[s]}">${opts(ds, sel)}</optgroup>` : ''; }).join('');
const quadrasOpts = () => [['-', 'Geral / administrativo (-)'], ...E().quadras.map(q => [q, 'Quadra ' + q])];

/* ---------- Novo lançamento ---------- */
function formNovo(def = {}) {
  const h = hoje();
  const f = folha('Novo lançamento', `<form id="fNovo" class="form">
    <div class="seg full" id="segTipo"><button type="button" data-v="saida" class="${def.tipo !== 'entrada' ? 'on' : ''}">Saída (pagar)</button><button type="button" data-v="entrada" class="${def.tipo === 'entrada' ? 'on' : ''}">Entrada (receber)</button></div>
    <label class="chk"><input type="checkbox" name="realizado" ${def.realizado ? 'checked' : ''}> Já foi pago/recebido (lança direto no extrato)</label>
    <label>Descrição<input name="item" required placeholder="Ex.: Mendonça NF, Parcela terreno, Venda" list="dlItens" value="${esc(def.item || '')}"></label>
    <div class="g2"><label>Fornecedor / cliente<input name="forn" list="dlForn" value="${esc(def.forn || '')}"></label>
      <label>Destino<select name="destino" required>${optsDestino(def.destino || (def.tipo === 'entrada' ? 'Venda' : 'Material'))}</select></label></div>
    <div class="g2"><label>Valor (R$)<input name="valor" inputmode="decimal" required placeholder="0,00" value="${def.valor ? fNum2.format(Math.abs(def.valor)) : ''}"></label>
      <label>O valor informado é<select name="modo"><option value="total">Total (dividir entre lotes e parcelas)</option><option value="cada">De cada lançamento</option></select></label></div>
    <div class="g2"><label>Quadra<select name="quadra">${opts(quadrasOpts(), def.quadra || '-')}</select></label>
      <label>Vencimento${'<input type="date" name="venc" required value="' + (def.venc || h) + '">'}</label></div>
    <div id="boxLotes"></div>
    <div class="g2"><label>Parcelas<input name="parcelas" type="number" min="1" max="240" value="1" inputmode="numeric"></label>
      <label>Intervalo<select name="intervalo"><option value="1">Mensal</option><option value="0.5">Quinzenal</option><option value="7d">Semanal</option></select></label></div>
    <div id="boxReal" class="g2"><label>Data do pagamento<input type="date" name="pag" value="${def.pag || h}"></label><label>Conta<select name="conta">${opts(E().contas.map(c => c.nome), def.conta)}</select></label></div>
    <div class="g2"><label>Quantidade <span class="muted">(opcional, para histórico de preço)</span><input name="qtd" inputmode="decimal" placeholder="Ex.: 200"></label><label>Unidade<input name="unid" list="dlUnid" placeholder="sc, m³, un, kg"></label></div><datalist id="dlUnid"><option value="sc"><option value="m³"><option value="m²"><option value="un"><option value="kg"><option value="t"><option value="milheiro"><option value="m"></datalist>
    <label class="chk"><input type="checkbox" name="nf"> Tem nota fiscal</label>
    <label>Observação<textarea name="obs" rows="2"></textarea></label>
    <label>Anexo (foto do boleto ou da nota)<input type="file" name="anexo" accept="image/*,application/pdf"></label>
    <div class="preview" id="prevNovo"></div>
    <div class="acts"><button type="button" class="btn" data-close>Cancelar</button><button class="btn primary" type="submit">Salvar</button></div>
    <datalist id="dlItens">${[...new Set(S.list.slice(-3000).map(t => t.item.replace(/\s*[-–]\s*[A-Z]\d?\s*[\d.]*\s*$/, '').replace(/\s*\(\d+\/\d+\)\s*/, ' ').trim()))].slice(0, 300).map(x => `<option value="${esc(x)}">`).join('')}</datalist>
    <datalist id="dlForn">${[...new Set(S.list.map(t => t.forn).filter(Boolean))].map(x => `<option value="${esc(x)}">`).join('')}</datalist>
  </form>`);
  const F = f.el.querySelector('#fNovo');
  let tipo = def.tipo === 'entrada' ? 'entrada' : 'saida';
  $$('#segTipo button', F).forEach(b => b.onclick = () => { tipo = b.dataset.v; $$('#segTipo button', F).forEach(x => x.classList.toggle('on', x === b)); F.destino.value = tipo === 'entrada' ? 'Venda' : 'Material'; prev(); });
  const lotesBox = () => {
    const q = F.quadra.value, ls = q === '-' ? [] : lotesDe(q);
    $('#boxLotes', F).innerHTML = ls.length ? `<div class="lbl">Lotes <span class="muted">(nenhum marcado = custo geral da quadra)</span> <button type="button" class="btn link sm" data-lt="all">Todos</button><button type="button" class="btn link sm" data-lt="none">Limpar</button></div>
      <div class="chips">${ls.map(l => `<label class="pick"><input type="checkbox" name="lotes" value="${esc(l)}" ${String(def.lote) === String(l) ? 'checked' : ''}><span>${esc(l)}</span></label>`).join('')}</div>` : '';
    $$('[data-lt]', F).forEach(b => b.onclick = () => { $$('[name=lotes]', F).forEach(c => c.checked = b.dataset.lt === 'all'); prev(); });
  };
  const real = () => { $('#boxReal', F).classList.toggle('hidden', !F.realizado.checked); };
  const gerar = () => {
    const v = parseBR(F.valor.value); const n = Math.max(1, Math.min(240, parseInt(F.parcelas.value) || 1));
    const lotes = $$('[name=lotes]:checked', F).map(c => c.value); const L = lotes.length ? lotes : ['-'];
    const total = L.length * n;
    if (!(v > 0)) return { erro: 'Informe um valor maior que zero.', total };
    let valores;
    if (F.modo.value === 'cada') valores = Array(total).fill(r2(v));
    else { const cents = Math.round(v * 100); const base = Math.floor(cents / total); let resto = cents - base * total; valores = Array.from({ length: total }, () => (base + (resto-- > 0 ? 1 : 0)) / 100); }
    const q = F.quadra.value, base = F.item.value.trim(), venc0 = F.venc.value, iv = F.intervalo.value;
    const out = []; let k = 0; const grupo = total > 1 ? uid() : '';
    for (let i = 0; i < n; i++) {
      const venc = iv === '7d' ? addDays(venc0, 7 * i) : iv === '0.5' ? (i % 2 ? addDays(addMonths(venc0, Math.floor(i / 2)), 15) : addMonths(venc0, i / 2)) : addMonths(venc0, i);
      for (const l of L) {
        let item = base + (n > 1 ? ` (${i + 1}/${n})` : '') + (l !== '-' ? ` - ${q} ${l}` : (q !== '-' ? ` - ${q}` : ''));
        const valor = (tipo === 'saida' ? -1 : 1) * valores[k++];
        const t = { id: uid(), item, valor, juros: 0, nf: F.nf.checked ? 1 : 0, local: E().local, quadra: q, lote: l, destino: F.destino.value, venc, status: 'aberto', forn: F.forn.value.trim(), obs: F.obs.value.trim(), grupo, criado: Date.now() };
        if (F.realizado.checked) { t.status = 'pago'; t.pag = n > 1 ? venc : (F.pag.value || venc); t.conta = F.conta.value; }
        const qt = parseBR(F.qtd.value); if (qt > 0) { t.qtd = Math.round((F.modo.value === 'cada' ? qt : qt / total) * 1000) / 1000; t.unid = F.unid.value.trim(); }
        if (def.conc && total === 1) t.conc = def.conc;
        out.push(t);
      }
    }
    return { lista: out, total };
  };
  const prev = () => {
    const g = gerar(); const p = $('#prevNovo', F);
    if (!F.item.value.trim()) { p.innerHTML = ''; return; }
    if (g.erro) { p.innerHTML = `<span class="muted">${g.total} lançamento(s)</span>`; return; }
    const tot = sum(g.lista, t => t.valor);
    p.innerHTML = `<b>${g.total} lançamento${g.total > 1 ? 's' : ''}</b> · total <b class="${cls(tot)}">${brl2(tot)}</b><br><span class="muted">Ex.: ${esc(g.lista[0].item)} · ${brl2(g.lista[0].valor)} · vence ${dBR(g.lista[0].venc)}${g.total > 1 ? ` … último ${dBR(g.lista[g.lista.length - 1].venc)}` : ''}</span>`;
  };
  F.quadra.onchange = () => { lotesBox(); prev(); };
  F.realizado.onchange = () => { real(); prev(); };
  F.addEventListener('input', prev); F.addEventListener('change', prev);
  lotesBox(); real(); prev();
  F.onsubmit = async e => {
    e.preventDefault(); const g = gerar(); if (g.erro) { toast(g.erro, 'err'); return; }
    if (g.total > 60 && !(await confirmar(`Isso vai criar <b>${g.total}</b> lançamentos. Continuar?`, 'Criar', false))) return;
    const arq = F.anexo.files[0];
    if (arq) { try { const aid = await salvarAnexo(arq); g.lista.forEach(t => t.anexo = aid); } catch (err) { toast(err.message, 'err'); return; } }
    salvarTitulos(g.lista); f.fechar(); toast(g.total > 1 ? `${g.total} lançamentos criados` : 'Lançamento salvo');
  };
}

/* ---------- Editar ---------- */
function formEditar(t) {
  const lots = t.quadra && t.quadra !== '-' ? lotesDe(t.quadra) : [];
  const f = folha('Editar lançamento', `<form id="fEd" class="form">
    <div class="seg full" id="segTipo"><button type="button" data-v="saida" class="${t.valor < 0 ? 'on' : ''}">Saída</button><button type="button" data-v="entrada" class="${t.valor >= 0 ? 'on' : ''}">Entrada</button></div>
    <label>Descrição<input name="item" required value="${esc(t.item)}"></label>
    <div class="g2"><label>Fornecedor / cliente<input name="forn" value="${esc(t.forn || '')}"></label><label>Destino<select name="destino">${optsDestino(t.destino)}${todosDestinos().includes(t.destino) ? '' : `<option selected>${esc(t.destino)}</option>`}</select></label></div>
    <div class="g2"><label>Valor (R$)<input name="valor" inputmode="decimal" required value="${fNum2.format(Math.abs(t.valor))}"></label><label>Vencimento<input type="date" name="venc" required value="${esc(t.venc)}"></label></div>
    <div class="g2"><label>Quadra<select name="quadra">${opts(quadrasOpts(), t.quadra || '-')}${(t.quadra && t.quadra !== '-' && !E().quadras.includes(t.quadra)) ? `<option selected>${esc(t.quadra)}</option>` : ''}</select></label>
      <label>Lote<select name="lote">${opts([['-', 'Sem lote (-)'], ...lots], t.lote || '-')}${(t.lote && t.lote !== '-' && !lots.includes(t.lote)) ? `<option selected>${esc(t.lote)}</option>` : ''}</select></label></div>
    ${pago(t) ? `<div class="g3"><label>Pago em<input type="date" name="pag" value="${esc(t.pag || t.venc)}"></label><label>Conta<select name="conta">${opts(E().contas.map(c => c.nome), t.conta)}</select></label><label>Juros/multa (R$)<input name="juros" inputmode="decimal" value="${t.juros ? fNum2.format(Math.abs(t.juros)) : ''}"></label></div>` : ''}
    <div class="g2"><label>Quantidade<input name="qtd" inputmode="decimal" value="${t.qtd ? String(t.qtd).replace('.', ',') : ''}"></label><label>Unidade<input name="unid" value="${esc(t.unid || '')}"></label></div>
    <label class="chk"><input type="checkbox" name="nf" ${t.nf ? 'checked' : ''}> Tem nota fiscal</label>
    <label>Observação<textarea name="obs" rows="2">${esc(t.obs || '')}</textarea></label>
    <div class="acts"><button type="button" class="btn" data-close>Cancelar</button><button class="btn primary" type="submit">Salvar</button></div></form>`);
  const F = f.el.querySelector('#fEd'); let tipo = t.valor < 0 ? 'saida' : 'entrada';
  $$('#segTipo button', F).forEach(b => b.onclick = () => { tipo = b.dataset.v; $$('#segTipo button', F).forEach(x => x.classList.toggle('on', x === b)); });
  F.quadra.onchange = () => { const q = F.quadra.value; F.lote.innerHTML = opts([['-', 'Sem lote (-)'], ...(q === '-' ? [] : lotesDe(q))], '-'); };
  F.onsubmit = e => {
    e.preventDefault(); const v = parseBR(F.valor.value); if (!(v >= 0)) { toast('Valor inválido', 'err'); return; }
    const sg = tipo === 'saida' ? -1 : 1;
    const n = { ...t, item: F.item.value.trim(), forn: F.forn.value.trim(), destino: F.destino.value, valor: sg * v, venc: F.venc.value, quadra: F.quadra.value, lote: F.lote.value, nf: F.nf.checked ? 1 : 0, obs: F.obs.value.trim() };
    { const qt = parseBR(F.qtd.value); if (qt > 0) { n.qtd = qt; n.unid = F.unid.value.trim(); } else { delete n.qtd; delete n.unid; } }
    if (pago(t)) { n.pag = F.pag.value; n.conta = F.conta.value; const j = parseBR(F.juros.value); n.juros = j > 0 ? sg * j : 0; }
    salvarTitulos([n]); f.fechar(); toast('Alteração salva');
  };
}

/* ---------- Baixa ---------- */
function formBaixa(ids) {
  const ts = ids.map(id => S.byId.get(id)).filter(t => t && !pago(t)); if (!ts.length) return;
  const um = ts.length === 1, t0 = ts[0], tot = sum(ts, t => t.valor);
  const ent = tot > 0;
  const f = folha(um ? (ent ? 'Registrar recebimento' : 'Registrar pagamento') : `Dar baixa em ${ts.length} lançamentos`, `<form id="fBx" class="form">
    <div class="resumo">${um ? `<b>${esc(t0.item)}</b><br><span class="muted">Vencimento ${dBR(t0.venc)}</span>` : `<b>${ts.length} lançamentos</b>`}<div class="big ${cls(tot)}">${brl2(tot)}</div></div>
    <div class="g2"><label>Data do ${ent ? 'recebimento' : 'pagamento'}<input type="date" name="pag" required value="${hoje()}"></label><label>Conta<select name="conta">${opts(E().contas.map(c => c.nome))}</select></label></div>
    ${um ? `<div class="g2"><label>Valor ${ent ? 'recebido' : 'pago'} (sem juros)<input name="valor" inputmode="decimal" value="${fNum2.format(Math.abs(t0.valor))}"></label><label>Juros / multa (R$)<input name="juros" inputmode="decimal" placeholder="0,00"></label></div>` : `<label>Juros / multa total (R$) <span class="muted">— rateado proporcionalmente</span><input name="juros" inputmode="decimal" placeholder="0,00"></label>`}
    <div class="acts"><button type="button" class="btn" data-close>Cancelar</button><button class="btn primary" type="submit">Confirmar baixa</button></div></form>`);
  const F = f.el.querySelector('#fBx');
  F.onsubmit = e => {
    e.preventDefault(); const j = parseBR(F.juros.value) || 0;
    const out = ts.map(t => {
      const sg = t.valor < 0 ? -1 : 1; const n = { ...t, status: 'pago', pag: F.pag.value, conta: F.conta.value };
      if (um) { const v = parseBR(F.valor.value); if (v >= 0) n.valor = sg * v; n.juros = j > 0 ? sg * j : 0; }
      else n.juros = j > 0 ? sg * r2(j * Math.abs(t.valor) / Math.abs(tot || 1)) : 0;
      return n;
    });
    salvarTitulos(out); f.fechar(); listaSel.clear(); toast(um ? 'Baixa registrada' : `${out.length} baixas registradas`); refreshView();
  };
}

/* ---------- Detalhe ---------- */
function detalhe(t) {
  const venc = !pago(t) && t.venc < hoje();
  const grupo = t.grupo ? S.list.filter(x => x.grupo === t.grupo) : [];
  const f = folha(pago(t) ? (t.valor < 0 ? 'Pagamento realizado' : 'Recebimento realizado') : (t.valor < 0 ? 'Conta a pagar' : 'Conta a receber'), `
    <div class="resumo"><b>${esc(t.item)}</b><div class="big ${cls(t.valor)}">${brl2(t.valor)}</div>
      ${pago(t) ? `<span class="chip ok">Pago em ${dBR(t.pag)} · ${esc(t.conta || '')}</span>` : venc ? `<span class="chip bad">Vencida em ${dBR(t.venc)}</span>` : `<span class="chip warn">Vence em ${dBR(t.venc)}</span>`}</div>
    <dl class="dl">
      <dt>Destino</dt><dd>${esc(t.destino)}</dd><dt>Quadra · lote</dt><dd>${esc(t.quadra || '-')} · ${esc(t.lote || '-')}</dd>
      <dt>Vencimento</dt><dd>${dBR(t.venc)}</dd>${t.juros ? `<dt>Juros / multa</dt><dd class="neg">${brl2(t.juros)}</dd>` : ''}
      ${t.forn ? `<dt>Fornecedor</dt><dd>${esc(t.forn)}</dd>` : ''}<dt>Nota fiscal</dt><dd>${t.nf ? 'Sim' : 'Não'}</dd>
      ${t.obs ? `<dt>Observação</dt><dd>${esc(t.obs)}</dd>` : ''}${grupo.length > 1 ? `<dt>Grupo</dt><dd>${grupo.length} lançamentos (${grupo.filter(pago).length} pagos)</dd>` : ''}
    </dl>
    <div id="anexoBox">${t.anexo ? `<button class="btn" id="verAnexo">${ic('clip', 16)} Ver anexo</button>` : `<label class="btn file">${ic('clip', 16)} Anexar foto ou PDF<input type="file" id="addAnexo" accept="image/*,application/pdf" hidden></label>`}</div>
    <div class="acts wrap">
      ${pago(t) ? `<button class="btn" id="dEst">Estornar baixa</button>` : `<button class="btn primary" id="dBx">${t.valor < 0 ? 'Registrar pagamento' : 'Registrar recebimento'}</button>`}
      <button class="btn" id="dEd">Editar</button><button class="btn" id="dDup">Duplicar</button><button class="btn danger" id="dDel">Excluir</button>
      ${grupo.filter(x => !pago(x)).length > 1 ? `<button class="btn danger ghost" id="dDelG">Excluir ${grupo.filter(x => !pago(x)).length} em aberto do grupo</button>` : ''}
    </div>`);
  const on = (id, fn) => { const b = f.el.querySelector('#' + id); if (b) b.onclick = fn; };
  on('dBx', () => { f.fechar(); formBaixa([t.id]); });
  on('dEst', async () => { if (await confirmar('Estornar a baixa? O lançamento volta para "em aberto" e sai do extrato.', 'Estornar')) { const n = { ...t, status: 'aberto', juros: 0 }; delete n.pag; delete n.conta; salvarTitulos([n]); f.fechar(); toast('Baixa estornada'); } });
  on('dEd', () => { f.fechar(); formEditar(t); });
  on('dDup', () => { f.fechar(); formNovo({ tipo: t.valor < 0 ? 'saida' : 'entrada', item: t.item.replace(/\s*-\s*\S+\s+\S+$/, ''), forn: t.forn, destino: t.destino, valor: t.valor, quadra: t.quadra, lote: t.lote, venc: addMonths(t.venc, 1) }); });
  on('dDel', async () => { if (await confirmar(`Excluir <b>${esc(t.item)}</b>?`, 'Excluir')) { excluirTitulos([t.id]); f.fechar(); toast('Lançamento excluído'); } });
  on('dDelG', async () => { const ab = grupo.filter(x => !pago(x)); if (await confirmar(`Excluir os <b>${ab.length}</b> lançamentos em aberto deste grupo? Os já pagos ficam.`, 'Excluir')) { excluirTitulos(ab.map(x => x.id)); f.fechar(); toast(ab.length + ' lançamentos excluídos'); } });
  on('verAnexo', () => verAnexo(t.anexo));
  const inp = f.el.querySelector('#addAnexo');
  if (inp) inp.onchange = async () => { try { const aid = await salvarAnexo(inp.files[0]); salvarTitulos([{ ...t, anexo: aid }]); f.fechar(); toast('Anexo salvo'); } catch (e) { toast(e.message, 'err'); } };
}

/* ---------- Anexos ---------- */
async function salvarAnexo(file) {
  let mime = file.type || 'application/octet-stream', b64;
  if (mime.startsWith('image/')) {
    const img = await new Promise((res, rej) => { const i = new Image(); i.onload = () => res(i); i.onerror = () => rej(new Error('Não consegui abrir a imagem')); i.src = URL.createObjectURL(file); });
    let w = img.naturalWidth, h = img.naturalHeight, max = 1600, q = 0.72;
    for (let tent = 0; tent < 5; tent++) {
      const s = Math.min(1, max / Math.max(w, h)); const c = document.createElement('canvas'); c.width = Math.round(w * s); c.height = Math.round(h * s);
      c.getContext('2d').drawImage(img, 0, 0, c.width, c.height); b64 = c.toDataURL('image/jpeg', q).split(',')[1];
      if (b64.length < 700000) break; max *= 0.75; q -= 0.08;
    }
    mime = 'image/jpeg';
  } else {
    if (file.size > 650000) throw new Error('PDF muito grande (máx. ~650 KB). Tire uma foto em vez disso.');
    b64 = await new Promise(res => { const r = new FileReader(); r.onload = () => res(String(r.result).split(',')[1]); r.readAsDataURL(file); });
  }
  const id = uid();
  gravar(doc(db, 'anexos', id), { b64, mime, nome: file.name || '', criado: Date.now() }).catch(onErroEscrita);
  return id;
}
async function verAnexo(id) {
  try {
    const s = await getDoc(doc(db, 'anexos', id)); if (!s.exists()) { toast('Anexo ainda não disponível neste aparelho (precisa de internet).', 'err'); return; }
    const a = s.data(); const url = 'data:' + a.mime + ';base64,' + a.b64;
    folha('Anexo', a.mime.startsWith('image/') ? `<img src="${url}" class="anexo-img" alt="Anexo">` : `<p><a class="btn" href="${url}" download="${esc(a.nome || 'anexo.pdf')}">Abrir PDF</a></p>`, { largo: true });
  } catch (e) { toast('Não foi possível abrir o anexo: ' + (e.code || e.message), 'err'); }
}

/* =========================================================
   Componentes de lista
   ========================================================= */
const listaSel = new Set();
function linhaTitulo(t, o = {}) {
  const h = hoje(); const venc = !pago(t) && t.venc < h; const hj = !pago(t) && t.venc === h;
  const d = o.data === 'pag' ? (t.pag || t.venc) : t.venc;
  const meta = [t.quadra && t.quadra !== '-' ? t.quadra + (t.lote && t.lote !== '-' ? ' · ' + t.lote : '') : 'Geral', t.destino, t.forn, t.nf ? 'NF' : '', t.anexo ? '📎' : ''].filter(Boolean).map(esc).join(' · ');
  return `<div class="tr ${o.sel ? 'selmode' : ''} ${listaSel.has(t.id) ? 'sel' : ''}" data-act="${o.sel ? 'tsel' : 'abrir'}" data-id="${t.id}">
    ${o.sel ? `<span class="cb">${listaSel.has(t.id) ? '✓' : ''}</span>` : ''}
    <div class="td-date"><b>${d ? d.slice(8, 10) + '/' + d.slice(5, 7) : '—'}</b><small>${d ? d.slice(0, 4) : ''}</small></div>
    <div class="td-main"><div class="td-item">${esc(t.item)}</div><div class="td-meta">${meta}</div></div>
    <div class="td-val"><b class="${cls(t.valor)}">${brl2(t.valor)}</b>${t.juros ? `<small class="neg">juros ${brl2(t.juros)}</small>` : ''}${o.saldo !== undefined ? `<small class="muted">saldo ${brl(o.saldo)}</small>` : ''}
      ${!pago(t) ? (venc ? '<span class="chip bad">vencida</span>' : hj ? '<span class="chip warn">hoje</span>' : '') : (o.mostrarPago ? '<span class="chip ok">pago</span>' : '')}</div>
    ${o.baixa && !pago(t) ? `<button class="btn sm" data-act="baixar" data-id="${t.id}">Baixar</button>` : ''}
  </div>`;
}
ACOES.tsel = d => { if (listaSel.has(d.id)) listaSel.delete(d.id); else listaSel.add(d.id); refreshView(); };

/* Agrupa lançamentos rateados por lote (mesma descrição, data, quadra, destino e situação). */
const reEsc = s => String(s).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
function baseItem(t) { if (!t.lote || t.lote === '-') return t.item; return t.item.replace(new RegExp('\\s*[-–]\\s*' + reEsc(t.quadra) + '\\s+' + reEsc(t.lote) + '\\s*$'), ''); }
function agrupar(rows, campo) {
  const out = [], idx = new Map();
  for (const t of rows) {
    if (!t.lote || t.lote === '-') { out.push({ t }); continue; }
    const d = campo === 'pag' ? (t.pag || t.venc) : t.venc;
    const k = [t.status, d, t.quadra, t.destino, Math.sign(t.valor), baseItem(t)].join('|');
    let g = idx.get(k);
    if (!g) { g = { itens: [] }; idx.set(k, g); out.push(g); }
    g.itens.push(t);
  }
  return out.map(g => g.t ? g : (g.itens.length === 1 ? { t: g.itens[0] } : g));
}
function linhaGrupo(g, o = {}) {
  const t = g.itens[0], ids = g.itens.map(x => x.id), tot = sum(g.itens, x => x.valor), jur = sum(g.itens, x => x.juros || 0);
  const h = hoje(); const venc = !pago(t) && t.venc < h; const hj = !pago(t) && t.venc === h;
  const d = o.data === 'pag' ? (t.pag || t.venc) : t.venc;
  const ls = g.itens.map(x => x.lote).sort(lotSort);
  const allSel = ids.every(id => listaSel.has(id));
  return `<div class="tr grp ${o.sel ? 'selmode' : ''} ${o.sel && allSel ? 'sel' : ''}" data-act="${o.sel ? 'gsel' : 'grp'}" data-ids="${ids.join(',')}">
    ${o.sel ? `<span class="cb">${allSel ? '✓' : ''}</span>` : ''}
    <div class="td-date"><b>${d ? d.slice(8, 10) + '/' + d.slice(5, 7) : '—'}</b><small>${d ? d.slice(0, 4) : ''}</small></div>
    <div class="td-main"><div class="td-item">${esc(baseItem(t))} <span class="chip">${g.itens.length} lotes</span></div><div class="td-meta">${esc(t.quadra)} · lotes ${esc(ls.length > 8 ? ls.slice(0, 8).join(', ') + '…' : ls.join(', '))} · ${esc(t.destino)}</div></div>
    <div class="td-val"><b class="${cls(tot)}">${brl2(tot)}</b>${jur ? `<small class="neg">juros ${brl2(jur)}</small>` : ''}<small class="muted">${g.itens.length} × ${brl2(t.valor)}</small>
      ${!pago(t) ? (venc ? '<span class="chip bad">vencida</span>' : hj ? '<span class="chip warn">hoje</span>' : '') : (o.mostrarPago ? '<span class="chip ok">pago</span>' : '')}</div>
    ${o.baixa && !pago(t) ? `<button class="btn sm" data-act="baixarG" data-ids="${ids.join(',')}">Baixar</button>` : ''}
  </div>`;
}
const linhaEntrada = (e, o) => e.t ? linhaTitulo(e.t, o) : linhaGrupo(e, o);
ACOES.gsel = d => { const ids = d.ids.split(','); const all = ids.every(id => listaSel.has(id)); ids.forEach(id => all ? listaSel.delete(id) : listaSel.add(id)); refreshView(); };
ACOES.baixarG = d => formBaixa(d.ids.split(','));
ACOES.grp = d => {
  const ts = d.ids.split(',').map(id => S.byId.get(id)).filter(Boolean); if (!ts.length) return;
  const ab = ts.filter(t => !pago(t)); const tot = sum(ts, t => t.valor);
  const f = folha(baseItem(ts[0]), `<div class="resumo"><b>${ts.length} lançamentos · quadra ${esc(ts[0].quadra)}</b><div class="big ${cls(tot)}">${brl2(tot)}</div><span class="muted">${esc(ts[0].destino)} · ${pago(ts[0]) ? 'pago em ' + dBR(ts[0].pag) : 'vence ' + dBR(ts[0].venc)}</span></div>
    <div class="acts wrap">${ab.length ? `<button class="btn primary" id="gBx">${ts[0].valor < 0 ? 'Registrar pagamento' : 'Registrar recebimento'} de ${ab.length === ts.length ? 'todos' : ab.length + ' em aberto'}</button>` : ''}<button class="btn danger ghost" id="gDel">Excluir todos</button></div>
    <div class="tlist">${ts.sort((a, b) => lotSort(a.lote, b.lote)).map(t => linhaTitulo(t, { mostrarPago: true })).join('')}</div>`, { largo: true });
  const bx = f.el.querySelector('#gBx'); if (bx) bx.onclick = () => { f.fechar(); formBaixa(ab.map(t => t.id)); };
  f.el.querySelector('#gDel').onclick = async () => { if (await confirmar(`Excluir os <b>${ts.length}</b> lançamentos deste grupo?`, 'Excluir')) { excluirTitulos(ts.map(t => t.id)); f.fechar(); toast(ts.length + ' lançamentos excluídos'); } };
  f.el.addEventListener('click', ev => { if (ev.target.closest('[data-act="abrir"]')) f.fechar(); });
};
function kpi(lab, val, sub, c, act) { return `<div class="card kpi ${act ? 'click' : ''}" ${act ? `data-act="ir" data-href="${act}"` : ''}><div class="lab">${lab}</div><div class="val ${c || ''}">${val}</div><div class="sub">${sub || ''}</div></div>`; }

/* =========================================================
   Gráficos
   ========================================================= */
const charts = {};
function grafico(id, cfg) { if (charts[id]) { charts[id].destroy(); delete charts[id]; } const c = document.getElementById(id); if (c) charts[id] = new Chart(c, cfg); }
function graficoSaldo(id, dias, limite) {
  let ds = dias;
  if (limite) { const lim = addDays(hoje(), limite); ds = dias.filter(x => x.d <= lim); ds = ds.concat([{ d: lim, v: ds[ds.length - 1].v }]); }
  const minIn = ds.reduce((m, x) => x.v < m.v ? x : m, ds[0]);
  grafico(id, { type: 'line', data: { labels: ds.map(x => dBR(x.d)), datasets: [
    { label: 'Saldo', data: ds.map(x => x.v), borderColor: css('--bal'), borderWidth: 2, stepped: true, pointRadius: 0, fill: { target: { value: 0 }, above: 'rgba(46,139,110,.12)', below: 'rgba(194,73,61,.14)' } },
    { label: 'Menor saldo', data: ds.map(x => x === minIn ? x.v : null), pointRadius: 5, pointBackgroundColor: css('--out'), showLine: false }
  ] }, options: { responsive: true, maintainAspectRatio: false, animation: false, interaction: { mode: 'nearest', intersect: false },
    plugins: { legend: { display: false }, tooltip: { callbacks: { label: c => brl(c.parsed.y) } } },
    scales: { x: { ticks: { color: css('--muted'), maxTicksLimit: 6, maxRotation: 0 }, grid: { display: false } }, y: { ticks: { color: css('--muted'), callback: v => short(v) }, grid: { color: css('--line') } } } } });
}

/* =========================================================
   Telas
   ========================================================= */
const VIEWS = {};

VIEWS.inicio = {
  titulo: 'Início',
  refresh(el) {
    const h = hoje(), L = S.list, ab = L.filter(t => !pago(t));
    if (!L.length && !S.estrutura) { el.innerHTML = `<div class="empty card"><h2>Bem-vindo!</h2><p>O sistema está vazio. Para começar, importe sua planilha <b>SFG Projeto</b> com todo o histórico.</p><a class="btn primary" href="#/dados">Importar planilha</a></div>`; return; }
    const sd = saldos(), pj = projecao();
    const venc = ab.filter(t => t.valor < 0 && t.venc < h), p7 = ab.filter(t => t.valor < 0 && t.venc >= h && t.venc <= addDays(h, 7)), r30 = ab.filter(t => t.valor > 0 && t.venc <= addDays(h, 30));
    const prox = ab.filter(t => t.venc <= addDays(h, 14)).sort((a, b) => a.venc.localeCompare(b.venc) || a.valor - b.valor);
    const neg = pj.dias.find(x => x.v < 0);
    const r30txt = r30.length ? [...new Set(r30.map(t => t.quadra + (t.lote && t.lote !== '-' ? ' · ' + t.lote : '')))].slice(0, 2).join(', ') : 'nada previsto';
    const vencDesde = venc.length ? venc.map(t => t.venc).sort()[0] : '';
    // saldo no fim de cada mês
    const meses = []; let m0 = h.slice(0, 7) + '-01';
    for (let i = 0; i < 8; i++) { const ini = addMonths(m0, i); const fim = addDays(addMonths(ini, 1), -1); let v = pj.base; for (const x of pj.dias) if (x.d <= fim) v = x.v; meses.push({ k: ini.slice(0, 7), v }); }
    const qs = E().quadras.map(resumoQuadra).sort((a, b) => b.proj - a.proj); const qmax = Math.max(1, ...qs.map(q => Math.abs(q.proj)));
    const hora = new Date().getHours(); const saud = hora < 12 ? 'Bom dia' : hora < 18 ? 'Boa tarde' : 'Boa noite';
    const nome = (S.user?.displayName || '').split(' ')[0] || ((S.user?.email || '').startsWith('luis') ? 'Luis' : '');
    el.innerHTML = `<div class="eyebrow">${new Date().toLocaleDateString('pt-BR', { weekday: 'long', day: 'numeric', month: 'long' })}</div>
    <h2 class="saud">${saud}${nome ? ', ' + esc(nome) : ''}</h2>
    <div class="grid kpis k4">
      <div class="card kpi hero click" data-act="ir" data-href="#/extrato"><div class="lab">Saldo em caixa</div><div class="val">${brl(sd.total)}</div><div class="sub">${sd.contas.map(c => esc(c.nome) + ' ' + brl(c.saldo)).join(' · ')}</div></div>
      ${kpi('Vencidas a pagar', brl(sum(venc, t => t.valor)), venc.length ? venc.length + ' conta(s) · desde ' + dBR(vencDesde).slice(0, 5) : 'nenhuma', venc.length ? 'neg' : '', '#/pagar')}
      ${kpi('A pagar em 7 dias', brl(sum(p7, t => t.valor)), p7.length + ' lançamento(s)', '', '#/pagar')}
      ${kpi('A receber em 30 dias', brl(sum(r30, t => t.valor)), esc(r30txt), 'pos', '#/receber')}
    </div>
    ${neg ? `<div class="alerta-caixa">${ic('fluxo', 20)}<span>O caixa fica negativo a partir de <b>${dBR(neg.d).slice(0, 5)}</b> e chega a <b>${brl(pj.min.v)}</b> em <b>${dBR(pj.min.d).slice(0, 5)}</b>. Planeje um aporte ou antecipe recebíveis.</span><a class="btn link" href="#/fluxo">Ver fluxo</a></div>` : ''}
    <div class="grid lg">
      <div class="card"><h2>Saldo projetado <span class="hint">fim de cada mês · contas a pagar e a receber</span></h2><div class="chartbox"><canvas id="cIniM"></canvas></div></div>
      <div class="card"><h2>Resultado projetado por quadra</h2>${qs.map(q => `<a class="qbar" href="#/obras/${encodeURIComponent(q.q)}"><b>${esc(q.q)}</b><span class="bar"><i style="width:${Math.max(0, q.proj) / qmax * 100}%"></i></span><span style="text-align:right" class="${cls(q.proj)}">${short(q.proj)}</span></a>`).join('')}
        <p class="muted small" style="border-top:1px solid var(--line);padding-top:10px;margin:10px 0 0">Total projetado <b style="color:var(--ink)">${short(sum(qs, q => q.proj))}</b> · previsto ${short(sum(qs, q => q.prevRes))}</p></div>
    </div>
    <div class="card"><h2>Próximos vencimentos <span class="hint">14 dias, inclui vencidas</span><button class="btn sm" data-act="ics">${ic('calendario', 16)} Lembretes no calendário</button></h2><div class="tlist">${(() => { const g = agrupar(prox, 'venc'); return g.slice(0, 15).map(e => linhaEntrada(e, { baixa: true })).join('') + (g.length > 15 ? `<a class="btn link" href="#/pagar">Ver todos (${prox.length} lançamentos)</a>` : ''); })() || '<p class="muted">Nada vencendo.</p>'}</div></div>`;
    grafico('cIniM', { type: 'bar', data: { labels: meses.map(x => MES3[+x.k.slice(5, 7) - 1] + '/' + x.k.slice(2, 4)), datasets: [{ data: meses.map(x => x.v), backgroundColor: meses.map(x => x.v < 0 ? css('--out') : css('--bal')), borderRadius: 6, maxBarThickness: 46 }] },
      options: { responsive: true, maintainAspectRatio: false, animation: false, plugins: { legend: { display: false }, tooltip: { callbacks: { label: c => brl(c.parsed.y) } } },
        scales: { x: { ticks: { color: css('--muted') }, grid: { display: false } }, y: { ticks: { color: css('--muted'), callback: v => short(v), maxTicksLimit: 6 }, grid: { color: css('--line') } } } } });
  }
};

function viewLista(modo) {
  const F = { q: '', per: modo === 'extrato' ? 'mes' : 'todos', st: 'aberto', quadra: '', dest: '', conta: '', ini: '', fim: '' };
  let selMode = false, limite = 200, agrup = true;
  const per = modo === 'extrato'
    ? [['mes', 'Este mês'], ['mesant', 'Mês passado'], ['30d', 'Últimos 30 dias'], ['ano', 'Este ano'], ['todos', 'Tudo'], ['custom', 'Período…']]
    : [['todos', 'Todas as datas'], ['vencidas', 'Vencidas'], ['hoje', 'Vencem hoje'], ['7', 'Próximos 7 dias'], ['30', 'Próximos 30 dias'], ['mes', 'Este mês'], ['proxmes', 'Próximo mês'], ['custom', 'Período…']];
  return {
    titulo: modo === 'pagar' ? 'Contas a pagar' : modo === 'receber' ? 'Contas a receber' : 'Extrato',
    montar(el) {
      listaSel.clear(); selMode = false; limite = 200;
      el.innerHTML = `<div class="filtros card">
        <input type="search" id="fq" placeholder="Buscar descrição, fornecedor…" value="${esc(F.q)}">
        ${modo !== 'extrato' ? `<select id="fst">${opts([['aberto', 'Em aberto'], ['pago', modo === 'pagar' ? 'Pagas' : 'Recebidas'], ['todos', 'Todas']], F.st)}</select>` : `<select id="fconta">${opts(E().contas.map(c => c.nome), F.conta, 'Todas as contas')}</select>`}
        <select id="fper">${opts(per, F.per)}</select>
        <span id="fcustom" class="${F.per === 'custom' ? '' : 'hidden'}"><input type="date" id="fini" value="${F.ini}"> até <input type="date" id="ffim" value="${F.fim}"></span>
        <select id="fquadra">${opts(quadrasOpts().map(([v, l]) => [v, l.replace('Quadra ', '')]), F.quadra, 'Todas as quadras')}</select>
        <select id="fdest">${opts(todosDestinos(), F.dest, 'Todos os destinos')}</select>
        <label class="tgl"><input type="checkbox" id="fgrp" ${agrup ? 'checked' : ''}> Agrupar lotes</label>
        ${modo !== 'extrato' ? `<button class="btn" id="bSel">Selecionar</button>` : ''}
      </div><div id="lsaldo"></div><div class="card"><div id="ltot" class="ltot"></div><div id="lrows" class="tlist"></div><div id="lmore"></div></div><div id="selbar" class="selbar hidden"></div>`;
      const bind = (id, k, ev = 'change') => $('#' + id, el)?.addEventListener(ev, e => { F[k] = e.target.value; limite = 200; if (k === 'per') $('#fcustom', el).classList.toggle('hidden', F.per !== 'custom'); this.refresh(el); });
      bind('fq', 'q', 'input'); bind('fst', 'st'); bind('fconta', 'conta'); bind('fper', 'per'); bind('fini', 'ini'); bind('ffim', 'fim'); bind('fquadra', 'quadra'); bind('fdest', 'dest');
      $('#fgrp', el).onchange = e => { agrup = e.target.checked; limite = 200; this.refresh(el); };
      const bs = $('#bSel', el); if (bs) bs.onclick = () => { selMode = !selMode; listaSel.clear(); bs.classList.toggle('on', selMode); bs.textContent = selMode ? 'Cancelar seleção' : 'Selecionar'; this.refresh(el); };
    },
    sair() { listaSel.clear(); },
    refresh(el) {
      const h = hoje(); const st = modo === 'extrato' ? 'pago' : F.st;
      const campo = st === 'aberto' ? 'venc' : st === 'pago' ? 'pag' : 'venc';
      const q = F.q.trim().toLowerCase();
      const mes = h.slice(0, 7), mesAnt = addMonths(h.slice(0, 7) + '-01', -1).slice(0, 7), prox = addMonths(h.slice(0, 7) + '-01', 1).slice(0, 7);
      const okPer = d => {
        d = d || ''; switch (F.per) {
          case 'vencidas': return d < h; case 'hoje': return d === h; case '7': return d >= h && d <= addDays(h, 7); case '30': return d >= h && d <= addDays(h, 30);
          case 'mes': return d.startsWith(mes); case 'mesant': return d.startsWith(mesAnt); case 'proxmes': return d.startsWith(prox); case '30d': return d >= addDays(h, -30) && d <= h;
          case 'ano': return d.startsWith(h.slice(0, 4)); case 'custom': return (!F.ini || d >= F.ini) && (!F.fim || d <= F.fim); default: return true;
        }
      };
      let rows = S.list.filter(t => {
        if (modo === 'pagar' && !(t.valor < 0)) return false;
        if (modo === 'receber' && !(t.valor > 0)) return false;
        if (st === 'aberto' && pago(t)) return false; if (st === 'pago' && !pago(t)) return false;
        if (F.per === 'vencidas' && pago(t)) return false;
        if (F.quadra && t.quadra !== F.quadra) return false; if (F.dest && t.destino !== F.dest) return false;
        if (F.conta && t.conta !== F.conta) return false;
        if (q && !(t.item.toLowerCase().includes(q) || (t.forn || '').toLowerCase().includes(q) || (t.obs || '').toLowerCase().includes(q))) return false;
        return okPer(campo === 'pag' ? (t.pag || t.venc) : t.venc);
      });
      const desc = campo === 'pag';
      rows.sort((a, b) => { const x = (campo === 'pag' ? (a.pag || a.venc) : a.venc) || '', y = (campo === 'pag' ? (b.pag || b.venc) : b.venc) || ''; return desc ? y.localeCompare(x) || (b.criado || 0) - (a.criado || 0) : x.localeCompare(y) || (a.criado || 0) - (b.criado || 0); });
      // saldo corrido no extrato
      let saldoMap = null;
      if (modo === 'extrato' && !q && !F.quadra && !F.dest) {
        const sd = saldos(); let s = F.conta ? (E().contas.find(c => c.nome === F.conta)?.inicial || 0) : sum(E().contas, c => +c.inicial || 0);
        const all = S.list.filter(t => pago(t) && (!F.conta || t.conta === F.conta)).sort((a, b) => (a.pag || a.venc || '').localeCompare(b.pag || b.venc || '') || (a.criado || 0) - (b.criado || 0));
        saldoMap = new Map(); for (const t of all) { s += t.valor + (t.juros || 0); saldoMap.set(t.id, s); }
        $('#lsaldo', el).innerHTML = `<div class="grid kpis">${kpi('Saldo total', brl(sd.total), 'todas as contas', cls(sd.total))}${sd.contas.map(c => kpi(esc(c.nome), brl(c.saldo), c.inicial ? 'saldo inicial ' + brl(c.inicial) : '', cls(c.saldo))).join('')}</div>`;
      } else if (modo === 'extrato') $('#lsaldo', el).innerHTML = '';
      const tot = sum(rows, t => t.valor + (pago(t) ? (t.juros || 0) : 0));
      const ent = sum(rows, t => Math.max(0, t.valor)), sai = sum(rows, t => Math.min(0, t.valor) + Math.min(0, pago(t) ? (t.juros || 0) : 0));
      $('#ltot', el).innerHTML = `<span><b>${rows.length.toLocaleString('pt-BR')}</b> lançamento(s)</span>${modo === 'extrato' ? `<span>Entradas <b class="pos">${brl(ent)}</b></span><span>Saídas <b class="neg">${brl(sai)}</b></span>` : ''}<span>Total <b class="${cls(tot)}">${brl2(tot)}</b></span>${selMode ? `<button class="btn link sm" id="bSelAll">Selecionar todos em aberto</button>` : ''}`;
      const sa = $('#bSelAll', el); if (sa) sa.onclick = () => { rows.filter(t => !pago(t)).forEach(t => listaSel.add(t.id)); this.refresh(el); };
      let html = '', mesAtual = null;
      const dOf = x => ((campo === 'pag' ? (x.pag || x.venc) : x.venc) || '');
      const subMes = {}; for (const x of rows) { const m = dOf(x).slice(0, 7); subMes[m] = (subMes[m] || 0) + x.valor + (pago(x) ? (x.juros || 0) : 0); }
      const entradas = agrup ? agrupar(rows, campo) : rows.map(t => ({ t }));
      for (const e of entradas.slice(0, limite)) {
        const t = e.t || e.itens[0]; const m = dOf(t).slice(0, 7);
        if (m !== mesAtual) { mesAtual = m; html += `<div class="mhead"><span>${mesNome(m)}</span><span class="${cls(subMes[m])}">${brl(subMes[m])}</span></div>`; }
        html += linhaEntrada(e, { sel: selMode, data: campo, saldo: saldoMap && e.t ? saldoMap.get(t.id) : undefined, mostrarPago: st === 'todos' });
      }
      $('#lrows', el).innerHTML = html || '<p class="muted pad">Nenhum lançamento encontrado.</p>';
      $('#lmore', el).innerHTML = entradas.length > limite ? `<button class="btn full" id="bMais">Mostrar mais (${(entradas.length - limite).toLocaleString('pt-BR')} linhas restantes)</button>` : '';
      const bm = $('#bMais', el); if (bm) bm.onclick = () => { limite += 300; this.refresh(el); };
      const sb = $('#selbar', el);
      if (selMode && listaSel.size) {
        const sel = [...listaSel].map(id => S.byId.get(id)).filter(Boolean);
        sb.classList.remove('hidden');
        sb.innerHTML = `<span><b>${sel.length}</b> selecionado(s) · <b class="${cls(sum(sel, t => t.valor))}">${brl2(sum(sel, t => t.valor))}</b></span><span class="acts"><button class="btn danger ghost" id="sbDel">Excluir</button><button class="btn primary" id="sbBx">Dar baixa</button></span>`;
        $('#sbBx', sb).onclick = () => formBaixa(sel.filter(t => !pago(t)).map(t => t.id));
        $('#sbDel', sb).onclick = async () => { if (await confirmar(`Excluir <b>${sel.length}</b> lançamentos?`, 'Excluir')) { excluirTitulos(sel.map(t => t.id)); listaSel.clear(); toast(sel.length + ' excluídos'); } };
      } else sb.classList.add('hidden');
    }
  };
}
VIEWS.pagar = viewLista('pagar');
VIEWS.receber = viewLista('receber');
VIEWS.extrato = viewLista('extrato');

/* ---------- Obras ---------- */
function pctBar(p, custo) {
  if (p === null || p === undefined || !isFinite(p)) return '<span class="muted">—</span>';
  const over = custo && p > 1.0001;
  return `<span class="pbar"><span class="bar"><i style="width:${Math.min(100, p * 100)}%;${over ? 'background:var(--out)' : ''}"></i></span><span>${fPct.format(p)}</span></span>`;
}
VIEWS.obras = {
  titulo: a => a.length === 2 ? `Lote ${a[0]} · ${a[1]}` : a.length === 1 ? 'Quadra ' + a[0] : 'Obras · ' + E().local,
  refresh(el, a) { if (a.length === 2) return obraLote(el, a[0], a[1]); if (a.length === 1) return obraQuadra(el, a[0]); return obrasResumo(el); }
};
function obrasResumo(el) {
  const rows = E().quadras.map(resumoQuadra); const T = k => sum(rows, r => r[k]);
  el.innerHTML = `<div class="grid kpis">${kpi('Lotes', T('n'), E().quadras.length + ' quadras', '')}${kpi('Custo realizado', brl(T('custo')), '', 'neg')}${kpi('A pagar', brl(T('pagar')), '', 'neg')}${kpi('Receita realizada', brl(T('receita')), '', 'pos')}${kpi('A receber', brl(T('receber')), '', 'pos')}${kpi('Resultado projetado', brl(T('proj')), 'previsto ' + brl(T('prevRes')), cls(T('proj')))}</div>
  <div class="card"><h2>Quadras <span class="hint">toque para abrir</span></h2><div class="tbl-wrap"><table class="tbl"><thead><tr><th class="l">Quadra</th><th>Lotes</th><th>Custo realizado</th><th>A pagar</th><th>Receita</th><th>A receber</th><th>Resultado projetado</th><th>Previsto</th><th>Custo vs orçado</th><th>Checklist</th></tr></thead><tbody>
  ${rows.map(r => `<tr class="click" data-act="ir" data-href="#/obras/${encodeURIComponent(r.q)}"><td class="l"><b>${esc(r.q)}</b></td><td>${r.n}</td><td class="neg">${brl(r.custo)}</td><td class="neg">${brl(r.pagar)}</td><td class="pos">${brl(r.receita)}</td><td class="pos">${brl(r.receber)}</td><td class="${cls(r.proj)}"><b>${brl(r.proj)}</b></td><td>${r.prevCusto ? brl(r.prevRes) : '—'}</td><td>${pctBar(r.prevCusto ? (r.custo + r.pagar) / r.prevCusto : null, true)}</td><td>${pctBar(r.cl)}</td></tr>`).join('') || '<tr><td colspan="10" class="l muted">Nenhuma quadra cadastrada. Vá em Cadastros.</td></tr>'}
  </tbody></table></div></div>`;
}
function obraQuadra(el, q) {
  const s = resumoQuadra(q);
  const real = byDest(S.list.filter(t => t.quadra === q && pago(t))), pag = byDest(S.list.filter(t => t.quadra === q && !pago(t) && t.valor < 0)), p = previstoLote(q);
  const sec = secoesViab(); const labs = [...sec.op, ...sec.nonop].filter(k => Math.abs(p[k] || 0) > .5 || Math.abs(real[k] || 0) > .5 || Math.abs(pag[k] || 0) > .5);
  el.innerHTML = `<div class="grid kpis">${kpi('Lotes', s.n, '', '')}${kpi('Custo realizado', brl(s.custo), s.n ? short(s.custo / s.n) + ' por lote' : '', 'neg')}${kpi('A pagar', brl(s.pagar), '', 'neg')}${kpi('Receita realizada', brl(s.receita), '', 'pos')}${kpi('A receber', brl(s.receber), '', 'pos')}${kpi('Resultado projetado', brl(s.proj), 'previsto ' + brl(s.prevRes), cls(s.proj))}${kpi('Checklist', s.cl === null ? '—' : fPct.format(s.cl), 'documentos', '', '#/checklist/' + encodeURIComponent(q))}</div>
  <div class="grid two">
    <div class="card"><h2>Previsto x realizado <span class="hint">quadra inteira</span> <button class="btn sm" id="bOrc">Editar orçamento</button></h2>
      <div class="legend"><span style="--c:var(--pend)">Previsto</span><span style="--c:var(--out)">Realizado</span><span style="--c:var(--out-soft)">A pagar</span></div><div class="chartbox tall"><canvas id="cQ"></canvas></div></div>
    <div class="card"><h2>Lotes <span class="hint">toque para ver a viabilidade</span></h2><div class="tbl-wrap"><table class="tbl"><thead><tr><th class="l">Lote</th><th>Custo real.</th><th>A pagar</th><th>Venda</th><th class="l">Status</th><th>Checklist</th></tr></thead><tbody>
    ${lotesDe(q).map(l => {
      const ts = S.list.filter(t => t.quadra === q && t.lote === l);
      const custo = sum(ts.filter(t => pago(t) && t.destino !== 'Aporte'), t => Math.min(t.valor, 0) + Math.min(t.juros || 0, 0));
      const venda = sum(ts.filter(t => pago(t) && t.destino === 'Venda'), t => t.valor), aRec = sum(ts.filter(t => !pago(t) && t.valor > 0), t => t.valor);
      const prox = ts.filter(t => !pago(t) && t.valor > 0).map(t => t.venc).sort()[0];
      const st = venda > 0 && aRec < .5 ? '<span class="chip ok">Recebido</span>' : venda > 0 ? '<span class="chip ok">Parcial</span>' : aRec > 0 ? `<span class="chip warn">A receber ${prox ? dBR(prox) : ''}</span>` : '<span class="muted">—</span>';
      return `<tr class="click" data-act="ir" data-href="#/obras/${encodeURIComponent(q)}/${encodeURIComponent(l)}"><td class="l"><b>${esc(l)}</b></td><td class="neg">${brl(custo)}</td><td class="neg">${brl(sum(ts.filter(t => !pago(t) && t.valor < 0), t => t.valor))}</td><td class="pos">${brl(venda + aRec)}</td><td class="l">${st}</td><td>${pctBar(ckPct(q, l))}</td></tr>`;
    }).join('') || '<tr><td colspan="6" class="l muted">Sem lotes cadastrados</td></tr>'}</tbody></table></div></div>
  </div>`;
  $('#bOrc', el).onclick = () => formOrcamento(q);
  const soft = css('--out') + '66';
  grafico('cQ', { type: 'bar', data: { labels: labs, datasets: [
    { label: 'Previsto', data: labs.map(k => -(p[k] || 0) * s.n), backgroundColor: css('--pend'), borderRadius: 3, stack: 'p' },
    { label: 'Realizado', data: labs.map(k => -Math.min(0, real[k] || 0)), backgroundColor: css('--out'), borderRadius: 3, stack: 'r' },
    { label: 'A pagar', data: labs.map(k => -Math.min(0, pag[k] || 0)), backgroundColor: soft, borderRadius: 3, stack: 'r' }
  ] }, options: { indexAxis: 'y', responsive: true, maintainAspectRatio: false, animation: false, interaction: { mode: 'index', intersect: false }, plugins: { legend: { display: false }, tooltip: { callbacks: { label: c => c.dataset.label + ': ' + brl(c.parsed.x) } } },
    scales: { x: { stacked: true, ticks: { color: css('--muted'), maxTicksLimit: 5, callback: v => short(v) }, grid: { color: css('--line') } }, y: { stacked: true, ticks: { color: css('--ink') }, grid: { display: false } } } } });
}
function obraLote(el, q, l) {
  const ts = S.list.filter(t => t.quadra === q && t.lote === l);
  const P = previstoLote(q), R = byDest(ts.filter(pago)), X = byDest(ts.filter(t => !pago(t))), C = {};
  for (const k of new Set([...Object.keys(R), ...Object.keys(X)])) C[k] = (R[k] || 0) + (X[k] || 0);
  const sec = secoesViab();
  let html = '<thead><tr><th class="l">Destino</th><th>Previsto</th><th>Realizado</th><th>Δ</th><th>Em aberto</th><th>Projetado</th></tr></thead><tbody>';
  const linha = (k, p, r, x, c, rec, tot) => { const d = Math.abs(p) > .5 ? r / p - 1 : null; const bad = d === null ? false : rec ? d < -1e-4 : d > 1e-4; return `<tr class="${tot ? 'tot' : ''}"><td class="l">${esc(k)}</td><td>${brl(p)}</td><td class="${cls(r)}">${brl(r)}</td><td class="${d === null ? 'muted' : bad ? 'neg' : 'pos'}">${d === null ? '—' : (d > 0 ? '+' : '') + fPct.format(d)}</td><td class="${cls(x)}">${Math.abs(x) > .5 ? brl(x) : '—'}</td><td class="${cls(c)}">${brl(c)}</td></tr>`; };
  for (const s of ['op', 'nonop', 'rec']) {
    html += `<tr class="sec"><td colspan="6">${SECAO_NOME[s]}</td></tr>`; const T = { p: 0, r: 0, x: 0, c: 0 };
    for (const k of sec[s]) { const p = P[k] || 0, r = R[k] || 0, x = X[k] || 0, c = C[k] || 0; T.p += p; T.r += r; T.x += x; T.c += c; html += linha(k, p, r, x, c, s === 'rec'); }
    html += linha('Total', T.p, T.r, T.x, T.c, s === 'rec', true);
  }
  const iP = indicadores(P), iR = indicadores(R), iC = indicadores(C);
  const ind = [['ROI', 'roi', pct], ['Lucro operacional', 'lop', brl], ['Margem operacional', 'mop', pct], ['Lucro líquido', 'll', brl], ['Margem líquida', 'ml', pct]];
  const custo = sum(ts.filter(t => pago(t) && t.destino !== 'Aporte'), t => Math.min(t.valor, 0) + Math.min(t.juros || 0, 0));
  el.innerHTML = `<div class="grid kpis">${kpi('Custo realizado', brl(custo), 'previsto ' + brl(iP.op + iP.non), 'neg')}${kpi('Lucro líquido projetado', brl(iC.ll), 'previsto ' + brl(iP.ll), cls(iC.ll))}${kpi('ROI projetado', pct(iC.roi), 'previsto ' + pct(iP.roi), '')}${kpi('Checklist', ckPct(q, l) === null ? '—' : fPct.format(ckPct(q, l)), 'documentos', '', '#/checklist/' + encodeURIComponent(q))}</div>
  <div class="grid lg"><div class="card"><h2>Viabilidade <span class="hint">projetado = realizado + em aberto</span></h2><div class="tbl-wrap"><table class="tbl">${html}</tbody></table></div></div>
  <div class="card"><h2>Indicadores</h2><div class="tbl-wrap"><table class="tbl"><thead><tr><th class="l">Indicador</th><th>Previsto</th><th>Realizado</th><th>Projetado</th></tr></thead><tbody>${ind.map(([n, k, fm]) => `<tr><td class="l">${n}</td><td>${fm(iP[k])}</td><td class="${typeof iR[k] === 'number' ? cls(iR[k]) : ''}">${fm(iR[k])}</td><td class="${typeof iC[k] === 'number' ? cls(iC[k]) : ''}"><b>${fm(iC[k])}</b></td></tr>`).join('')}</tbody></table></div>
  <p class="muted small">Realizado considera só o que já foi pago/recebido. Enquanto a venda não entra, o resultado realizado fica negativo.</p></div></div>
  <div class="card"><h2>Lançamentos do lote <span class="hint">${ts.length}</span></h2><div class="tlist" id="lotLanc"></div><div id="lotMais"></div></div>`;
  const ord = ts.sort((a, b) => (dataRef(b) || '').localeCompare(dataRef(a) || ''));
  let lim = 25; const pinta = () => { $('#lotLanc', el).innerHTML = ord.slice(0, lim).map(t => linhaTitulo(t, { mostrarPago: true })).join('') || '<p class="muted">Nenhum lançamento.</p>'; $('#lotMais', el).innerHTML = ord.length > lim ? `<button class="btn full" id="bLm">Mostrar todos (${ord.length})</button>` : ''; const b = $('#bLm', el); if (b) b.onclick = () => { lim = 1e9; pinta(); }; };
  pinta();
}
function formOrcamento(q) {
  const atualQ = (S.orcamento && S.orcamento.quadras && S.orcamento.quadras[q]) || {};
  const ds = [...(E().destinos.op || []), ...(E().destinos.nonop || []), ...(E().destinos.rec || [])];
  const f = folha('Orçamento por lote · Quadra ' + q, `<form id="fOrc" class="form"><p class="muted">Valores previstos <b>por lote</b>. Custos e receitas em valor positivo.</p>
    ${['op', 'nonop', 'rec'].map(s => `<div class="lbl">${SECAO_NOME[s]}</div><div class="g2">${(E().destinos[s] || []).map(d => `<label>${esc(d)}<input name="${esc(d)}" inputmode="decimal" value="${atualQ[d] ? fNum2.format(atualQ[d]) : ''}" placeholder="0,00"></label>`).join('')}</div>`).join('')}
    <div class="acts"><button type="button" class="btn" data-close>Cancelar</button><button class="btn primary">Salvar orçamento</button></div></form>`);
  const F = f.el.querySelector('#fOrc');
  F.onsubmit = e => { e.preventDefault(); const o = {}; for (const d of ds) { const v = parseBR(F.elements[d]?.value); if (v > 0) o[d] = v; } gravar(doc(db, 'config', 'orcamento'), { quadras: { [q]: o } }, { merge: true }).catch(onErroEscrita); S.orcamento = S.orcamento || { quadras: {} }; S.orcamento.quadras = { ...(S.orcamento.quadras || {}), [q]: o }; f.fechar(); toast('Orçamento salvo'); refreshView(); };
}

/* ---------- Fluxo ---------- */
let fluxoH = 180;
VIEWS.fluxo = {
  titulo: 'Fluxo de caixa',
  refresh(el) {
    const h = hoje(), P = projecao(); const neg = P.dias.find(x => x.v < 0);
    const em = d => P.base + sum(P.ev.filter(e => e.d <= addDays(h, d)), e => e.v);
    const venc = S.list.filter(t => !pago(t) && t.venc < h);
    el.innerHTML = `${venc.length ? `<div class="alert"><b>${venc.length} lançamento(s) vencido(s) em aberto</b> (${brl(sum(venc, t => t.valor))}) estão considerados como se acontecessem hoje. Se já foram pagos, registre a baixa.</div>` : ''}
    <div class="grid kpis">${kpi('Saldo hoje', brl(P.base), dBR(h), cls(P.base))}${kpi('Em 30 dias', brl(em(30)), '', cls(em(30)))}${kpi('Em 90 dias', brl(em(90)), '', cls(em(90)))}${kpi('Menor saldo', brl(P.min.v), 'em ' + dBR(P.min.d), cls(P.min.v))}${kpi('Primeiro saldo negativo', neg ? dBR(neg.d) : 'nenhum', neg ? 'precisa de caixa ou aporte' : '', neg ? 'neg' : 'pos')}</div>
    <div class="card"><h2>Saldo projetado</h2><div class="seg" id="segH">${[[90, '3 meses'], [180, '6 meses'], [365, '12 meses'], [0, 'Tudo']].map(([v, l]) => `<button data-v="${v}" class="${v === fluxoH ? 'on' : ''}">${l}</button>`).join('')}</div><div class="chartbox tall"><canvas id="cFl"></canvas></div></div>
    <div class="card"><h2>Próximos 60 dias</h2><div class="tbl-wrap"><table class="tbl"><thead><tr><th class="l">Data</th><th class="l">Descrição</th><th class="l">Quadra</th><th>Valor</th><th>Saldo após</th></tr></thead><tbody>
    ${P.ev.filter(e => e.d <= addDays(h, 60)).map(e => `<tr class="click" data-act="abrir" data-id="${e.t.id}"><td class="l">${dBR(e.d)}${e.t.venc < h ? ' <span class="chip bad">vencida</span>' : ''}</td><td class="l wrap">${esc(e.t.item)}</td><td class="l">${esc(e.t.quadra)}</td><td class="${cls(e.v)}">${brl2(e.v)}</td><td class="${cls(e.bal)}">${brl(e.bal)}</td></tr>`).join('') || '<tr><td colspan="5" class="l muted">Nada previsto.</td></tr>'}</tbody></table></div></div>`;
    $$('#segH button', el).forEach(b => b.onclick = () => { fluxoH = +b.dataset.v; this.refresh(el); });
    graficoSaldo('cFl', P.dias, fluxoH || 0);
  }
};

/* ---------- Checklist ---------- */
VIEWS.checklist = {
  titulo: a => a[0] ? 'Checklist · ' + a[0] : 'Checklist de documentos',
  refresh(el, a) {
    const c = S.checklist || { docs: DOCS_PADRAO, lotes: {} }; const docs = c.docs || DOCS_PADRAO;
    const qs = E().quadras.filter(q => lotesDe(q).length); const q = a[0] && qs.includes(a[0]) ? a[0] : '';
    const lotes = (q ? [q] : qs).flatMap(x => lotesDe(x).map(l => [x, l]));
    const st = (x, l) => (c.lotes || {})[x + '|' + l] || '';
    el.innerHTML = `<div class="filtros card"><select id="ckq">${opts(qs, q, 'Todas as quadras')}</select><span class="muted">Toque numa bolinha para marcar ou desmarcar.</span></div>
    <div class="grid two"><div class="card"><h2>Por documento</h2><div class="tbl-wrap"><table class="tbl"><tbody>${docs.map((d, i) => { const n = lotes.filter(([x, l]) => st(x, l)[i] === '1').length; return `<tr><td class="l">${esc(d)}</td><td>${n}/${lotes.length}</td><td>${pctBar(lotes.length ? n / lotes.length : null)}</td></tr>`; }).join('')}</tbody></table></div></div>
    <div class="card"><h2>Por quadra</h2><div class="tbl-wrap"><table class="tbl"><thead><tr><th class="l">Quadra</th><th>Lotes</th><th>Completos</th><th>Andamento</th></tr></thead><tbody>${qs.map(x => { const ls = lotesDe(x); const full = ls.filter(l => docs.every((_, i) => st(x, l)[i] === '1')).length; return `<tr class="click" data-act="ir" data-href="#/checklist/${encodeURIComponent(x)}"><td class="l"><b>${esc(x)}</b></td><td>${ls.length}</td><td>${full}</td><td>${pctBar(ckPct(x))}</td></tr>`; }).join('')}</tbody></table></div></div></div>
    <div class="card"><h2>Matriz</h2><div class="tbl-wrap matriz"><table class="tbl cl"><thead><tr><th class="l">Quadra</th><th class="l">Lote</th>${docs.map(d => `<th><span>${esc(d)}</span></th>`).join('')}</tr></thead><tbody>
    ${lotes.map(([x, l]) => `<tr><td class="l">${esc(x)}</td><td class="l">${esc(l)}</td>${docs.map((d, i) => `<td><button class="dot ${st(x, l)[i] === '1' ? 'y' : ''}" data-act="ck" data-k="${esc(x + '|' + l)}" data-i="${i}" title="${esc(d)}" aria-label="${esc(d)}"></button></td>`).join('')}</tr>`).join('')}</tbody></table></div></div>`;
    $('#ckq', el).onchange = e => { location.hash = '#/checklist' + (e.target.value ? '/' + encodeURIComponent(e.target.value) : ''); };
  }
};
ACOES.ck = d => {
  const c = S.checklist = S.checklist || { docs: DOCS_PADRAO, lotes: {} }; c.lotes = c.lotes || {};
  const docs = c.docs || DOCS_PADRAO; let s = (c.lotes[d.k] || '').padEnd(docs.length, '0').split(''); const i = +d.i;
  s[i] = s[i] === '1' ? '0' : '1'; c.lotes[d.k] = s.join('');
  gravar(doc(db, 'config', 'checklist'), { docs, lotes: { [d.k]: c.lotes[d.k] } }, { merge: true }).catch(onErroEscrita);
  const b = document.querySelector(`[data-act="ck"][data-k="${CSS.escape(d.k)}"][data-i="${i}"]`); if (b) b.classList.toggle('y', s[i] === '1');
};

/* ---------- Cadastros ---------- */
function expandirLotes(txt) {
  const out = [];
  for (const p of String(txt).split(/[,;\n]+/).map(s => s.trim()).filter(Boolean)) {
    const m = p.match(/^(\d+)\s*(?:-|a|até)\s*(\d+)$/i);
    if (m) { const a = +m[1], b = +m[2]; for (let i = Math.min(a, b); i <= Math.max(a, b); i++) out.push(String(i)); } else out.push(p);
  }
  return [...new Set(out)].sort(lotSort);
}
const usados = (k, v) => S.list.filter(t => t[k] === v).length;
VIEWS.cadastros = {
  titulo: 'Cadastros',
  refresh(el) {
    const e = E();
    el.innerHTML = `<div class="grid two">
      <div class="card"><h2>Obra</h2><div class="kv"><span>Local</span><b>${esc(e.local)}</b><button class="btn sm" data-act="cadLocal">Alterar</button></div></div>
      <div class="card"><h2>Contas bancárias <button class="btn sm" data-act="cadConta">Adicionar</button></h2><div class="tbl-wrap"><table class="tbl"><thead><tr><th class="l">Conta</th><th>Saldo inicial</th><th>Saldo atual</th><th></th></tr></thead><tbody>
        ${saldos().contas.map(c => `<tr><td class="l"><b>${esc(c.nome)}</b></td><td>${brl2(c.inicial)}</td><td class="${cls(c.saldo)}">${brl2(c.saldo)}</td><td>${e.contas.some(x => x.nome === c.nome) ? `<button class="btn sm" data-act="cadConta" data-nome="${esc(c.nome)}">Editar</button>` : ''}</td></tr>`).join('')}</tbody></table></div></div>
    </div>
    <div class="card"><h2>Quadras e lotes <button class="btn sm" data-act="cadQuadra">Adicionar quadra</button></h2><div class="tbl-wrap"><table class="tbl"><thead><tr><th class="l">Quadra</th><th>Lotes</th><th class="l">Numeração</th><th>Lançamentos</th><th></th></tr></thead><tbody>
      ${e.quadras.map(q => { const ls = lotesDe(q); return `<tr><td class="l"><b>${esc(q)}</b></td><td>${ls.length}</td><td class="l wrap muted small">${esc(ls.slice(0, 30).join(', '))}${ls.length > 30 ? '…' : ''}</td><td>${usados('quadra', q)}</td><td><button class="btn sm" data-act="cadQuadra" data-q="${esc(q)}">Editar</button></td></tr>`; }).join('')}</tbody></table></div></div>
    <div class="card"><h2>Destinos (categorias)</h2><div class="grid two">${['op', 'nonop', 'rec', 'aporte'].map(s => `<div><div class="lbl">${SECAO_NOME[s]} <button class="btn link sm" data-act="cadDest" data-s="${s}">+ adicionar</button></div><div class="chips">${(e.destinos[s] || []).map(d => `<span class="chip">${esc(d)} <button class="x" data-act="rmDest" data-s="${s}" data-d="${esc(d)}" aria-label="Remover">×</button></span>`).join('')}</div></div>`).join('')}</div></div>
    <div class="card"><h2>Documentos do checklist <button class="btn sm" data-act="cadDoc">Adicionar</button></h2><div class="chips">${((S.checklist && S.checklist.docs) || DOCS_PADRAO).map(d => `<span class="chip">${esc(d)}</span>`).join('')}</div></div>`;
  }
};
const clone = o => JSON.parse(JSON.stringify(o));
ACOES.cadLocal = () => { const f = folha('Local da obra', `<form class="form" id="fx"><label>Nome<input name="v" value="${esc(E().local)}" required></label><div class="acts"><button type="button" class="btn" data-close>Cancelar</button><button class="btn primary">Salvar</button></div></form>`); f.el.querySelector('#fx').onsubmit = e => { e.preventDefault(); const n = clone(E()); n.local = e.target.v.value.trim(); salvarEstrutura(n); f.fechar(); }; };
ACOES.cadConta = d => {
  const c = E().contas.find(x => x.nome === d.nome);
  const f = folha(c ? 'Editar conta' : 'Nova conta bancária', `<form class="form" id="fx"><label>Nome<input name="nome" value="${esc(c?.nome || '')}" required ${c ? 'readonly' : ''}></label><label>Saldo inicial (R$)<input name="ini" inputmode="decimal" value="${c ? fNum2.format(c.inicial || 0) : '0,00'}"></label>
    <div class="acts">${c && !usados('conta', c.nome) ? '<button type="button" class="btn danger ghost" id="rm">Remover</button>' : ''}<button type="button" class="btn" data-close>Cancelar</button><button class="btn primary">Salvar</button></div></form>`);
  const F = f.el.querySelector('#fx');
  F.onsubmit = e => { e.preventDefault(); const n = clone(E()); const nome = F.nome.value.trim(); const ini = parseBR(F.ini.value) || 0; const i = n.contas.findIndex(x => x.nome === nome); if (i >= 0) n.contas[i].inicial = ini; else n.contas.push({ nome, inicial: ini }); salvarEstrutura(n); f.fechar(); };
  const rm = f.el.querySelector('#rm'); if (rm) rm.onclick = () => { const n = clone(E()); n.contas = n.contas.filter(x => x.nome !== c.nome); salvarEstrutura(n); f.fechar(); };
};
ACOES.cadQuadra = d => {
  const q = d.q; const ls = q ? lotesDe(q) : [];
  const f = folha(q ? 'Quadra ' + q : 'Nova quadra', `<form class="form" id="fx"><label>Nome da quadra<input name="q" value="${esc(q || '')}" required ${q ? 'readonly' : ''} placeholder="Ex.: K4"></label>
    <label>Lotes <span class="muted">— use vírgulas e intervalos: 1-45, 19.1, 19.2</span><textarea name="l" rows="4">${esc(ls.join(', '))}</textarea></label>
    <div class="acts">${q && !usados('quadra', q) ? '<button type="button" class="btn danger ghost" id="rm">Remover quadra</button>' : ''}<button type="button" class="btn" data-close>Cancelar</button><button class="btn primary">Salvar</button></div></form>`);
  const F = f.el.querySelector('#fx');
  F.onsubmit = e => { e.preventDefault(); const n = clone(E()); const nome = F.q.value.trim(); if (!q && n.quadras.includes(nome)) { toast('Essa quadra já existe', 'err'); return; } if (!n.quadras.includes(nome)) n.quadras.push(nome); n.lotes[nome] = expandirLotes(F.l.value); salvarEstrutura(n); f.fechar(); toast('Quadra salva'); };
  const rm = f.el.querySelector('#rm'); if (rm) rm.onclick = () => { const n = clone(E()); n.quadras = n.quadras.filter(x => x !== q); delete n.lotes[q]; salvarEstrutura(n); f.fechar(); };
};
ACOES.cadDest = d => { const f = folha('Novo destino · ' + SECAO_NOME[d.s], `<form class="form" id="fx"><label>Nome<input name="v" required></label><div class="acts"><button type="button" class="btn" data-close>Cancelar</button><button class="btn primary">Adicionar</button></div></form>`); f.el.querySelector('#fx').onsubmit = e => { e.preventDefault(); const v = e.target.v.value.trim(); if (todosDestinos().includes(v)) { toast('Esse destino já existe', 'err'); return; } const n = clone(E()); n.destinos[d.s] = [...(n.destinos[d.s] || []), v]; salvarEstrutura(n); f.fechar(); }; };
ACOES.rmDest = async d => { const n = usados('destino', d.d); if (n) { toast(`"${d.d}" tem ${n} lançamentos e não pode ser removido`, 'err'); return; } if (await confirmar(`Remover o destino <b>${esc(d.d)}</b>?`, 'Remover')) { const e = clone(E()); e.destinos[d.s] = e.destinos[d.s].filter(x => x !== d.d); salvarEstrutura(e); } };
ACOES.cadDoc = () => { const f = folha('Novo documento do checklist', `<form class="form" id="fx"><label>Nome<input name="v" required></label><div class="acts"><button type="button" class="btn" data-close>Cancelar</button><button class="btn primary">Adicionar</button></div></form>`); f.el.querySelector('#fx').onsubmit = e => { e.preventDefault(); const docs = [...((S.checklist && S.checklist.docs) || DOCS_PADRAO), e.target.v.value.trim()]; S.checklist = { ...(S.checklist || { lotes: {} }), docs }; gravar(doc(db, 'config', 'checklist'), { docs }, { merge: true }).catch(onErroEscrita); f.fechar(); refreshView(); }; };

/* =========================================================
   Vendas e repasses Caixa
   ========================================================= */
const ETAPAS_V = [
  ['disp', 'Disponível', '#9AA8BE'], ['aprov', 'Cliente aprovado', '#7E93B8'], ['analise', 'Análise Caixa', '#5D78A6'],
  ['contrato', 'Contrato assinado', '#3F5F92'], ['registro', 'Registro / habite-se', '#23477F'], ['recebido', 'Repasse recebido', '#1F7A57']
];
const etapaNome = k => (ETAPAS_V.find(e => e[0] === k) || ['', k])[1];
function infoVendaLote(q, l) {
  const ts = S.list.filter(t => t.quadra === q && t.lote === l && t.destino === 'Venda');
  const rec = sum(ts.filter(pago), t => t.valor), aberto = sum(ts.filter(t => !pago(t)), t => t.valor);
  const prox = ts.filter(t => !pago(t)).map(t => t.venc).sort()[0] || '';
  const ult = ts.filter(pago).map(t => t.pag || t.venc).sort().pop() || '';
  const v = ((S.vendas && S.vendas.lotes) || {})[q + '|' + l] || null;
  let etapa = v && v.etapa, sugerida = false;
  if (!etapa) { sugerida = true; etapa = rec > 0 && aberto < 0.5 ? 'recebido' : aberto > 0 ? 'contrato' : 'disp'; }
  return { q, l, k: q + '|' + l, v, etapa, sugerida, rec, aberto, total: rec + aberto, prox, ult };
}
let vendQ = '';
VIEWS.vendas = {
  titulo: 'Vendas e repasses',
  refresh(el) {
    const h = hoje(); const qs = E().quadras.filter(q => lotesDe(q).length);
    if (vendQ && !qs.includes(vendQ)) vendQ = '';
    const lots = (vendQ ? [vendQ] : qs).flatMap(q => lotesDe(q).map(l => infoVendaLote(q, l)));
    const fq = t => t.destino === 'Venda' && (!vendQ || t.quadra === vendQ);
    const aRec = sum(S.list.filter(t => !pago(t) && fq(t)), t => t.valor), recebido = sum(S.list.filter(t => pago(t) && fq(t)), t => t.valor);
    const proxT = S.list.filter(t => !pago(t) && fq(t) && t.valor > 0).sort((a, b) => a.venc.localeCompare(b.venc))[0];
    const parados = lots.filter(x => x.v && x.v.data && !['disp', 'recebido'].includes(x.etapa) && x.v.data < addDays(h, -30));
    const sug = lots.filter(x => x.sugerida && x.etapa === 'contrato').length;
    el.innerHTML = `<div class="filtros card"><label class="tgl">Quadra <select id="vq">${opts(qs, vendQ, 'Todas')}</select></label><span class="muted small">Toque num lote para registrar cliente, etapa e datas.${sug ? ` ${sug} lote(s) estão com etapa <b>sugerida</b> a partir dos lançamentos: confirme quando puder.` : ''}</span></div>
    <div class="grid kpis k4">${kpi('A receber em vendas', brl(aRec), S.list.filter(t => !pago(t) && fq(t)).length + ' parcela(s)', 'pos', '#/receber')}${kpi('Já recebido', brl(recebido), '', '')}${kpi('Próximo repasse', proxT ? dBR(proxT.venc) : '—', proxT ? esc(proxT.quadra + ' · ' + proxT.lote + ' · ' + brl(proxT.valor)) : '', '')}${kpi('Parados há +30 dias', parados.length, 'na mesma etapa', parados.length ? 'neg' : '')}</div>
    <div class="kanban">${ETAPAS_V.map(([k, nome, cor]) => {
      const cs = lots.filter(x => x.etapa === k);
      return `<section class="kcol"><h3><span style="display:flex;gap:8px;align-items:center"><span class="dia" style="--c:${cor}"></span>${nome}</span><span class="muted">${cs.length}</span></h3><div class="ktot">${short(sum(cs, x => x.total))}</div>
        ${cs.slice(0, 60).map(x => `<button type="button" class="kcard" data-act="venda" data-k="${esc(x.k)}"><span style="display:flex;justify-content:space-between;gap:8px"><b>${esc(x.q)} · ${esc(x.l)}</b><b>${x.total ? short(x.total) : ''}</b></span>
          <small>${esc(x.v && x.v.cliente || 'Sem cliente')}</small>
          <small style="color:${x.v && x.v.data && x.v.data < addDays(h, -30) && !['disp', 'recebido'].includes(k) ? 'var(--warn)' : 'var(--muted)'}">${k === 'recebido' ? (x.ult ? 'Recebido ' + dBR(x.ult) : 'Recebido') : x.prox ? 'Repasse previsto ' + dBR(x.prox) : x.v && x.v.data ? 'Desde ' + dBR(x.v.data) : ''}${x.sugerida && k === 'contrato' ? ' · sugerida' : ''}</small></button>`).join('')}
        ${cs.length > 60 ? `<span class="muted small" style="padding:4px">+ ${cs.length - 60} lotes</span>` : ''}</section>`;
    }).join('')}</div>`;
    $('#vq', el).onchange = e => { vendQ = e.target.value; this.refresh(el); };
  }
};
ACOES.venda = d => formVenda(d.k);
function formVenda(k) {
  const [q, l] = k.split('|'); const x = infoVendaLote(q, l); const v = x.v || {};
  const hist = (v.hist || []).slice().reverse();
  const f = folha(`Quadra ${q} · lote ${l}`, `<form class="form" id="fV">
    <div class="resumo"><div class="g3"><div><span class="muted small">Recebido</span><div class="pos"><b>${brl(x.rec)}</b></div></div><div><span class="muted small">A receber</span><div><b>${brl(x.aberto)}</b></div></div><div><span class="muted small">Próximo vencimento</span><div><b>${x.prox ? dBR(x.prox) : '—'}</b></div></div></div></div>
    <div class="g2"><label>Etapa<select name="etapa">${opts(ETAPAS_V.map(e => [e[0], e[1]]), x.etapa)}</select></label><label>Desde<input type="date" name="data" value="${esc(v.data || hoje())}"></label></div>
    <div class="g2"><label>Cliente<input name="cliente" value="${esc(v.cliente || '')}" autocomplete="off"></label><label>Telefone<input name="tel" type="tel" value="${esc(v.tel || '')}"></label></div>
    <label>Correspondente / agência<input name="corresp" value="${esc(v.corresp || '')}" list="dlCorr"></label>
    <label>Observação<textarea name="obs" rows="2">${esc(v.obs || '')}</textarea></label>
    <datalist id="dlCorr">${[...new Set(Object.values((S.vendas && S.vendas.lotes) || {}).map(o => o.corresp).filter(Boolean))].map(c => `<option value="${esc(c)}">`).join('')}</datalist>
    ${hist.length ? `<div class="lbl">Histórico</div><div class="tlist small">${hist.map(hh => `<div class="kv"><span>${dBR(hh.data)}</span><b>${esc(etapaNome(hh.etapa))}</b></div>`).join('')}</div>` : ''}
    <div class="acts"><a class="btn" href="#/obras/${encodeURIComponent(q)}/${encodeURIComponent(l)}" data-close>Ver lançamentos do lote</a><button type="button" class="btn" data-close>Cancelar</button><button class="btn primary rw">Salvar</button></div></form>`);
  const F = f.el.querySelector('#fV');
  F.onsubmit = e => {
    e.preventDefault();
    const novo = { etapa: F.etapa.value, data: F.data.value, cliente: F.cliente.value.trim(), tel: F.tel.value.trim(), corresp: F.corresp.value.trim(), obs: F.obs.value.trim(), hist: v.hist || [] };
    if (novo.etapa !== v.etapa) novo.hist = [...novo.hist, { etapa: novo.etapa, data: novo.data }];
    S.vendas = S.vendas || { lotes: {} }; S.vendas.lotes = { ...(S.vendas.lotes || {}), [k]: novo };
    gravar(doc(db, 'config', 'vendas'), { lotes: { [k]: novo } }, { merge: true }).catch(onErroEscrita);
    f.fechar(); toast('Lote ' + q + ' · ' + l + ' atualizado'); refreshView();
  };
}

/* =========================================================
   Conciliação bancária
   ========================================================= */
const hashTxt = s => { let h = 5381; for (let i = 0; i < s.length; i++) h = ((h << 5) + h + s.charCodeAt(i)) >>> 0; return h.toString(36); };
const contaId = c => String(c).replace(/[\/.#$\[\]]/g, '-') || 'conta';
function lerExtratoBanco(txt) {
  if (/<OFX>|OFXHEADER|<STMTTRN>/i.test(txt)) {
    const tag = (b, n) => { const m = b.match(new RegExp('<' + n + '>\\s*([^<\\r\\n]+)', 'i')); return m ? m[1].trim() : ''; };
    const linhas = txt.split(/<STMTTRN>/i).slice(1).map((b, i) => {
      const dt = tag(b, 'DTPOSTED'); const d = dt.slice(0, 4) + '-' + dt.slice(4, 6) + '-' + dt.slice(6, 8);
      const v = parseFloat(tag(b, 'TRNAMT').replace(',', '.'));
      const desc = [tag(b, 'NAME'), tag(b, 'MEMO')].filter(Boolean).filter((x, j, a) => a.indexOf(x) === j).join(' · ') || '(sem descrição)';
      const fit = tag(b, 'FITID') || (d + '|' + v + '|' + desc + '|' + i);
      return { id: 'k' + hashTxt(fit), d, desc, v };
    }).filter(l => /^\d{4}-\d{2}-\d{2}$/.test(l.d) && isFinite(l.v));
    const bal = txt.match(/<LEDGERBAL>[\s\S]*?<BALAMT>\s*([-\d.,]+)/i), dta = txt.match(/<LEDGERBAL>[\s\S]*?<DTASOF>\s*(\d{8})/i);
    return { linhas, saldo: bal ? parseFloat(bal[1].replace(',', '.')) : null, saldoData: dta ? dta[1].slice(0, 4) + '-' + dta[1].slice(4, 6) + '-' + dta[1].slice(6, 8) : '' };
  }
  // CSV: data; descrição; valor (separador ; ou ,)
  const rows = txt.split(/\r?\n/).filter(r => r.trim());
  const sep = (rows[0].match(/;/g) || []).length >= (rows[0].match(/,/g) || []).length ? ';' : ',';
  const cell = r => r.split(sep).map(c => c.replace(/^"|"$/g, '').trim());
  const linhas = [];
  rows.forEach((r, i) => {
    const c = cell(r); let d = '', v = NaN, desc = '';
    for (const x of c) { let m; if (!d && (m = x.match(/^(\d{2})\/(\d{2})\/(\d{4})/))) d = `${m[3]}-${m[2]}-${m[1]}`; else if (!d && /^\d{4}-\d{2}-\d{2}/.test(x)) d = x.slice(0, 10); }
    for (let j = c.length - 1; j >= 0; j--) { if (/^-?R?\$?\s?-?[\d.]+,\d{2}$|^-?\d+\.\d{2}$/.test(c[j])) { v = parseBR(c[j].replace(/R\$\s?/, '')); if (/^-/.test(c[j]) || /^R?\$?\s?-/.test(c[j])) v = -Math.abs(v); break; } }
    desc = c.filter(x => x && !/^[\d\/.,\-R$\s]+$/.test(x)).sort((a, b) => b.length - a.length)[0] || '(sem descrição)';
    if (d && isFinite(v)) linhas.push({ id: 'k' + hashTxt(d + '|' + v + '|' + desc + '|' + i), d, desc, v });
  });
  return { linhas, saldo: null, saldoData: '' };
}
function analisarConc(conta) {
  const ex = S.extratos && S.extratos[contaId(conta)]; if (!ex) return null;
  const linhas = Object.values(ex.linhas || {}).map(j => { try { return JSON.parse(j); } catch (e) { return null; } }).filter(Boolean).sort((a, b) => b.d.localeCompare(a.d));
  if (!linhas.length) return { ex, linhas: [] };
  const ini = addDays(linhas[linhas.length - 1].d, -40), fim = addDays(linhas[0].d, 40);
  const ts = S.list.filter(t => { const d = dataRef(t) || ''; return d >= ini && d <= fim; });
  const conc = new Map(); for (const t of S.list) if (t.conc) (conc.get(t.conc) || conc.set(t.conc, []).get(t.conc)).push(t);
  const livres = ts.filter(t => !t.conc);
  const grupos = new Map(); for (const t of livres) { if (!t.lote || t.lote === '-') continue; const k = [t.status, dataRef(t), t.quadra, t.destino, Math.sign(t.valor), baseItem(t)].join('|'); (grupos.get(k) || grupos.set(k, []).get(k)).push(t); }
  const usado = new Set(); const dias = (a, b) => Math.abs((new Date(a) - new Date(b)) / 864e5);
  const val = t => t.valor + (pago(t) ? (t.juros || 0) : 0);
  const out = linhas.map(l => {
    if (conc.has(l.id)) return { ...l, st: 'ok', ts: conc.get(l.id) };
    if (l.ign) return { ...l, st: 'ign', ts: [] };
    let c = livres.find(t => !usado.has(t.id) && pago(t) && (t.conta || '') === conta && Math.abs(val(t) - l.v) < 0.01 && dias(t.pag || t.venc, l.d) <= 5);
    if (c) { usado.add(c.id); return { ...l, st: 'auto', ts: [c] }; }
    for (const g of grupos.values()) if (g.length > 1 && pago(g[0]) && (g[0].conta || '') === conta && !g.some(t => usado.has(t.id)) && Math.abs(sum(g, val) - l.v) < 0.01 && dias(g[0].pag || g[0].venc, l.d) <= 5) { g.forEach(t => usado.add(t.id)); return { ...l, st: 'auto', ts: g }; }
    c = livres.find(t => !usado.has(t.id) && !pago(t) && Math.abs(t.valor - l.v) < 0.01 && dias(t.venc, l.d) <= 25);
    if (c) { usado.add(c.id); return { ...l, st: 'sug', ts: [c] }; }
    for (const g of grupos.values()) if (g.length > 1 && !pago(g[0]) && !g.some(t => usado.has(t.id)) && Math.abs(sum(g, t => t.valor) - l.v) < 0.01 && dias(g[0].venc, l.d) <= 25) { g.forEach(t => usado.add(t.id)); return { ...l, st: 'sug', ts: g }; }
    return { ...l, st: 'novo', ts: [] };
  });
  return { ex, linhas: out };
}
let concConta = '', concTodas = false, concCache = null;
const ST_CONC = { ok: ['Conciliada', 'ok'], auto: ['Encontrada', 'ok'], sug: ['Sugestão', 'warn'], novo: ['Sem lançamento', 'bad'], ign: ['Ignorada', ''] };
VIEWS.conciliacao = {
  titulo: 'Conciliação bancária',
  refresh(el) {
    const contas = E().contas.map(c => c.nome); if (!contas.includes(concConta)) concConta = contas[0];
    const a = analisarConc(concConta); concCache = a;
    const sd = saldos().contas.find(c => c.nome === concConta);
    const ls = a ? a.linhas : [];
    const pend = ls.filter(l => ['sug', 'novo', 'auto'].includes(l.st));
    const vis = concTodas ? ls : pend;
    el.innerHTML = `<div class="filtros card"><label class="tgl">Conta <select id="cc">${opts(contas, concConta)}</select></label>
      <label class="btn primary file rw">Importar extrato (OFX ou CSV)<input type="file" id="cfile" accept=".ofx,.OFX,.csv,.txt" hidden></label>
      ${ls.some(l => l.st === 'auto') ? `<button class="btn rw" id="cAll">Confirmar todas encontradas (${ls.filter(l => l.st === 'auto').length})</button>` : ''}
      <label class="tgl"><input type="checkbox" id="cTodas" ${concTodas ? 'checked' : ''}> Mostrar já conciliadas</label></div>
    ${!a ? `<div class="card empty"><h2>Nenhum extrato importado para ${esc(concConta)}</h2><p class="muted">No app ou internet banking do banco, exporte o extrato do período em <b>OFX</b> (ou CSV) e importe aqui. O sistema casa cada linha com os lançamentos, inclusive rateios por lote.</p></div>` : `
    <div class="grid kpis k4">${kpi('Saldo no banco', a.ex.saldo != null ? brl2(a.ex.saldo) : '—', a.ex.saldoData ? 'em ' + dBR(a.ex.saldoData) : 'o arquivo não trouxe saldo', '')}${kpi('Saldo no sistema', brl2(sd ? sd.saldo : 0), esc(concConta), '')}${kpi('Conciliadas', ls.filter(l => l.st === 'ok').length + ' de ' + ls.filter(l => l.st !== 'ign').length, '', 'pos')}${kpi('Precisam de você', pend.length, ls.filter(l => l.st === 'sug').length + ' sugestão(ões) · ' + ls.filter(l => l.st === 'novo').length + ' sem lançamento', pend.length ? 'neg' : '')}</div>
    ${a.ex.saldo != null && sd && Math.abs(a.ex.saldo - sd.saldo) > 0.01 && !pend.length ? `<div class="alert">Tudo conciliado, mas o saldo do banco difere do sistema em ${brl2(a.ex.saldo - sd.saldo)}. Pode haver lançamentos anteriores ao período do extrato ou saldo inicial da conta a ajustar em Cadastros.</div>` : ''}
    <div class="card"><div class="tbl-wrap"><table class="tbl ftbl"><thead><tr><th class="l">Data</th><th class="l">No banco</th><th>Valor</th><th class="l">Situação</th><th class="l">No sistema</th><th></th></tr></thead><tbody>
    ${vis.map(l => { const [n, c] = ST_CONC[l.st]; const t0 = l.ts[0]; return `<tr><td class="l">${dBR(l.d)}</td><td class="l wrap">${esc(l.desc)}</td><td class="${cls(l.v)}"><b>${brl2(l.v)}</b></td><td class="l st"><span class="chip ${c}">${n}</span></td>
      <td class="l wrap">${t0 ? `<a href="#" data-act="abrir" data-id="${t0.id}">${esc(l.ts.length > 1 ? baseItem(t0) + ' · ' + l.ts.length + ' lotes' : t0.item)}</a><br><span class="muted small">${pago(t0) ? 'Pago ' + dBR(t0.pag) : 'Em aberto · vence ' + dBR(t0.venc)}</span>` : l.st === 'novo' ? '<span class="muted">Nenhum lançamento com este valor e data</span>' : ''}</td>
      <td class="rw">${l.st === 'auto' ? `<button class="btn sm" data-act="cOk" data-id="${l.id}">Confirmar</button>` : l.st === 'sug' ? `<button class="btn sm primary" data-act="cBx" data-id="${l.id}">Dar baixa</button>` : l.st === 'novo' ? `<button class="btn sm" data-act="cNovo" data-id="${l.id}">Criar lançamento</button> <button class="btn sm link" data-act="cIgn" data-id="${l.id}">Ignorar</button>` : l.st === 'ok' ? `<button class="btn sm link" data-act="cDes" data-id="${l.id}">Desfazer</button>` : `<button class="btn sm link" data-act="cIgn" data-id="${l.id}" data-v="0">Reativar</button>`}</td></tr>`; }).join('') || '<tr><td colspan="6" class="l muted">Nada pendente. Tudo conciliado.</td></tr>'}
    </tbody></table></div></div>`}`;
    $('#cc', el).onchange = e => { concConta = e.target.value; this.refresh(el); };
    $('#cTodas', el).onchange = e => { concTodas = e.target.checked; this.refresh(el); };
    const fi = $('#cfile', el); if (fi) fi.onchange = async e => {
      const file = e.target.files[0]; if (!file) return;
      try {
        const buf = new Uint8Array(await file.arrayBuffer());
        let txt = new TextDecoder('utf-8').decode(buf); if (txt.includes('\uFFFD')) txt = new TextDecoder('windows-1252').decode(buf);
        const r = lerExtratoBanco(txt); if (!r.linhas.length) { toast('Não encontrei lançamentos nesse arquivo.', 'err'); return; }
        const linhas = {}; const atuais = (S.extratos && S.extratos[contaId(concConta)] && S.extratos[contaId(concConta)].linhas) || {};
        r.linhas.forEach(l => { if (!atuais[l.id]) linhas[l.id] = JSON.stringify(l); });
        const dados = { conta: concConta, linhas, importado: Date.now() }; if (r.saldo != null) { dados.saldo = r.saldo; dados.saldoData = r.saldoData; }
        S.extratos = S.extratos || {}; const cur = S.extratos[contaId(concConta)] || { linhas: {} }; S.extratos[contaId(concConta)] = { ...cur, ...dados, linhas: { ...(cur.linhas || {}), ...linhas } };
        gravar(doc(db, 'extratos', contaId(concConta)), dados, { merge: true }).catch(onErroEscrita);
        toast(`${r.linhas.length} linhas lidas · ${Object.keys(linhas).length} novas`); this.refresh(el);
      } catch (err) { console.error(err); toast('Não consegui ler o extrato: ' + err.message, 'err'); }
    };
    const ca = $('#cAll', el); if (ca) ca.onclick = () => { const out = []; ls.filter(l => l.st === 'auto').forEach(l => l.ts.forEach(t => out.push({ ...t, conc: l.id }))); salvarTitulos(out); toast(out.length + ' lançamento(s) conciliados'); };
  }
};
const linhaConc = id => concCache && concCache.linhas.find(l => l.id === id);
ACOES.cOk = d => { const l = linhaConc(d.id); if (!l) return; salvarTitulos(l.ts.map(t => ({ ...t, conc: l.id }))); toast('Conciliado'); };
ACOES.cDes = d => { const l = linhaConc(d.id); if (!l) return; salvarTitulos(l.ts.map(t => { const n = { ...t }; delete n.conc; return n; })); };
ACOES.cBx = async d => {
  const l = linhaConc(d.id); if (!l) return;
  if (!(await confirmar(`Registrar ${l.ts.length > 1 ? 'os <b>' + l.ts.length + '</b> lançamentos' : '<b>' + esc(l.ts[0].item) + '</b>'} como ${l.v < 0 ? 'pago' : 'recebido'} em <b>${dBR(l.d)}</b> na conta <b>${esc(concConta)}</b>?`, 'Dar baixa', false))) return;
  salvarTitulos(l.ts.map(t => ({ ...t, status: 'pago', pag: l.d, conta: concConta, conc: l.id }))); toast('Baixa registrada e conciliada');
};
ACOES.cNovo = d => { const l = linhaConc(d.id); if (!l) return; formNovo({ tipo: l.v < 0 ? 'saida' : 'entrada', realizado: true, item: l.desc, valor: Math.abs(l.v), venc: l.d, pag: l.d, conta: concConta, conc: l.id, destino: l.v < 0 ? 'Administrativo' : 'Aporte', quadra: '-' }); };
ACOES.cIgn = d => {
  const ex = S.extratos && S.extratos[contaId(concConta)]; if (!ex || !ex.linhas[d.id]) return;
  const l = JSON.parse(ex.linhas[d.id]); l.ign = d.v !== '0'; if (!l.ign) delete l.ign; ex.linhas[d.id] = JSON.stringify(l);
  gravar(doc(db, 'extratos', contaId(concConta)), { linhas: { [d.id]: ex.linhas[d.id] } }, { merge: true }).catch(onErroEscrita); refreshView();
};

/* =========================================================
   Cronograma físico-financeiro
   ========================================================= */
const ETAPAS_OBRA = ['Não iniciada', 'Terraplenagem', 'Fundação', 'Alvenaria', 'Cobertura', 'Instalações', 'Acabamento', 'Entregue'];
function dadosObra(q) {
  const o = ((S.obras && S.obras.quadras) || {})[q] || {};
  const med = (o.medicoes || []).slice().sort((a, b) => a.data.localeCompare(b.data));
  return { etapa: o.etapa || '', med, fis: med.length ? med[med.length - 1].pct : null };
}
VIEWS.cronograma = {
  titulo: 'Cronograma físico-financeiro',
  refresh(el) {
    const h = hoje(); const meses = [...Array(14)].map((_, i) => addMonths(h.slice(0, 7) + '-01', i).slice(0, 7));
    const rows = E().quadras.filter(q => lotesDe(q).length).map(q => {
      const r = resumoQuadra(q), o = dadosObra(q);
      const fin = r.prevCusto ? r.custo / r.prevCusto : null, comp = r.prevCusto ? (r.custo + r.pagar) / r.prevCusto : null;
      const rep = meses.map(m => sum(S.list.filter(t => !pago(t) && t.quadra === q && t.destino === 'Venda' && (t.venc < h ? h : t.venc).startsWith(m)), t => t.valor));
      const entregue = o.etapa === 'Entregue' || (!o.etapa && r.receber < 0.5 && r.receita > 0);
      return { q, n: r.n, o, fin, comp, rep, entregue, alerta: o.fis !== null && fin !== null && fin * 100 - o.fis > 10 };
    });
    el.innerHTML = `<div class="card"><div class="legend"><span style="--c:var(--navy)">Físico (medição na obra)</span><span style="--c:var(--steel)">Financeiro (gasto ÷ orçado)</span><span style="--c:#1F7A57">Repasse previsto</span></div>
    <div class="tbl-wrap"><table class="tbl"><thead><tr><th class="l">Quadra</th><th>Lotes</th><th class="l">Etapa</th><th class="l">Físico × financeiro</th><th>Comprometido</th><th class="l"><div class="gant" style="--n:${meses.length}">${meses.map(m => `<span>${MES3[+m.slice(5, 7) - 1]}</span>`).join('')}</div></th><th></th></tr></thead><tbody>
    ${rows.map(r => `<tr><td class="l"><b>${esc(r.q)}</b></td><td>${r.n}</td>
      <td class="l"><span class="chip ${r.entregue ? 'ok' : ''}">${esc(r.o.etapa || (r.entregue ? 'Entregue' : 'Não informada'))}</span>${r.alerta ? '<br><span class="chip warn" style="margin-top:4px">Gasto à frente da obra</span>' : ''}</td>
      <td class="l"><div class="dual"><span class="pbar" style="justify-content:flex-start"><span class="bar"><i style="width:${r.o.fis !== null ? r.o.fis : (r.entregue ? 100 : 0)}%;background:var(--navy)"></i></span><span>${r.o.fis !== null ? fPct.format(r.o.fis / 100) : r.entregue ? '100%' : '—'}</span></span>
        <span class="pbar" style="justify-content:flex-start"><span class="bar"><i style="width:${Math.min(100, (r.fin || 0) * 100)}%;background:${r.fin > 1 ? 'var(--out)' : 'var(--steel)'}"></i></span><span class="${r.fin > 1 ? 'neg' : ''}">${pct(r.fin)}</span></span></div></td>
      <td class="${r.comp > 1 ? 'neg' : ''}">${pct(r.comp)}</td>
      <td class="l"><div class="gant" style="--n:${meses.length}">${r.rep.map((v, i) => `<span>${v > 0.5 ? `<i class="dia" title="${esc(MES3[+meses[i].slice(5, 7) - 1] + '/' + meses[i].slice(2, 4) + ': ' + brl(v))}"></i>` : ''}</span>`).join('')}</div></td>
      <td><button class="btn sm rw" data-act="medicao" data-q="${esc(r.q)}">Medição</button></td></tr>`).join('')}
    </tbody></table></div><p class="muted small">Financeiro = custo já pago ÷ orçamento previsto da quadra. Comprometido inclui o que está em contas a pagar. Repasses: contas a receber de venda por mês.</p></div>`;
  }
};
ACOES.medicao = d => {
  const q = d.q, o = dadosObra(q);
  const f = folha('Medição · quadra ' + q, `<form class="form" id="fM">
    <div class="g3"><label>Etapa atual<select name="etapa">${opts(ETAPAS_OBRA, o.etapa || 'Fundação')}</select></label><label>% físico executado<input name="pct" type="number" min="0" max="100" step="1" inputmode="numeric" value="${o.fis ?? ''}" required></label><label>Data<input type="date" name="data" value="${hoje()}"></label></div>
    <label>Observação<input name="obs" placeholder="Ex.: 12 casas na cobertura"></label>
    ${o.med.length ? `<div class="lbl">Medições anteriores</div>${o.med.slice().reverse().map(m => `<div class="kv"><span>${dBR(m.data)}</span><b>${m.pct}%</b><span class="muted small">${esc(m.obs || '')}</span></div>`).join('')}` : ''}
    <div class="acts"><button type="button" class="btn" data-close>Cancelar</button><button class="btn primary">Salvar medição</button></div></form>`);
  const F = f.el.querySelector('#fM');
  F.onsubmit = e => {
    e.preventDefault(); const p = Math.max(0, Math.min(100, +F.pct.value || 0));
    const med = [...o.med.filter(m => m.data !== F.data.value), { data: F.data.value, pct: p, obs: F.obs.value.trim() }];
    const etapa = p >= 100 ? 'Entregue' : F.etapa.value;
    S.obras = S.obras || { quadras: {} }; S.obras.quadras = { ...(S.obras.quadras || {}), [q]: { etapa, medicoes: med } };
    gravar(doc(db, 'config', 'obras'), { quadras: { [q]: { etapa, medicoes: med } } }, { merge: true }).catch(onErroEscrita);
    f.fechar(); toast('Medição registrada'); refreshView();
  };
};

/* =========================================================
   Simulador de viabilidade
   ========================================================= */
const SIM_CAMPOS = [['lotes', 'Número de lotes', 1, ''], ['venda', 'Preço de venda por casa (R$)', 1, 'R$'], ['obra', 'Custo de obra por casa (R$)', 1, 'R$'], ['terreno', 'Terreno por lote (R$)', 1, 'R$'], ['corr', 'Corretagem (% da venda)', 0, '%'], ['imp', 'Impostos (% da venda)', 0, '%'], ['adm', 'Administrativo por lote (R$)', 1, 'R$']];
function simDeQuadra(q) {
  const P = previstoLote(q); const sec = secoesViab(); const venda = P['Venda'] || 0;
  const obra = -sum(sec.op.filter(k => k !== 'Terreno' && k !== 'Juros e Multa'), k => P[k] || 0);
  return { nome: 'Base ' + q, lotes: lotesDe(q).length || 1, venda, obra, terreno: -(P['Terreno'] || 0), corr: venda ? r2(-(P['Corretagem'] || 0) / venda * 100) : 5, imp: venda ? r2(-(P['Impostos'] || 0) / venda * 100) : 2.5, adm: -((P['Administrativo'] || 0) + (P['Caixa'] || 0) + (P['Aditivo'] || 0)) };
}
function calcSim(p) {
  const corr = p.venda * p.corr / 100, imp = p.venda * p.imp / 100, custo = p.obra + p.terreno + p.adm;
  const lucro = p.venda - custo - corr - imp;
  return { corr, imp, lucro, total: lucro * p.lotes, vgv: p.venda * p.lotes, custoTotal: (custo + corr + imp) * p.lotes, margem: p.venda ? lucro / p.venda : null, roi: custo ? lucro / custo : null, minimo: (1 - (p.corr + p.imp) / 100) > 0 ? custo / (1 - (p.corr + p.imp) / 100) : null };
}
let simAtual = null;
VIEWS.simulador = {
  titulo: 'Simulador de viabilidade',
  montar(el) {
    if (!simAtual) { const q = E().quadras.includes('K3') ? 'K3' : E().quadras[0]; simAtual = q ? simDeQuadra(q) : { nome: 'Novo empreendimento', lotes: 40, venda: 190000, obra: 110000, terreno: 25000, corr: 4.7, imp: 2.6, adm: 250 }; }
    el.innerHTML = `<div class="grid lg"><div class="card"><h2>Premissas <label class="tgl" style="margin-left:auto">Carregar orçamento de <select id="sQ">${opts(E().quadras, '', 'quadra…')}</select></label></h2>
      <form class="form" id="fSim"><label>Nome do cenário<input name="nome" value="${esc(simAtual.nome)}"></label>
      <div class="g2">${SIM_CAMPOS.map(([k, n, , u]) => `<label>${n}<input name="${k}" inputmode="decimal" value="${u === '%' ? String(simAtual[k]).replace('.', ',') : k === 'lotes' ? simAtual[k] : fNum2.format(simAtual[k])}"></label>`).join('')}</div>
      <div class="acts left rw"><button class="btn primary" type="submit">Salvar cenário</button></div></form></div>
      <div class="card" id="simRes"></div></div><div class="card"><h2>Cenários salvos</h2><div id="simLista"></div></div>`;
    const F = $('#fSim', el);
    const ler = () => { const p = { nome: F.nome.value.trim() || 'Cenário' }; for (const [k] of SIM_CAMPOS) { const v = parseBR(F[k].value); p[k] = isFinite(v) ? v : 0; } return p; };
    const calc = () => { simAtual = ler(); const p = simAtual, r = calcSim(p);
      $('#simRes', el).innerHTML = `<h2>Resultado</h2><div class="grid kpis" style="grid-template-columns:repeat(2,minmax(0,1fr))">
        <div class="card kpi hero"><div class="lab">Lucro do empreendimento</div><div class="val">${short(r.total)}</div><div class="sub">${p.lotes} lote(s)</div></div>
        ${kpi('Lucro por casa', brl(r.lucro), '', cls(r.lucro))}${kpi('Margem líquida', pct(r.margem), 'sobre a venda', '')}${kpi('ROI', pct(r.roi), 'sobre o custo', '')}</div>
        <table class="tbl"><tbody>
        <tr><td class="l">VGV (venda total)</td><td>${brl(r.vgv)}</td></tr><tr><td class="l">Custo total (com corretagem e impostos)</td><td>${brl(-r.custoTotal)}</td></tr>
        <tr><td class="l">Por casa: obra + terreno + adm.</td><td>${brl(-(p.obra + p.terreno + p.adm))}</td></tr><tr><td class="l">Por casa: corretagem + impostos</td><td>${brl(-(r.corr + r.imp))}</td></tr>
        <tr class="tot"><td class="l">Preço mínimo para empatar</td><td>${r.minimo ? brl(r.minimo) : '—'}</td></tr></tbody></table>`; };
    F.addEventListener('input', calc);
    F.onsubmit = e => { e.preventDefault(); const p = ler(); const id = p.id || uid(); const sv = { ...p, id, salvo: Date.now() }; S.simulacoes = S.simulacoes || { cenarios: {} }; S.simulacoes.cenarios = { ...(S.simulacoes.cenarios || {}), [id]: sv }; gravar(doc(db, 'config', 'simulacoes'), { cenarios: { [id]: sv } }, { merge: true }).catch(onErroEscrita); toast('Cenário salvo'); this.refresh(el); };
    $('#sQ', el).onchange = e => { if (!e.target.value) return; simAtual = simDeQuadra(e.target.value); this.montar(el); this.refresh(el); };
    calc();
  },
  refresh(el) {
    const cs = Object.values((S.simulacoes && S.simulacoes.cenarios) || {}).sort((a, b) => (b.salvo || 0) - (a.salvo || 0));
    const box = $('#simLista', el); if (!box) return;
    box.innerHTML = cs.length ? `<div class="tbl-wrap"><table class="tbl"><thead><tr><th class="l">Cenário</th><th>Lotes</th><th>Venda</th><th>Lucro/casa</th><th>Margem</th><th>Lucro total</th><th></th></tr></thead><tbody>${cs.map(c => { const r = calcSim(c); return `<tr><td class="l"><b>${esc(c.nome)}</b></td><td>${c.lotes}</td><td>${brl(c.venda)}</td><td class="${cls(r.lucro)}">${brl(r.lucro)}</td><td>${pct(r.margem)}</td><td>${short(r.total)}</td><td><button class="btn sm" data-act="simAbrir" data-id="${c.id}">Abrir</button> <button class="btn sm link rw" data-act="simDel" data-id="${c.id}">Excluir</button></td></tr>`; }).join('')}</tbody></table></div>` : '<p class="muted">Nenhum cenário salvo ainda.</p>';
  }
};
ACOES.simAbrir = d => { const c = S.simulacoes?.cenarios?.[d.id]; if (!c) return; simAtual = { ...c }; VIEWS.simulador.montar($('#view')); VIEWS.simulador.refresh($('#view')); window.scrollTo(0, 0); };
ACOES.simDel = async d => { const c = S.simulacoes?.cenarios?.[d.id]; if (!c || !(await confirmar(`Excluir o cenário <b>${esc(c.nome)}</b>?`, 'Excluir'))) return; delete S.simulacoes.cenarios[d.id]; gravar(doc(db, 'config', 'simulacoes'), { cenarios: { [d.id]: deleteField() } }, { merge: true }).catch(onErroEscrita); refreshView(); };

/* =========================================================
   Fornecedores e histórico de preço
   ========================================================= */
function statsFornecedores() {
  const m = new Map(); const cad = (S.fornecedores && S.fornecedores.lista) || {};
  for (const n of Object.keys(cad)) m.set(n, { nome: n, cad: cad[n], pago: 0, aberto: 0, n: 0, ult: '' });
  for (const t of S.list) { if (!t.forn) continue; let s = m.get(t.forn); if (!s) { s = { nome: t.forn, cad: null, pago: 0, aberto: 0, n: 0, ult: '' }; m.set(t.forn, s); } s.n++; if (pago(t)) { s.pago += t.valor + (t.juros || 0); const d = t.pag || t.venc; if (d > s.ult) s.ult = d; } else s.aberto += t.valor; }
  return [...m.values()].sort((a, b) => a.pago - b.pago);
}
let fornQ = '';
VIEWS.fornecedores = {
  titulo: 'Fornecedores',
  montar(el) {
    el.innerHTML = `<div class="filtros card"><input type="search" id="fnq" placeholder="Buscar fornecedor…" value="${esc(fornQ)}"><button class="btn primary rw" data-act="fornNovo">+ Novo fornecedor</button></div><div id="fnBox"></div>`;
    $('#fnq', el).oninput = e => { fornQ = e.target.value; this.refresh(el); };
  },
  refresh(el) {
    const q = fornQ.trim().toLowerCase(); const fs = statsFornecedores().filter(f => !q || f.nome.toLowerCase().includes(q));
    const semForn = S.list.filter(t => !t.forn && t.valor < 0).length;
    $('#fnBox', el).innerHTML = `${semForn ? `<div class="alert">${semForn.toLocaleString('pt-BR')} saídas ainda não têm fornecedor. Abra um fornecedor e use <b>Vincular lançamentos pela descrição</b> para preencher de uma vez (ex.: tudo que tem "Mendonça" na descrição).</div>` : ''}
    <div class="card"><div class="tbl-wrap"><table class="tbl"><thead><tr><th class="l">Fornecedor</th><th class="l">Categoria</th><th>Lançamentos</th><th>Total pago</th><th>Em aberto</th><th>Última compra</th></tr></thead><tbody>
    ${fs.map(f => `<tr class="click" data-act="forn" data-n="${esc(f.nome)}"><td class="l"><b>${esc(f.nome)}</b></td><td class="l muted">${esc(f.cad?.cat || '')}</td><td>${f.n}</td><td class="neg">${brl(f.pago)}</td><td>${f.aberto ? brl(f.aberto) : '—'}</td><td>${f.ult ? dBR(f.ult) : '—'}</td></tr>`).join('') || '<tr><td colspan="6" class="l muted">Nenhum fornecedor ainda. Cadastre um ou preencha o campo Fornecedor nos lançamentos.</td></tr>'}
    </tbody></table></div></div>`;
  }
};
function salvarForn(nome, dados) { S.fornecedores = S.fornecedores || { lista: {} }; S.fornecedores.lista = { ...(S.fornecedores.lista || {}), [nome]: dados }; gravar(doc(db, 'config', 'fornecedores'), { lista: { [nome]: dados } }, { merge: true }).catch(onErroEscrita); }
ACOES.fornNovo = () => {
  const f = folha('Novo fornecedor', `<form class="form" id="fF"><label>Nome<input name="nome" required></label><div class="g2"><label>CNPJ / CPF<input name="cnpj"></label><label>Telefone<input name="tel" type="tel"></label></div><label>Categoria<select name="cat">${opts(todosDestinos(), 'Material')}</select></label><label>Observação<input name="obs"></label><div class="acts"><button type="button" class="btn" data-close>Cancelar</button><button class="btn primary">Salvar</button></div></form>`);
  f.el.querySelector('#fF').onsubmit = e => { e.preventDefault(); const F = e.target; const n = F.nome.value.trim(); salvarForn(n, { cnpj: F.cnpj.value.trim(), tel: F.tel.value.trim(), cat: F.cat.value, obs: F.obs.value.trim() }); f.fechar(); refreshView(); ACOES.forn({ n }); };
};
ACOES.forn = d => {
  const nome = d.n; const cad = ((S.fornecedores && S.fornecedores.lista) || {})[nome] || {};
  const ts = S.list.filter(t => t.forn === nome).sort((a, b) => (dataRef(b) || '').localeCompare(dataRef(a) || ''));
  // compras com quantidade: agrupa rateios
  const compras = new Map(); for (const t of ts) { if (!(t.qtd > 0)) continue; const k = t.grupo || t.id; const c = compras.get(k) || { item: baseItem(t).replace(/\s*\(\d+\/\d+\)\s*$/, ''), d: t.venc, qtd: 0, valor: 0, unid: t.unid || '' }; c.qtd += t.qtd; c.valor += Math.abs(t.valor); compras.set(k, c); }
  const cs = [...compras.values()].sort((a, b) => b.d.localeCompare(a.d));
  const f = folha(nome, `<div class="grid kpis" style="grid-template-columns:repeat(3,minmax(0,1fr))">${kpi('Total pago', brl(sum(ts.filter(pago), t => t.valor + (t.juros || 0))), '', 'neg')}${kpi('Em aberto', brl(sum(ts.filter(t => !pago(t)), t => t.valor)), '', '')}${kpi('Lançamentos', ts.length, '', '')}</div>
    <form class="form card" id="fF"><div class="g2"><label>CNPJ / CPF<input name="cnpj" value="${esc(cad.cnpj || '')}"></label><label>Telefone<input name="tel" type="tel" value="${esc(cad.tel || '')}"></label></div><div class="g2"><label>Categoria<select name="cat">${opts(todosDestinos(), cad.cat || 'Material')}</select></label><label>Observação<input name="obs" value="${esc(cad.obs || '')}"></label></div><div class="acts left rw"><button class="btn">Salvar cadastro</button></div></form>
    <div class="card rw"><h2>Vincular lançamentos pela descrição</h2><div class="filtros" style="padding:0"><input id="vtxt" value="${esc(nome.split(' ')[0])}" placeholder="texto da descrição"><button class="btn" id="vBtn">Vincular</button><span id="vPrev" class="muted small"></span></div></div>
    <div class="card"><h2>Histórico de preço <span class="hint">lançamentos com quantidade informada</span></h2>${cs.length ? `${cs.length > 1 ? '<div class="chartbox" style="height:200px"><canvas id="cPreco"></canvas></div>' : ''}<div class="tbl-wrap"><table class="tbl"><thead><tr><th class="l">Data</th><th class="l">Item</th><th>Qtd.</th><th>Total</th><th>Preço unitário</th></tr></thead><tbody>${cs.map(c => `<tr><td class="l">${dBR(c.d)}</td><td class="l">${esc(c.item)}</td><td>${c.qtd.toLocaleString('pt-BR', { maximumFractionDigits: 2 })} ${esc(c.unid)}</td><td>${brl2(c.valor)}</td><td><b>${brl2(c.valor / c.qtd)}</b>${c.unid ? ' / ' + esc(c.unid) : ''}</td></tr>`).join('')}</tbody></table></div>` : '<p class="muted small">Para acompanhar preço, preencha <b>Quantidade</b> e <b>Unidade</b> ao lançar uma compra (ex.: 200 sacos de cimento).</p>'}</div>
    <div class="card"><h2>Lançamentos</h2><div class="tlist">${ts.slice(0, 40).map(t => linhaTitulo(t, { mostrarPago: true })).join('') || '<p class="muted">Nenhum lançamento.</p>'}</div></div>`, { largo: true });
  f.el.querySelector('#fF').onsubmit = e => { e.preventDefault(); const F = e.target; salvarForn(nome, { cnpj: F.cnpj.value.trim(), tel: F.tel.value.trim(), cat: F.cat.value, obs: F.obs.value.trim() }); toast('Cadastro salvo'); };
  const vt = f.el.querySelector('#vtxt'), vp = f.el.querySelector('#vPrev');
  const cand = () => { const s = vt.value.trim().toLowerCase(); return s.length < 3 ? [] : S.list.filter(t => !t.forn && t.item.toLowerCase().includes(s)); };
  const pv = () => { const c = cand(); vp.textContent = c.length ? `${c.length} lançamento(s) sem fornecedor contêm "${vt.value.trim()}" (${brl(sum(c, t => t.valor))})` : 'Nenhum lançamento encontrado'; };
  vt.oninput = pv; pv();
  f.el.querySelector('#vBtn').onclick = async () => { const c = cand(); if (!c.length) return; if (await confirmar(`Definir <b>${esc(nome)}</b> como fornecedor de <b>${c.length}</b> lançamento(s)?`, 'Vincular', false)) { salvarTitulos(c.map(t => ({ ...t, forn: nome }))); f.fechar(); toast(c.length + ' lançamentos vinculados'); } };
  if (cs.length > 1) { const ord = cs.slice().reverse(); grafico('cPreco', { type: 'line', data: { labels: ord.map(c => dBR(c.d)), datasets: [{ data: ord.map(c => c.valor / c.qtd), borderColor: css('--bal'), backgroundColor: css('--bal'), tension: 0.2 }] }, options: { responsive: true, maintainAspectRatio: false, animation: false, plugins: { legend: { display: false }, tooltip: { callbacks: { label: c => brl2(c.parsed.y) } } }, scales: { y: { ticks: { callback: v => brl2(v), color: css('--muted'), maxTicksLimit: 5 }, grid: { color: css('--line') } }, x: { ticks: { color: css('--muted') }, grid: { display: false } } } } }); }
};

/* =========================================================
   Lembretes no calendário (.ics)
   ========================================================= */
ACOES.ics = () => {
  const h = hoje(); const ab = S.list.filter(t => !pago(t) && t.venc >= h && t.venc <= addDays(h, 90));
  if (!ab.length) { toast('Nenhum vencimento nos próximos 90 dias.'); return; }
  const escI = s => String(s).replace(/\\/g, '\\\\').replace(/;/g, '\\;').replace(/,/g, '\\,').replace(/\n/g, '\\n');
  const fold = l => { const out = []; while (l.length > 73) { out.push(l.slice(0, 73)); l = ' ' + l.slice(73); } out.push(l); return out.join('\r\n'); };
  const stamp = new Date().toISOString().replace(/[-:]/g, '').slice(0, 15) + 'Z';
  const ev = []; const dias = [...new Set(ab.map(t => t.venc))].sort();
  for (const d of dias) for (const tipo of ['pagar', 'receber']) {
    const ts = ab.filter(t => t.venc === d && (tipo === 'pagar' ? t.valor < 0 : t.valor > 0)); if (!ts.length) continue;
    const g = agrupar(ts, 'venc'); const tot = sum(ts, t => t.valor);
    const desc = g.slice(0, 20).map(e => e.t ? `${e.t.item}: ${brl2(e.t.valor)}` : `${baseItem(e.itens[0])} (${e.itens.length} lotes): ${brl2(sum(e.itens, x => x.valor))}`).join('\n') + (g.length > 20 ? `\n+ ${g.length - 20} itens` : '');
    const ymd = d.replace(/-/g, ''); const fimD = addDays(d, 1).replace(/-/g, '');
    ev.push(['BEGIN:VEVENT', `UID:sfg-${tipo}-${ymd}@sfgincorporadora`, `DTSTAMP:${stamp}`, `DTSTART;VALUE=DATE:${ymd}`, `DTEND;VALUE=DATE:${fimD}`,
      fold(`SUMMARY:${escI(`SFG: ${tipo === 'pagar' ? 'pagar' : 'receber'} ${brl(Math.abs(tot))} (${g.length})`)}`), fold(`DESCRIPTION:${escI(desc)}`),
      'BEGIN:VALARM', 'ACTION:DISPLAY', fold(`DESCRIPTION:${escI(`SFG: ${tipo} ${brl(Math.abs(tot))} amanhã`)}`), 'TRIGGER:-PT15H', 'END:VALARM', 'END:VEVENT'].join('\r\n'));
  }
  const ics = ['BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//SFG Incorporadora//Financeiro//PT-BR', 'CALSCALE:GREGORIAN', 'X-WR-CALNAME:SFG vencimentos', ...ev, 'END:VCALENDAR'].join('\r\n');
  const blob = new Blob([ics], { type: 'text/calendar;charset=utf-8' });
  if (/iPhone|iPad|Macintosh/i.test(navigator.userAgent) && 'ontouchend' in document) { location.href = URL.createObjectURL(blob); }
  else baixarArquivo(blob, `SFG vencimentos ${h}.ics`);
  toast(`${ev.length} lembretes (próximos 90 dias). Abra o arquivo para adicionar à agenda.`);
};

/* =========================================================
   Relatório mensal (PDF)
   ========================================================= */
let relMes = '';
function dadosRelatorio(m) {
  const ini = m + '-01', fim = addDays(addMonths(ini, 1), -1), h = hoje();
  const pagos = S.list.filter(pago); const base = sum(E().contas, c => +c.inicial || 0);
  const ate = d => base + sum(pagos.filter(t => (t.pag || t.venc) <= d), t => t.valor + (t.juros || 0));
  const mes = pagos.filter(t => { const d = t.pag || t.venc; return d >= ini && d <= fim; });
  const porDest = {}; for (const t of mes) if (t.valor < 0 || t.juros) { if (t.valor < 0) porDest[t.destino] = (porDest[t.destino] || 0) + t.valor; if (t.juros) porDest['Juros e Multa'] = (porDest['Juros e Multa'] || 0) + t.juros; }
  const ab = S.list.filter(t => !pago(t)); const pj = projecao();
  const fluxo = [...Array(6)].map((_, i) => { const mi = addMonths(h.slice(0, 7) + '-01', i), mf = addDays(addMonths(mi, 1), -1); let v = pj.base; for (const x of pj.dias) if (x.d <= mf) v = x.v; return [mesNome(mi.slice(0, 7)), v]; });
  return {
    m, ini, fim, saldoIni: ate(addDays(ini, -1)), saldoFim: ate(fim),
    entradas: sum(mes.filter(t => t.valor > 0 && t.destino !== 'Aporte'), t => t.valor), aportes: sum(mes.filter(t => t.destino === 'Aporte'), t => t.valor),
    saidas: sum(mes, t => Math.min(0, t.valor) + Math.min(0, t.juros || 0)),
    porDest: Object.entries(porDest).sort((a, b) => a[1] - b[1]),
    quadras: E().quadras.map(resumoQuadra),
    vencidas: agrupar(ab.filter(t => t.valor < 0 && t.venc < h).sort((a, b) => a.venc.localeCompare(b.venc)), 'venc'),
    prox: agrupar(ab.filter(t => t.venc >= h && t.venc <= addDays(h, 30)).sort((a, b) => a.venc.localeCompare(b.venc)), 'venc'),
    fluxo, min: pj.min
  };
}
const entTxt = e => e.t ? [dBR(e.t.venc), e.t.item, e.t.quadra, e.t.valor] : [dBR(e.itens[0].venc), baseItem(e.itens[0]) + ` (${e.itens.length} lotes)`, e.itens[0].quadra, sum(e.itens, x => x.valor)];
VIEWS.relatorios = {
  titulo: 'Relatório mensal',
  refresh(el) {
    const h = hoje(); if (!relMes) relMes = h.slice(0, 7);
    const ms = [...Array(18)].map((_, i) => addMonths(h.slice(0, 7) + '-01', -i).slice(0, 7));
    const r = dadosRelatorio(relMes);
    el.innerHTML = `<div class="filtros card"><label class="tgl">Mês <select id="rM">${opts(ms.map(m => [m, mesNome(m)]), relMes)}</select></label><button class="btn primary" id="rPdf">Baixar PDF</button><span class="muted small">O PDF traz o resumo do mês, despesas por destino, resultado por quadra, contas vencidas, próximos 30 dias e o fluxo projetado.</span></div>
    <div class="grid kpis k4">${kpi('Saldo inicial', brl(r.saldoIni), dBR(addDays(r.ini, -1)), cls(r.saldoIni))}${kpi('Entradas', brl(r.entradas), r.aportes ? 'aportes ' + brl(r.aportes) : '', 'pos')}${kpi('Saídas', brl(r.saidas), 'inclui juros', 'neg')}${kpi('Saldo final', brl(r.saldoFim), dBR(r.fim), cls(r.saldoFim))}</div>
    <div class="grid two"><div class="card"><h2>Saídas por destino</h2><table class="tbl"><tbody>${r.porDest.map(([d, v]) => `<tr><td class="l">${esc(d)}</td><td class="neg">${brl(v)}</td></tr>`).join('') || '<tr><td class="l muted">Sem saídas no mês</td></tr>'}</tbody></table></div>
    <div class="card"><h2>Fluxo projetado</h2><table class="tbl"><tbody>${r.fluxo.map(([n, v]) => `<tr><td class="l">${esc(n)}</td><td class="${cls(v)}">${brl(v)}</td></tr>`).join('')}</tbody></table></div></div>`;
    $('#rM', el).onchange = e => { relMes = e.target.value; this.refresh(el); };
    $('#rPdf', el).onclick = () => gerarPDF(relMes).catch(e => { console.error(e); toast('Erro ao gerar PDF: ' + e.message, 'err'); });
  }
};
async function imgParaPNG(url, w) {
  const img = await new Promise((res, rej) => { const i = new Image(); i.onload = () => res(i); i.onerror = rej; i.src = url; });
  const r = (img.naturalHeight || 508) / (img.naturalWidth || 908); const c = document.createElement('canvas'); c.width = w; c.height = Math.round(w * r);
  c.getContext('2d').drawImage(img, 0, 0, c.width, c.height); return { data: c.toDataURL('image/png'), r };
}
async function gerarPDF(m) {
  toast('Gerando PDF…');
  const r = dadosRelatorio(m);
  const P = s => String(s).replace(/\u2212/g, '-').replace(/[\u202f\u00a0]/g, ' ');
  const pdf = new jsPDF({ unit: 'pt', format: 'a4' }); const W = pdf.internal.pageSize.getWidth();
  const NAVY = [11, 32, 70];
  pdf.setFillColor(...NAVY); pdf.rect(0, 0, W, 86, 'F');
  try { const lg = await imgParaPNG(LOGO_BRANCO, 360); pdf.addImage(lg.data, 'PNG', 36, 20, 82, 82 * lg.r); } catch (e) { /* sem logo */ }
  pdf.setTextColor(255, 255, 255); pdf.setFont('helvetica', 'bold'); pdf.setFontSize(17); pdf.text('Relatório financeiro', W - 36, 40, { align: 'right' });
  pdf.setFont('helvetica', 'normal'); pdf.setFontSize(11); pdf.text(P(`${mesNome(m)} · ${E().local}`), W - 36, 58, { align: 'right' });
  pdf.setFontSize(8.5); pdf.text(P('SFG Incorporadora · gerado em ' + new Date().toLocaleString('pt-BR', { dateStyle: 'short', timeStyle: 'short' })), W - 36, 74, { align: 'right' });
  let y = 110; pdf.setTextColor(15, 27, 45);
  const titulo = t => { if (y > 740) { pdf.addPage(); y = 50; } pdf.setFont('helvetica', 'bold'); pdf.setFontSize(12); pdf.setTextColor(...NAVY); pdf.text(t, 36, y); y += 8; };
  const tabela = (head, body, colsNum) => { autoTable(pdf, { startY: y, head: [head], body: body.map(r => r.map(P)), margin: { left: 36, right: 36 }, styles: { fontSize: 8.5, cellPadding: 4, textColor: [15, 27, 45] }, headStyles: { fillColor: NAVY, textColor: 255, fontStyle: 'bold' }, alternateRowStyles: { fillColor: [243, 245, 249] }, columnStyles: Object.fromEntries(colsNum.map(i => [i, { halign: 'right' }])), didParseCell: dd => { if (colsNum.includes(dd.column.index)) dd.cell.styles.halign = 'right'; } }); y = pdf.lastAutoTable.finalY + 22; };
  titulo('Resumo do mês');
  tabela(['Saldo inicial', 'Entradas', 'Aportes', 'Saídas', 'Saldo final'], [[brl(r.saldoIni), brl(r.entradas), brl(r.aportes), brl(r.saidas), brl(r.saldoFim)]], [0, 1, 2, 3, 4]);
  titulo('Saídas por destino');
  tabela(['Destino', 'Valor'], r.porDest.length ? r.porDest.map(([d, v]) => [d, brl2(v)]) : [['Sem saídas no mês', '']], [1]);
  titulo('Resultado por quadra');
  tabela(['Quadra', 'Lotes', 'Custo realizado', 'A pagar', 'Receita', 'A receber', 'Projetado', 'Previsto'], r.quadras.map(q => [q.q, q.n, brl(q.custo), brl(q.pagar), brl(q.receita), brl(q.receber), brl(q.proj), q.prevCusto ? brl(q.prevRes) : '—']), [1, 2, 3, 4, 5, 6, 7]);
  titulo('Contas vencidas em aberto');
  tabela(['Vencimento', 'Descrição', 'Quadra', 'Valor'], r.vencidas.length ? r.vencidas.slice(0, 40).map(e => { const x = entTxt(e); return [x[0], x[1], x[2], brl2(x[3])]; }) : [['—', 'Nenhuma', '', '']], [3]);
  titulo('Próximos 30 dias');
  tabela(['Vencimento', 'Descrição', 'Quadra', 'Valor'], r.prox.length ? r.prox.slice(0, 60).map(e => { const x = entTxt(e); return [x[0], x[1], x[2], brl2(x[3])]; }) : [['—', 'Nada previsto', '', '']], [3]);
  titulo('Fluxo de caixa projetado (saldo no fim do mês)');
  tabela(['Mês', 'Saldo projetado'], r.fluxo.map(([n, v]) => [n, brl(v)]), [1]);
  pdf.setFontSize(9); pdf.setTextColor(91, 107, 130); pdf.text(P(`Menor saldo projetado: ${brl(r.min.v)} em ${dBR(r.min.d)}.`), 36, y);
  const np = pdf.getNumberOfPages(); for (let i = 1; i <= np; i++) { pdf.setPage(i); pdf.setFontSize(8); pdf.setTextColor(140); pdf.text(`${i} / ${np}`, W - 36, pdf.internal.pageSize.getHeight() - 20, { align: 'right' }); }
  await baixarArquivo(pdf.output('blob'), `SFG Relatório ${m}.pdf`);
}

/* =========================================================
   Usuários e acesso
   ========================================================= */
const REGRAS = `rules_version = '2';
service cloud.firestore {
  match /databases/{database}/documents {
    function logado() { return request.auth != null && request.auth.token.email != null; }
    function dono() { return logado() && request.auth.token.email in ['${DONOS.join("', '")}']; }
    function temAcesso() { return exists(/databases/$(database)/documents/config/acesso); }
    function acesso() { return get(/databases/$(database)/documents/config/acesso).data; }
    function editor() { return dono() || (logado() && temAcesso() && request.auth.token.email in acesso().get('editores', [])); }
    function leitor() { return editor() || (logado() && temAcesso() && request.auth.token.email in acesso().get('leitores', [])); }
    match /config/acesso { allow read: if leitor(); allow write: if dono(); }
    match /{col}/{id} {
      allow read: if leitor();
      allow write: if editor() && !(col == 'config' && id == 'acesso');
    }
  }
}`;
VIEWS.usuarios = {
  titulo: 'Usuários e acesso',
  refresh(el) {
    const a = S.acesso || {}; const souDono = papel() === 'dono';
    const linha = (em, p, pode) => `<tr><td class="l"><b>${esc(em)}</b></td><td class="l"><span class="chip ${p === 'Dono' ? 'ok' : p === 'Editor' ? '' : 'warn'}">${p}</span></td><td>${pode && souDono ? `<button class="btn sm link" data-act="usrDel" data-e="${esc(em)}">Remover</button>` : ''}</td></tr>`;
    el.innerHTML = `<div class="grid two"><div class="card"><h2>Quem tem acesso</h2><table class="tbl"><tbody>${DONOS.map(e => linha(e, 'Dono', false)).join('')}${(a.editores || []).map(e => linha(e, 'Editor', true)).join('')}${(a.leitores || []).map(e => linha(e, 'Somente leitura', true)).join('')}</tbody></table>
      ${souDono ? `<form class="form" id="fU" style="margin-top:14px"><div class="g2"><label>E-mail<input name="email" type="email" required></label><label>Permissão<select name="p"><option value="leitores">Somente leitura (sócio, contador)</option><option value="editores">Editor (lança e dá baixa)</option></select></label></div><div class="acts left"><button class="btn primary">Dar acesso</button></div></form>` : '<p class="muted small">Só o dono pode alterar os acessos.</p>'}</div>
    <div class="card"><h2>Como liberar um novo usuário</h2><ol class="small" style="padding-left:18px;line-height:1.7"><li>No <b>Firebase → Authentication → Usuários</b>, clique em <b>Adicionar usuário</b> com o e-mail e uma senha para a pessoa.</li><li><b>Uma única vez:</b> em <b>Firestore → Regras</b>, troque o texto pelas regras abaixo e clique em <b>Publicar</b>. Elas leem esta lista de acessos.</li><li>Inclua o e-mail aqui ao lado, com a permissão certa.</li><li>Mande o endereço do sistema e a senha para a pessoa. Quem tem "somente leitura" vê tudo, mas não consegue lançar, editar nem excluir.</li></ol>
      <pre class="regras">${esc(REGRAS)}</pre><button class="btn" id="cpR">Copiar regras</button></div></div>`;
    const fu = $('#fU', el); if (fu) fu.onsubmit = e => {
      e.preventDefault(); const em = fu.email.value.trim().toLowerCase(); if (DONOS.includes(em)) return;
      const n = { editores: (a.editores || []).filter(x => x !== em), leitores: (a.leitores || []).filter(x => x !== em) }; n[fu.p.value].push(em);
      S.acesso = n; setDoc(doc(db, 'config', 'acesso'), n).catch(onErroEscrita); toast('Acesso liberado para ' + em); refreshView();
    };
    $('#cpR', el).onclick = async () => { try { await navigator.clipboard.writeText(REGRAS); toast('Regras copiadas'); } catch (e) { toast('Selecione o texto e copie manualmente.', 'err'); } };
  }
};
ACOES.usrDel = async d => { const a = S.acesso || {}; if (!(await confirmar(`Remover o acesso de <b>${esc(d.e)}</b>?`, 'Remover'))) return; const n = { editores: (a.editores || []).filter(x => x !== d.e), leitores: (a.leitores || []).filter(x => x !== d.e) }; S.acesso = n; setDoc(doc(db, 'config', 'acesso'), n).catch(onErroEscrita); refreshView(); };

/* ---------- Mais (celular) ---------- */
VIEWS.mais = {
  titulo: 'Mais',
  refresh(el) { el.innerHTML = NAV.filter(([k]) => !['inicio', 'pagar', 'receber', 'obras'].includes(k)).map(([k, n]) => k === '#' ? `<div class="eyebrow" style="margin:14px 4px 6px">${n}</div>` : `<a class="menu-item" href="#/${k}">${ic(k)}<span>${n}</span></a>`).join(''); }
};

/* ---------- Dados ---------- */
VIEWS.dados = {
  titulo: 'Dados e backup',
  refresh(el) {
    el.innerHTML = `<div class="grid two">
      <div class="card"><h2>Exportar</h2><p class="muted">Gera uma planilha com Extrato, Contas a Pagar, Contas a Receber, Previsto e Checklist. Faça isso de tempos em tempos como backup.</p><div class="acts left"><button class="btn primary" id="bXls">Exportar Excel</button><button class="btn" id="bJson">Backup completo (.json)</button></div></div>
      <div class="card"><h2>Importar</h2><p class="muted">Importa a planilha <b>SFG Projeto</b> (abas Extrato, Contas a Pagar, Contas a Receber, Previsto e Checklist) ou um backup .json. <b>Substitui todos os dados atuais.</b> Precisa de internet.</p>
        <div class="acts left"><label class="btn file">Importar planilha .xlsx<input type="file" id="iXls" accept=".xlsx,.xlsm" hidden></label><label class="btn file">Restaurar backup .json<input type="file" id="iJson" accept=".json,application/json" hidden></label></div><p id="impStatus" class="muted"></p></div>
      <div class="card"><h2>Conta</h2><div class="kv"><span>Usuário</span><b>${esc(S.user?.email || '')}</b></div><div class="kv"><span>Lançamentos</span><b>${S.list.length.toLocaleString('pt-BR')}</b></div><div class="kv"><span>Versão</span><b>${VERSAO}</b></div><div class="acts left"><button class="btn danger ghost" id="bSair">Sair deste aparelho</button></div></div>
    </div>`;
    $('#bXls', el).onclick = exportarExcel; $('#bJson', el).onclick = exportarJSON;
    $('#iXls', el).onchange = e => importarArquivo(e.target.files[0], 'xlsx');
    $('#iJson', el).onchange = e => importarArquivo(e.target.files[0], 'json');
    $('#bSair', el).onclick = async () => { if (await confirmar('Sair? Alterações que ainda não sincronizaram podem se perder se você sair sem internet.', 'Sair')) signOut(auth); };
  }
};

async function baixarArquivo(blob, nome) {
  try { const file = new File([blob], nome, { type: blob.type }); if (navigator.canShare && navigator.canShare({ files: [file] }) && /iPhone|iPad|Android/i.test(navigator.userAgent)) { await navigator.share({ files: [file], title: nome }); return; } } catch (e) { if (e.name === 'AbortError') return; }
  const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = nome; document.body.appendChild(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 2000);
}
function exportarExcel() {
  const cab = ['Item', 'R$', 'Juros e Multa', 'NF', 'Local', 'Quadra', 'Lote', 'Destino', 'Data', 'Vencimento', 'Conta', 'Fornecedor', 'Observação', 'Quantidade', 'Unidade'];
  const lin = t => [t.item, t.valor, t.juros || 0, !!t.nf, t.local || E().local, t.quadra, isNaN(+t.lote) || t.lote === '' ? t.lote : +t.lote, t.destino, pago(t) ? (t.pag || t.venc) : t.venc, t.venc, t.conta || '', t.forn || '', t.obs || '', t.qtd || '', t.unid || ''];
  const toDate = r => { for (const i of [8, 9]) if (r[i]) { const [y, m, d] = r[i].split('-').map(Number); r[i] = new Date(Date.UTC(y, m - 1, d)); } return r; };
  const sheet = rows => { const ws = XLSX.utils.aoa_to_sheet([cab, ...rows.map(lin).map(toDate)], { cellDates: true }); ws['!cols'] = [{ wch: 44 }, { wch: 14 }, { wch: 12 }, { wch: 5 }, { wch: 14 }, { wch: 7 }, { wch: 6 }, { wch: 16 }, { wch: 11 }, { wch: 11 }, { wch: 10 }, { wch: 20 }, { wch: 24 }]; return ws; };
  const wb = XLSX.utils.book_new();
  const ord = (a, b) => (dataRef(a) || '').localeCompare(dataRef(b) || '');
  XLSX.utils.book_append_sheet(wb, sheet(S.list.filter(pago).sort(ord)), 'Extrato');
  XLSX.utils.book_append_sheet(wb, sheet(S.list.filter(t => !pago(t) && t.valor <= 0).sort(ord)), 'Contas a Pagar');
  XLSX.utils.book_append_sheet(wb, sheet(S.list.filter(t => !pago(t) && t.valor > 0).sort(ord)), 'Contas a Receber');
  const qs = E().quadras, ds = [...(E().destinos.op || []), ...(E().destinos.nonop || []), ...(E().destinos.rec || [])];
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['Destino (por lote)', ...qs], ...ds.map(d => [d, ...qs.map(q => +(S.orcamento?.quadras?.[q]?.[d] || 0))])]), 'Previsto');
  const docs = (S.checklist && S.checklist.docs) || DOCS_PADRAO;
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['Local', 'Quadra', 'Lote', ...docs], ...qs.flatMap(q => lotesDe(q).map(l => { const s = S.checklist?.lotes?.[q + '|' + l] || ''; return [E().local, q, l, ...docs.map((_, i) => s[i] === '1')]; }))]), 'Checklist');
  const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  baixarArquivo(new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), `SFG Financeiro ${hoje()}.xlsx`);
}
function exportarJSON() {
  const data = { formato: 'sfg-financeiro', versao: 1, exportado: new Date().toISOString(), estrutura: S.estrutura, orcamento: S.orcamento, checklist: S.checklist, vendas: S.vendas, obras: S.obras, fornecedores: S.fornecedores, simulacoes: S.simulacoes, titulos: S.list.map(limpar) };
  baixarArquivo(new Blob([JSON.stringify(data)], { type: 'application/json' }), `SFG Financeiro backup ${hoje()}.json`);
}

/* Leitura da planilha SFG Projeto */
function lerPlanilha(buf) {
  const wb = XLSX.read(buf, { type: 'array' });
  const need = ['Extrato', 'Contas a Pagar', 'Contas a Receber'];
  const miss = need.filter(n => !wb.Sheets[n]); if (miss.length) throw new Error('Abas não encontradas: ' + miss.join(', '));
  const rows = name => { const ws = wb.Sheets[name]; if (!ws || !ws['!ref']) return []; const r = XLSX.utils.decode_range(ws['!ref']); r.s.r = 0; r.s.c = 0; return XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: true, range: r }); };
  const s = x => x === null || x === undefined ? '' : String(x).trim();
  const isNum = x => typeof x === 'number' && isFinite(x);
  const dt = x => isNum(x) ? new Date(Math.round((x - 25569) * 86400000)).toISOString().slice(0, 10) : (x instanceof Date ? isoLocal(x) : '');
  let n = 0; const base = Date.now();
  if (s((rows('Extrato')[0] || [])[0]) === 'Item') return lerExportacao(rows, s, isNum, dt, base);
  const titulos = [];
  const ledger = (name, status) => rows(name).slice(5).forEach(r => {
    if (!isNum(r[2])) return; const d = dt(r[9]);
    const t = { id: base.toString(36) + (n++).toString(36).padStart(5, '0'), item: s(r[1]), valor: r[2], juros: isNum(r[3]) ? r[3] : 0, nf: r[4] === true ? 1 : 0, local: s(r[5]), quadra: s(r[6]) || '-', lote: s(r[7]) || '-', destino: s(r[8]), venc: d, status, criado: base + n, grupo: '' };
    if (status === 'pago') { t.pag = d; t.conta = 'BTG'; }
    titulos.push(t);
  });
  ledger('Extrato', 'pago'); ledger('Contas a Pagar', 'aberto'); ledger('Contas a Receber', 'aberto');
  // previsto
  const pr = rows('Previsto'); const quadrasPrev = (pr[5] || []).slice(2).filter(v => v !== null && v !== '').map(s);
  const orc = {}; quadrasPrev.forEach(q => orc[q] = {});
  let sec = null;
  pr.slice(6).forEach(r => { const lab = s(r[1]); if (/Não Operacionais/i.test(lab)) { sec = 'nonop'; return; } if (/Operacionais/i.test(lab)) { sec = 'op'; return; } if (/Receitas/i.test(lab)) { sec = 'rec'; return; } if (sec && lab && lab !== 'Total' && lab !== 'Juros e Multa') quadrasPrev.forEach((q, i) => { const v = r[2 + i]; if (isNum(v) && v) orc[q][lab] = v; }); });
  // checklist
  const cr = rows('Checklist'); const docs = (cr[4] || []).slice(5).filter(v => v !== null && v !== '').map(s);
  const ck = {}; const ckLotes = {};
  cr.slice(5).forEach(r => { const q = s(r[2]), l = s(r[3]); if (!q || !l) return; ck[q + '|' + l] = docs.map((_, i) => r[5 + i] === true ? '1' : '0').join(''); (ckLotes[q] = ckLotes[q] || new Set()).add(l); });
  // estrutura
  const quadras = [...quadrasPrev]; const lotes = {};
  const addL = (q, l) => { if (!q || q === '-') return; if (!quadras.includes(q)) quadras.push(q); if (!l || l === '-') return; (lotes[q] = lotes[q] || new Set()).add(l); };
  titulos.forEach(t => addL(t.quadra, t.lote)); for (const q in ckLotes) ckLotes[q].forEach(l => addL(q, l));
  const destinos = clone(DEST_PADRAO); const conhecidos = new Set(Object.values(destinos).flat());
  titulos.forEach(t => { if (t.destino && !conhecidos.has(t.destino)) { destinos.nonop.push(t.destino); conhecidos.add(t.destino); } });
  const local = titulos.find(t => t.local)?.local || 'Nova Juazeiro';
  const estrutura = { local, quadras, lotes: Object.fromEntries(Object.entries(lotes).map(([q, set]) => [q, [...set].sort(lotSort)])), destinos, contas: [{ nome: 'BTG', inicial: 0 }, { nome: 'Caixa', inicial: 0 }] };
  return { titulos, estrutura, orcamento: { quadras: orc }, checklist: { docs: docs.length ? docs : DOCS_PADRAO, lotes: ck } };
}
/* Planilha gerada pelo próprio sistema (botão Exportar Excel) */
function lerExportacao(rows, s, isNum, dt, base) {
  let n = 0; const titulos = [];
  const ledger = (name, status) => rows(name).slice(1).forEach(r => {
    if (!isNum(r[1])) return; const venc = dt(r[9]) || dt(r[8]);
    const t = { id: base.toString(36) + (n++).toString(36).padStart(5, '0'), item: s(r[0]), valor: r[1], juros: isNum(r[2]) ? r[2] : 0, nf: r[3] === true ? 1 : 0, local: s(r[4]), quadra: s(r[5]) || '-', lote: s(r[6]) || '-', destino: s(r[7]), venc, status, forn: s(r[11]), obs: s(r[12]), criado: base + n, grupo: '' };
    if (isNum(r[13]) && r[13] > 0) { t.qtd = r[13]; t.unid = s(r[14]); }
    if (status === 'pago') { t.pag = dt(r[8]) || venc; t.conta = s(r[10]) || 'BTG'; }
    titulos.push(t);
  });
  ledger('Extrato', 'pago'); ledger('Contas a Pagar', 'aberto'); ledger('Contas a Receber', 'aberto');
  const pr = rows('Previsto'); const qs = (pr[0] || []).slice(1).map(s).filter(Boolean); const orc = {}; qs.forEach(q => orc[q] = {});
  pr.slice(1).forEach(r => { const d = s(r[0]); if (d) qs.forEach((q, i) => { if (isNum(r[1 + i]) && r[1 + i]) orc[q][d] = r[1 + i]; }); });
  const cr = rows('Checklist'); const docs = (cr[0] || []).slice(3).map(s).filter(Boolean); const ck = {};
  const quadras = [...qs], lotes = {};
  const addL = (q, l) => { if (!q || q === '-') return; if (!quadras.includes(q)) quadras.push(q); if (!l || l === '-') return; (lotes[q] = lotes[q] || new Set()).add(l); };
  cr.slice(1).forEach(r => { const q = s(r[1]), l = s(r[2]); if (!q || !l) return; ck[q + '|' + l] = docs.map((_, i) => r[3 + i] === true ? '1' : '0').join(''); addL(q, l); });
  titulos.forEach(t => addL(t.quadra, t.lote));
  const destinos = clone(DEST_PADRAO); const conhecidos = new Set(Object.values(destinos).flat());
  titulos.forEach(t => { if (t.destino && !conhecidos.has(t.destino)) { destinos.nonop.push(t.destino); conhecidos.add(t.destino); } });
  const contas = [...new Set(['BTG', 'Caixa', ...titulos.filter(t => t.conta).map(t => t.conta)])].map(nome => ({ nome, inicial: 0 }));
  return { titulos, estrutura: { local: titulos.find(t => t.local)?.local || 'Nova Juazeiro', quadras, lotes: Object.fromEntries(Object.entries(lotes).map(([q, set]) => [q, [...set].sort(lotSort)])), destinos, contas }, orcamento: { quadras: orc }, checklist: { docs: docs.length ? docs : DOCS_PADRAO, lotes: ck } };
}
async function importarArquivo(file, tipo) {
  if (!file) return; const st = $('#impStatus');
  if (!navigator.onLine) { toast('A importação precisa de internet.', 'err'); return; }
  try {
    st.textContent = 'Lendo arquivo…';
    let data;
    if (tipo === 'xlsx') data = lerPlanilha(new Uint8Array(await file.arrayBuffer()));
    else { data = JSON.parse(await file.text()); if (data.formato !== 'sfg-financeiro') throw new Error('Arquivo não é um backup do SFG Financeiro.'); }
    const pagos = data.titulos.filter(pago).length;
    const ok = await confirmar(`Encontrei <b>${data.titulos.length.toLocaleString('pt-BR')}</b> lançamentos (${pagos.toLocaleString('pt-BR')} realizados, ${(data.titulos.length - pagos).toLocaleString('pt-BR')} em aberto)${data.estrutura ? `, ${data.estrutura.quadras.length} quadras` : ''}.<br><br><b>Isso substitui todos os dados atuais do sistema.</b> Continuar?`, 'Importar e substituir');
    if (!ok) { st.textContent = ''; return; }
    st.textContent = 'Gravando…';
    await substituirTudo(data);
    st.textContent = 'Importação concluída.'; toast('Importação concluída: ' + data.titulos.length.toLocaleString('pt-BR') + ' lançamentos');
  } catch (e) { console.error(e); st.textContent = ''; toast('Erro na importação: ' + e.message, 'err'); }
}
async function substituirTudo(data) {
  if (somenteLeitura()) throw new Error('Seu acesso é somente leitura.');
  const ops = [];
  const porMes = {};
  for (const t of data.titulos) { const m = bucketDe(t); (porMes[m] = porMes[m] || {})[t.id] = JSON.stringify(limpar(t)); }
  for (const m of Object.keys(S.buckets)) if (!porMes[m]) ops.push(b => b.delete(doc(db, 'titulos', m)));
  for (const m in porMes) ops.push(b => b.set(doc(db, 'titulos', m), { items: porMes[m] }));
  if (data.estrutura) ops.push(b => b.set(doc(db, 'config', 'estrutura'), data.estrutura));
  if (data.orcamento) ops.push(b => b.set(doc(db, 'config', 'orcamento'), data.orcamento));
  if (data.checklist) ops.push(b => b.set(doc(db, 'config', 'checklist'), data.checklist));
  for (const k of ['vendas', 'obras', 'fornecedores', 'simulacoes']) if (data[k]) ops.push(b => b.set(doc(db, 'config', k), data[k]));
  for (let i = 0; i < ops.length; i += 20) { const b = writeBatch(db); ops.slice(i, i + 20).forEach(f => f(b)); await b.commit(); }
}

/* =========================================================
   Início
   ========================================================= */
if ('serviceWorker' in navigator && !EMU) {
  navigator.serviceWorker.register('./sw.js').catch(e => console.warn('SW', e));
  let recarregou = false;
  navigator.serviceWorker.addEventListener('controllerchange', () => { if (recarregou) return; recarregou = true; if (!$('.sheet-wrap')) location.reload(); });
}
matchMedia('(prefers-color-scheme: dark)').addEventListener('change', () => refreshView());
if (!CONFIGURADO) telaNaoConfigurado();
else onAuthStateChanged(auth, u => {
  if (u) { S.user = u; iniciarDados(); montarShell(); }
  else { S.user = null; pararDados(); document.removeEventListener('click', onClick); atual = { nome: null, args: [], view: null }; telaLogin(); }
});
window.__sfg = { S, hoje };
