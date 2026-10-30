import { initializeApp } from 'firebase/app';
import { initializeAuth, indexedDBLocalPersistence, browserLocalPersistence, onAuthStateChanged, signInWithEmailAndPassword, signOut, sendPasswordResetEmail, connectAuthEmulator } from 'firebase/auth';
import { initializeFirestore, persistentLocalCache, persistentMultipleTabManager, connectFirestoreEmulator, collection, doc, onSnapshot, setDoc, getDoc, getDocs, deleteField, writeBatch } from 'firebase/firestore';
/* Bibliotecas pesadas (planilha, gráficos, PDF) carregam só quando usadas: o app abre mais rápido. */
let XLSX = null, Chart = null, jsPDF = null, autoTable = null;
const _libs = {};
function carregarLib(nome) {
  if (_libs[nome]) return _libs[nome];
  _libs[nome] = new Promise((res, rej) => {
    const s = document.createElement('script'); s.src = `lib-${nome}.js?v=__VERSAO__`;
    s.onload = () => { if (nome === 'xlsx') XLSX = window.__XLSX; if (nome === 'chart') Chart = window.__Chart; if (nome === 'pdf') { jsPDF = window.__jsPDF; autoTable = window.__autoTable; } res(); };
    s.onerror = () => { _libs[nome] = null; rej(new Error('Não consegui carregar uma parte do app. Verifique a internet e tente de novo.')); };
    document.head.appendChild(s);
  });
  return _libs[nome];
}
import { firebaseConfig } from './config.js';
import qrcode from 'qrcode-generator';

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

/* ---------- Tema claro / escuro ---------- */
const TEMAS = [['auto', 'Automático'], ['claro', 'Claro'], ['escuro', 'Escuro']];
const temaPref = () => { try { const m = localStorage.getItem('sfgTema'); return m === 'claro' || m === 'escuro' ? m : 'auto'; } catch (e) { return 'auto'; } };
const escuroAgora = () => { const m = temaPref(); return m === 'escuro' || (m === 'auto' && matchMedia('(prefers-color-scheme: dark)').matches); };
function aplicarTema(m) {
  try { if (m === 'auto') localStorage.removeItem('sfgTema'); else localStorage.setItem('sfgTema', m); } catch (e) { /* ok */ }
  if (m === 'auto') delete document.documentElement.dataset.theme; else document.documentElement.dataset.theme = m === 'claro' ? 'light' : 'dark';
  const b = document.getElementById('btnTema'); if (b) { b.innerHTML = ic(m === 'claro' ? 'sol' : m === 'escuro' ? 'lua' : 'auto'); b.title = b.ariaLabel = 'Tema: ' + TEMAS.find(x => x[0] === m)[1] + ' (toque para trocar)'; }
  const lg = document.querySelector('.login-logo'); if (lg) lg.src = escuroAgora() ? 'icons/logo-branco-incorporadora.svg' : LOGO_COMPLETO;
  if (typeof refreshView === 'function' && document.getElementById('view')) refreshView();
}
const seletorTema = () => `<div class="seg" role="group" aria-label="Tema">${TEMAS.map(([v, l]) => `<button class="${temaPref() === v ? 'on' : ''}" data-act="tema" data-v="${v}">${ic(v === 'claro' ? 'sol' : v === 'escuro' ? 'lua' : 'auto', 16)} ${l}</button>`).join('')}</div>`;

/* ---------- Dias úteis e feriados ---------- */
function pascoa(y) { const a = y % 19, b = Math.floor(y / 100), c = y % 100, d = Math.floor(b / 4), e = b % 4, f = Math.floor((b + 8) / 25), g = Math.floor((b - f + 1) / 3), h = (19 * a + b - d - g + 15) % 30, i = Math.floor(c / 4), k = c % 4, l = (32 + 2 * e + 2 * i - h - k) % 7, m = Math.floor((a + 11 * h + 22 * l) / 451), mes = Math.floor((h + l - 7 * m + 114) / 31), dia = ((h + l - 7 * m + 114) % 31) + 1; return isoLocal(new Date(y, mes - 1, dia)); }
const _feriados = {};
function feriadosNacionais(y) {
  if (_feriados[y]) return _feriados[y];
  const m = new Map(); const p = pascoa(y);
  [['01-01', 'Confraternização Universal'], ['04-21', 'Tiradentes'], ['05-01', 'Dia do Trabalho'], ['09-07', 'Independência'], ['10-12', 'Nossa Senhora Aparecida'], ['11-02', 'Finados'], ['11-15', 'Proclamação da República'], ['11-20', 'Consciência Negra'], ['12-25', 'Natal']].forEach(([md, n]) => m.set(`${y}-${md}`, n));
  m.set(addDays(p, -48), 'Carnaval'); m.set(addDays(p, -47), 'Carnaval'); m.set(addDays(p, -2), 'Sexta-feira Santa'); m.set(addDays(p, 60), 'Corpus Christi');
  return (_feriados[y] = m);
}
function feriadoNome(iso) {
  const n = feriadosNacionais(+iso.slice(0, 4)).get(iso); if (n) return n;
  for (const f of E().feriados) if (f.d.length === 5 ? iso.slice(5) === f.d : iso === f.d) return f.nome || 'Feriado local';
  return '';
}
const DSEM = ['dom', 'seg', 'ter', 'qua', 'qui', 'sex', 'sáb'];
const diaSemana = iso => new Date(iso + 'T12:00').getDay();
function motivoNaoUtil(iso) { if (!/^\d{4}-\d{2}-\d{2}$/.test(iso || '')) return ''; const w = diaSemana(iso); if (w === 6) return 'sábado'; if (w === 0) return 'domingo'; return feriadoNome(iso); }
function proximoDiaUtil(iso) { let d = iso, n = 0; while (motivoNaoUtil(d) && n++ < 20) d = addDays(d, 1); return d; }
const dCurtoSem = iso => DSEM[diaSemana(iso)] + ' ' + dBR(iso).slice(0, 5);
const prefUtil = () => { try { return localStorage.getItem('sfgUtil') !== '0'; } catch (e) { return true; } };
const setPrefUtil = v => { try { localStorage.setItem('sfgUtil', v ? '1' : '0'); } catch (e) { /* ok */ } };

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
    contas: e.contas && e.contas.length ? e.contas : [{ nome: 'BTG', inicial: 0 }, { nome: 'Caixa', inicial: 0 }],
    feriados: e.feriados || [],
    obras: e.obras && e.obras.length ? e.obras : [e.local || 'Nova Juazeiro'],
    quadraObra: e.quadraObra || {},
    empresas: e.empresas && e.empresas.length ? e.empresas : [{ id: 'matriz', nome: 'SFG Incorporadora', cnpj: '', tipo: 'matriz' }],
    obraEmpresa: e.obraEmpresa || {},
    obraInfo: e.obraInfo || {}
  };
};
const empPorId = id => E().empresas.find(x => x.id === id);
const matrizId = () => (E().empresas.find(x => x.tipo === 'matriz') || E().empresas[0]).id;
const empDaObra = o => { const id = E().obraEmpresa[o]; return id && empPorId(id) ? id : matrizId(); };
const empT = t => t.emp && empPorId(t.emp) ? t.emp : t.ambito === 'empresa' ? matrizId() : empDaObra(obraT(t));
const empConta = nome => { const c = E().contas.find(x => x.nome === nome); return c && c.emp && empPorId(c.emp) ? c.emp : matrizId(); };
const variasEmpresas = () => E().empresas.length > 1;
const empDef = t => !!(t.emp && empPorId(t.emp));
const empDiverge = t => empDef(t) && t.emp !== empDaObra(obraT(t));
const empPendente = t => variasEmpresas() && (!empDef(t) || empDiverge(t));
const fCnpj = c => { const d = String(c || '').replace(/\D/g, ''); return d.length === 14 ? `${d.slice(0, 2)}.${d.slice(2, 5)}.${d.slice(5, 8)}/${d.slice(8, 12)}-${d.slice(12)}` : (c || ''); };
function cnpjValido(c) {
  const d = String(c || '').replace(/\D/g, ''); if (d.length !== 14 || /^(\d)\1+$/.test(d)) return false;
  const dv = n => { let s = 0, p = n - 7; for (let i = 0; i < n; i++) { s += +d[i] * p--; if (p < 2) p = 9; } const r = s % 11; return r < 2 ? 0 : 11 - r; };
  return dv(12) === +d[12] && dv(13) === +d[13];
}
const empRotulo = (x, curto) => x ? x.nome + (x.cnpj && !curto ? ' · ' + fCnpj(x.cnpj) : '') + (x.tipo === 'matriz' && curto ? ' (matriz)' : '') : '';
const empOpts = () => E().empresas.map(x => [x.id, (x.tipo === 'matriz' ? 'Matriz · ' : 'Filial · ') + x.nome + (x.cnpj ? ' · ' + fCnpj(x.cnpj).slice(-8) : '')]);
const obraDe = q => { const e = E(); const o = e.quadraObra[q]; return o && e.obras.includes(o) ? o : e.obras[0]; };
const obraT = t => t.ambito === 'empresa' ? '' : t.quadra && t.quadra !== '-' && E().quadras.includes(t.quadra) ? obraDe(t.quadra) : (t.local || E().obras[0]);
const ehEmpresa = t => t.ambito === 'empresa';
const variasObras = () => E().obras.length > 1;
const quadrasDaObra = o => E().quadras.filter(q => !o || obraDe(q) === o);
function secaoDe(dest) { const d = E().destinos; for (const s of ['op', 'nonop', 'rec', 'aporte']) if ((d[s] || []).includes(dest)) return s; return ['Aporte', 'Transferência'].includes(dest) ? 'aporte' : 'nonop'; }
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
const CAMPOS_HIST = [['item', 'Descrição'], ['valor', 'Valor'], ['venc', 'Vencimento'], ['status', 'Situação'], ['pag', 'Pago em'], ['conta', 'Conta'], ['juros', 'Juros/multa'], ['destino', 'Destino'], ['local', 'Obra'], ['quadra', 'Quadra'], ['lote', 'Lote'], ['emp', 'CNPJ'], ['forn', 'Fornecedor'], ['nfNum', 'NF'], ['obs', 'Observação'], ['etapaObra', 'Etapa da obra'], ['contrato', 'Contrato'], ['ambito', 'Pertence a'], ['aporteTipo', 'Tipo de aporte']];
const HIST_NO_LANC = 5;
function registrarHist(t, ant, opt) {
  const em = Date.now(), por = S.user?.email || ''; const log = x => { (opt._log = opt._log || []).push({ ...x, t: t.id, item: t.item, valor: t.valor }); return x; };
  if (opt.acao) { t.hist = [...(t.hist || []), log({ em, por, a: opt.acao })].slice(-HIST_NO_LANC); return; }
  if (!ant) { t.hist = [log({ em, por, a: 'criou' })]; t.criadoPor = por; return; }
  const m = {};
  for (const [k] of CAMPOS_HIST) { const a = ant[k] ?? '', b = t[k] ?? ''; if (JSON.stringify(a) !== JSON.stringify(b)) m[k] = [a, b]; }
  const na = anexosDe(ant).length, nb = anexosDe(t).length; if (na !== nb) m.anexos = [na, nb];
  if (!Object.keys(m).length) return;
  const a = ant.status !== t.status ? (t.status === 'pago' ? 'baixa' : 'estorno') : 'alterou';
  t.hist = [...(ant.hist || []), log({ em, por, a, m })].slice(-HIST_NO_LANC); t.alteradoPor = por;
}
function gravarHist(log) {
  if (!log || !log.length) return; const items = {};
  for (const x of log) items[x.t + '_' + x.em + '_' + Math.random().toString(36).slice(2, 5)] = JSON.stringify(x);
  gravar(doc(db, 'hist', hoje().slice(0, 7)), { items }, { merge: true }).catch(onErroEscrita);
  if (S.histCompleto) for (const k in items) S.histCompleto.push(JSON.parse(items[k]));
}
async function carregarHistCompleto() {
  if (S.histCompleto) return S.histCompleto;
  const s = await getDocs(collection(db, 'hist')); const out = [];
  s.forEach(d => { const it = d.data().items || {}; for (const k in it) { try { out.push(JSON.parse(it[k])); } catch (e) { /* ignora */ } } });
  return (S.histCompleto = out);
}
function salvarTitulos(lista, opt = {}) {
  if (somenteLeitura()) { toast('Seu acesso é somente leitura.', 'err'); return; }
  const porMes = {}, apagar = {};
  for (const t of lista) {
    if (!opt.sistema) registrarHist(t, S.byId.get(t.id), opt);
    t.alterado = Date.now();
    const novo = bucketDe(t);
    const antigo = S.byId.get(t.id)?.mes;
    if (antigo && antigo !== novo) (apagar[antigo] = apagar[antigo] || {})[t.id] = deleteField();
    (porMes[novo] = porMes[novo] || {})[t.id] = JSON.stringify(limpar(t));
  }
  for (const m in apagar) gravar(doc(db, 'titulos', m), { items: apagar[m] }, { merge: true }).catch(onErroEscrita);
  for (const m in porMes) gravar(doc(db, 'titulos', m), { items: porMes[m] }, { merge: true }).catch(onErroEscrita);
  if (opt._log) { gravarHist(opt._log.length > 400 ? opt._log.filter(x => x.a !== 'criou') : opt._log); delete opt._log; }
}
function excluirTitulos(ids, opt = {}) {
  if (somenteLeitura()) { toast('Seu acesso é somente leitura.', 'err'); return; }
  if (!opt.semLixeira) {
    const em = Date.now(), por = S.user?.email || ''; const items = {};
    for (const id of ids) { const t = S.byId.get(id); if (t) items[id] = JSON.stringify({ t: limpar(t), por, em }); }
    if (Object.keys(items).length) gravar(doc(db, 'lixeira', isoLocal(new Date(em)).slice(0, 7)), { items }, { merge: true }).catch(onErroEscrita);
    const velho = isoLocal(new Date(em - 92 * 864e5)).slice(0, 7);
    for (const b of new Set(Object.values(S.lixeiraOnde || {}))) if (b < velho) gravar(doc(db, 'lixeira', b), { items: {} }).catch(onErroEscrita);
  }
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
function projecao(lista = S.list, base0) {
  const h = hoje(), base = base0 === undefined ? saldos().total : base0;
  const ev = lista.filter(t => !pago(t)).map(t => ({ d: t.venc && t.venc > h ? t.venc : h, v: t.valor, t }));
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
const normTxt = s => String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
const docsCk = () => (S.checklist && S.checklist.docs) || DOCS_PADRAO;
/* Arquivo mais recente de um documento para o lote: no próprio lote, na quadra ou na obra. */
function arquivoDoc(q, l, doc) {
  const A = (S.docsArq && S.docsArq.arq) || (S.docsdrive && S.docsdrive.arq) || {};
  const c = [(A['L|' + q + '|' + l] || {})[doc], (A['Q|' + q] || {})[doc], (A['O|' + obraDe(q)] || {})[doc]].filter(Boolean);
  return c.sort((a, b) => (b.d || '').localeCompare(a.d || ''))[0] || null;
}
function estadoDoc(q, l, i) {
  const docs = docsCk(), doc = docs[i]; const s = ((S.checklist && S.checklist.lotes) || {})[q + '|' + l] || '';
  const manual = s[i] === '1'; const arq = arquivoDoc(q, l, doc);
  const dias = +(((S.checklist && S.checklist.validade) || {})[doc] || 0);
  const venc = !!(arq && dias && arq.d && addDays(arq.d, dias) < hoje());
  return { doc, manual, arq, venc, dias, feito: (manual && !venc) || (!!arq && !venc), vale: arq && dias && arq.d ? addDays(arq.d, dias) : '' };
}
function ckPct(q, l) {
  const docs = docsCk(); if (!docs.length) return null;
  const lots = l === undefined ? lotesDe(q) : [l]; if (!lots.length) return null;
  let ok = 0; for (const x of lots) for (let i = 0; i < docs.length; i++) if (estadoDoc(q, x, i).feito) ok++;
  return ok / (lots.length * docs.length);
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
  /* Coleções em blocos (um documento por mês ou por quadra): sem o limite de 1 MB de um documento só. */
  const blocos = (col, aoMudar) => S.unsubs.push(onSnapshot(collection(db, col), s => { const o = {}; s.forEach(d => { o[d.id] = d.data(); }); S['raw_' + col] = o; aoMudar(o); agendarRefresh(); }, e => console.warn(e)));
  for (const [col, fn] of BLOCOS_EXTRA) blocos(col, fn);
  blocos('lixeira', o => { const items = {}, onde = {}; for (const b in o) for (const k in (o[b].items || {})) { items[k] = o[b].items[k]; onde[k] = b; } S.lixeira = { items }; S.lixeiraOnde = onde; S.loaded.lixeira = true; });
  blocos('fluxoLotes', o => { const m = {}; for (const b in o) { const q = o[b].q; for (const l in (o[b].lotes || {})) m[q + '|' + l] = o[b].lotes[l]; } S.fluxoLotes = m; S.loaded.fluxoLotes = true; });
  blocos('docsArq', o => { const arq = {}, outros = {}; for (const b in o) { Object.assign(arq, o[b].arq || {}); Object.assign(outros, o[b].outros || {}); } S.docsArq = { arq, outros, blocos: Object.keys(o) }; S.loaded.docsArq = true; });
  const cfg = (id, key) => S.unsubs.push(onSnapshot(doc(db, 'config', id), snap => {
    S[key] = snap.exists() ? snap.data() : null; if (key === 'estrutura') S.loaded.estrutura = true; agendarRefresh();
  }, e => console.error(e)));
  cfg('estrutura', 'estrutura'); cfg('orcamento', 'orcamento'); cfg('checklist', 'checklist');
  cfg('drive', 'drive'); for (const k of CONFIG_EXTRA) cfg(k, k); cfg('fluxo', 'fluxo'); cfg('docsdrive', 'docsdrive'); cfg('agenda', 'agenda'); cfg('lixeira', 'lixeiraAntiga'); cfg('recorrencias', 'recorrencias'); cfg('vendas', 'vendas'); cfg('obras', 'obras'); cfg('fornecedores', 'fornecedores'); cfg('simulacoes', 'simulacoes');
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
let _migrou = false;
function migrarBlocos() {
  if (_migrou || somenteLeitura() || !S.user || !S.loaded.lixeira || !S.loaded.fluxoLotes || !S.loaded.docsArq || S.lixeiraAntiga === undefined || S.fluxo === undefined || S.docsdrive === undefined) return;
  _migrou = true;
  const lx = S.lixeiraAntiga && S.lixeiraAntiga.items;
  if (lx && Object.keys(lx).length) { const por = {}; for (const k in lx) { let em = Date.now(); try { em = JSON.parse(lx[k]).em || em; } catch (e) { /* ok */ } (por[isoLocal(new Date(em)).slice(0, 7)] = por[isoLocal(new Date(em)).slice(0, 7)] || {})[k] = lx[k]; } for (const b in por) gravar(doc(db, 'lixeira', b), { items: por[b] }, { merge: true }).catch(onErroEscrita); gravar(doc(db, 'config', 'lixeira'), { items: {} }).catch(onErroEscrita); }
  const fl = S.fluxo && S.fluxo.lotes;
  if (fl && Object.keys(fl).length) { const por = {}; for (const k in fl) { const [q, l] = k.split('|'); (por[q] = por[q] || {})[l] = fl[k]; } for (const q in por) gravar(doc(db, 'fluxoLotes', 'Q_' + idBloco(q)), { q, lotes: por[q] }, { merge: true }).catch(onErroEscrita); gravar(doc(db, 'config', 'fluxo'), { lotes: deleteField() }, { merge: true }).catch(onErroEscrita); }
  const dd = S.docsdrive;
  if (dd && (dd.arq || dd.outros)) { salvarArqDocs(dd.arq || {}, dd.outros || {}, true); salvarDocsDrive(dd); }
}
let _recT = 0;
function agendarRefresh() { agendarAgenda(); migrarBlocos(); if (typeof agendarPortais === 'function') agendarPortais(); clearTimeout(_recT); _recT = setTimeout(() => gerarRecorrencias(), 1200); if (_raf) return; _raf = requestAnimationFrame(() => { _raf = 0; refreshView(); }); }

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
  socios: '<path d="M21.2 15.9A10 10 0 1 1 8 2.8M22 12A10 10 0 0 0 12 2v10z"/>',
  empresas: '<path d="M3 21h18M5 21V5l7-3v19M12 9h7v12M8 8h1M8 12h1M8 16h1M15 13h1M15 17h1"/>',
  sol: '<circle cx="12" cy="12" r="4"/><path d="M12 2v2M12 20v2M4.9 4.9l1.4 1.4M17.7 17.7l1.4 1.4M2 12h2M20 12h2M4.9 19.1l1.4-1.4M17.7 6.3l1.4-1.4"/>',
  lua: '<path d="M21 12.8A9 9 0 1 1 11.2 3a7 7 0 0 0 9.8 9.8z"/>',
  auto: '<circle cx="12" cy="12" r="9"/><path d="M12 3v18A9 9 0 0 0 12 3z" fill="currentColor"/>',
  fornecedores: '<path d="M3 7h13v10H3zM16 10h3l2 3v4h-5M7 20a2 2 0 1 0 0-4 2 2 0 0 0 0 4zM18 20a2 2 0 1 0 0-4 2 2 0 0 0 0 4z"/>',
  usuarios: '<path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2M12 11a4 4 0 1 0 0-8 4 4 0 0 0 0 8z"/>',
  calendario: '<path d="M4 5h16v16H4zM4 9h16M8 3v4M16 3v4"/>',
  x: '<path d="M6 6l12 12M18 6L6 18"/>',
  voltar: '<path d="M15 6l-6 6 6 6"/>',
  historico: '<path d="M3 12a9 9 0 1 0 3-6.7L3 8M3 3v5h5M12 7v5l3 3"/>',
  recorrentes: '<path d="M17 2l4 4-4 4M3 11V9a3 3 0 0 1 3-3h15M7 22l-4-4 4-4M21 13v2a3 3 0 0 1-3 3H3"/>',
  demonstrativos: '<path d="M4 4h16v16H4zM4 9h16M9 9v11M14 13h3M14 16h3"/>',
  scan: '<path d="M3 7V5a2 2 0 0 1 2-2h2M17 3h2a2 2 0 0 1 2 2v2M21 17v2a2 2 0 0 1-2 2h-2M7 21H5a2 2 0 0 1-2-2v-2M7 8v8M10 8v8M13 8v8M16 8v8"/>',
  movimento: '<path d="M3 3v18h18M7 15l3-3 3 3 5-6"/><path d="M7 19v-2M11 19v-4M15 19v-3M19 19v-6"/>',
  etapas: '<rect x="3" y="4" width="6" height="5" rx="1"/><rect x="15" y="4" width="6" height="5" rx="1"/><rect x="9" y="15" width="6" height="5" rx="1"/><path d="M9 6.5h6M18 9v3H12v3"/>',
  busca: '<circle cx="11" cy="11" r="7"/><path d="M20 20l-4-4"/>',
  area_relatorios: '<path d="M14 3H6a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V9zM14 3v6h6M8 17v-3M12 17v-6M16 17v-4"/>',
  area_suprimentos: '<path d="M4 7h16l-1.5 11a2 2 0 0 1-2 2h-9a2 2 0 0 1-2-2zM9 7V5a3 3 0 0 1 6 0v2"/>',
  area_planejamento: '<path d="M3 4h18v16H3zM3 9h18M8 4v16M12 13h6M12 16h4"/>',
  area_documentos: '<path d="M9 11l2 2 4-4M5 4h14v16H5z"/>',
  area_sistema: '<circle cx="9" cy="8" r="3.5"/><path d="M3 20a6 6 0 0 1 12 0M17 8h4M17 12h4M18 16h3"/>',
  arrumar: '<path d="M3 6h13M3 12h9M3 18h6M17 14l2 2 4-4"/>',
  clip: '<path d="M21 11l-8.5 8.5a5 5 0 0 1-7-7L14 4a3.5 3.5 0 0 1 5 5l-8.5 8.5a2 2 0 0 1-3-3L15 7"/>'
};
const ic = (n, s = 20) => `<svg width="${s}" height="${s}" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${IC[n]}</svg>`;

/* Menu: telas parecidas ficam juntas numa área com abas. */
const AREAS = {
  relatorios: ['Relatórios', [['demonstrativos', 'DRE, DFC e BP'], ['relatorios', 'Relatório mensal'], ['movimento', 'Histórico financeiro'], ['ret', 'Impostos (RET)']]],
  suprimentos: ['Contratos e compras', [['contratos', 'Contratos'], ['compras', 'Compras e cotações']]],
  planejamento: ['Planejamento', [['orcamento', 'Orçamento por etapa'], ['cronograma', 'Cronograma'], ['simulador', 'Simulador']]],
  documentos: ['Documentos e etapas', [['checklist', 'Checklist'], ['etapas', 'Fluxo do lote'], ['diario', 'Diário de obra']]],
  sistema: ['Usuários e dados', [['usuarios', 'Usuários'], ['historico', 'Histórico e lixeira'], ['dados', 'Dados e backup']]]
};
const areaDe = nome => Object.keys(AREAS).find(a => AREAS[a][1].some(x => x[0] === nome));
const NAV = [
  ['#', 'Financeiro'], ['inicio', 'Início'], ['pagar', 'Contas a pagar'], ['receber', 'Contas a receber'], ['extrato', 'Extrato'], ['fluxo', 'Fluxo de caixa'], ['recorrentes', 'Contas recorrentes'], ['conciliacao', 'Conciliação'], ['@relatorios'],
  ['#', 'Obras'], ['obras', 'Quadras e lotes'], ['vendas', 'Vendas e repasses'], ['@suprimentos'], ['@planejamento'], ['@documentos'],
  ['#', 'Gestão'], ['fornecedores', 'Fornecedores'], ['empresas', 'Empresas (CNPJ)'], ['socios', 'Sócios e participações'], ['arrumar', 'Revisar e classificar'], ['cadastros', 'Cadastros'], ['@sistema']
].map(([k, n]) => k[0] === '@' ? [AREAS[k.slice(1)][1][0][0], AREAS[k.slice(1)][0], k.slice(1)] : [k, n]);
const LOGO_BRANCO = 'icons/logo-branco.svg', LOGO_AZUL = 'icons/logo-azul.svg', LOGO_COMPLETO = 'icons/logo-azul-incorporadora.svg';
const NAV_MOBILE = [['inicio', 'Início'], ['pagar', 'Pagar'], ['receber', 'Receber'], ['obras', 'Obras'], ['mais', 'Mais']];

/* =========================================================
   Telas de entrada
   ========================================================= */
function telaNaoConfigurado() {
  document.body.innerHTML = `<div class="center"><div class="login card"><img src="${escuroAgora() ? 'icons/logo-branco-incorporadora.svg' : LOGO_COMPLETO}" alt="SFG Incorporadora" class="login-logo"><h1>Sistema ainda não configurado</h1><p class="muted">Falta colar a configuração do Firebase em <code>src/config.js</code> e gerar o app novamente.</p></div></div>`;
}
function telaLogin(msg) {
  document.body.innerHTML = `<div class="center"><form class="login card" id="fLogin" autocomplete="on">
    <img src="${escuroAgora() ? 'icons/logo-branco-incorporadora.svg' : LOGO_COMPLETO}" alt="SFG Incorporadora" class="login-logo"><h1>Financeiro</h1><p class="muted">Acesso restrito</p>
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
    <aside class="side"><div class="brand"><img class="logo-img" src="${LOGO_BRANCO}" alt=""><div><b>SFG</b><small>INCORPORADORA</small></div><button class="side-tgl" id="sideTgl" title="Recolher menu" aria-label="Recolher ou expandir o menu"><svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M15 18l-6-6 6-6"/></svg><span>Recolher menu</span></button></div>
      <nav>${NAV.map(([k, n, ar]) => k === '#' ? `<div class="sec">${n}</div>` : `<a href="#/${k}" data-nav="${k}" title="${esc(n)}" ${ar ? `data-area="${ar}"` : ''}>${ic(ar ? 'area_' + ar : k, 18)}<span>${n}</span></a>`).join('')}</nav>
      <div class="side-foot"><span id="sideUser" class="muted"></span></div></aside>
    <div class="main">
      <header class="top"><button class="icon-btn back hidden" id="btnBack" aria-label="Voltar">${ic('voltar')}</button><img class="tlogo" src="${LOGO_BRANCO}" alt="SFG"><h1 id="title"></h1><span id="sync" class="sync"></span><button class="icon-btn busca-btn" data-act="busca" aria-label="Buscar ou perguntar" title="Buscar ou perguntar (tecla /)">${ic('busca')}</button><button class="icon-btn tema-btn" id="btnTema" data-act="temaCiclo" aria-label="Tema">${ic(temaPref() === 'claro' ? 'sol' : temaPref() === 'escuro' ? 'lua' : 'auto')}</button><button class="btn primary desk rw" data-act="novo">${ic('plus', 18)} Novo lançamento</button></header>
      <nav id="areaTabs" class="area-tabs hidden" aria-label="Seções"></nav>
      <div id="view" class="view"></div>
    </div>
    <nav class="bottom">${NAV_MOBILE.map(([k, n]) => `<a href="#/${k}" data-nav="${k}">${ic(k, 22)}<span>${n}</span></a>`).join('')}</nav>
    <button class="fab rw" data-act="novo" aria-label="Novo lançamento">${ic('plus', 26)}</button>
  </div><div id="sheets"></div><div id="toast"></div>`;
  $('#sideUser').textContent = S.user?.email || ''; document.body.classList.toggle('ro', somenteLeitura());
  $('#btnBack').onclick = () => history.back();
  { const ap = $('.app'); let rec = false; try { rec = localStorage.getItem('sfgMenu') === 'recolhido'; } catch (e) { /* sem storage */ }
    const aplica = () => { ap.classList.toggle('side-rec', rec); const b = $('#sideTgl'); b.title = rec ? 'Expandir menu' : 'Recolher menu'; b.querySelector('span').textContent = rec ? 'Expandir menu' : 'Recolher menu'; };
    $('#sideTgl').onclick = () => { rec = !rec; try { localStorage.setItem('sfgMenu', rec ? 'recolhido' : 'aberto'); } catch (e) { /* sem storage */ } aplica(); window.dispatchEvent(new Event('resize')); }; aplica(); }
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
  const area = areaDe(nome);
  $$('[data-nav]').forEach(a => {
    const k = a.dataset.nav;
    const on = (a.dataset.area ? a.dataset.area === area : k === nome) || (k === 'mais' && !['inicio', 'pagar', 'receber', 'obras'].includes(nome));
    a.classList.toggle('on', on);
  });
  const at = $('#areaTabs'); at.classList.toggle('hidden', !area);
  at.innerHTML = area ? AREAS[area][1].map(([k, n]) => `<a href="#/${k}" class="${k === nome ? 'on' : ''}">${n}</a>`).join('') : '';
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
const quadrasOpts = () => [['-', 'Geral / administrativo (-)'], ...E().quadras.map(q => [q, 'Quadra ' + q + (variasObras() ? ' · ' + obraDe(q) : '')])];
const campoObra = (sel, empSel) => {
  const o = sel || E().obras[0]; const c = [];
  if (variasObras()) c.push(`<label class="so-obra">Obra<select name="obra">${opts(E().obras, o)}</select></label>`);
  if (variasEmpresas()) c.push(`<label>CNPJ da nota / pagamento<select name="emp" required>${opts(empOpts(), empSel || '', 'Escolha o CNPJ…')}</select></label>`);
  return (c.length === 2 ? `<div class="g2">${c.join('')}</div>` : c.join('')) + (variasEmpresas() ? '<p class="hint-data" id="empHint"></p>' : '');
};
const obraDoForm = F => F.ambito && F.ambito.value === 'empresa' ? '' : F.quadra.value !== '-' && E().quadras.includes(F.quadra.value) ? obraDe(F.quadra.value) : (F.obra ? F.obra.value : E().obras[0]);
const TIPOS_APORTE = [['capital', 'Capital do sócio (integralização)'], ['emprestimo', 'Empréstimo do sócio (a devolver)'], ['intragrupo', 'Entre empresas do grupo (matriz ↔ filial)'], ['retirada', 'Retirada / distribuição de lucro'], ['devolucao', 'Devolução de empréstimo ao sócio'], ['transferencia', 'Transferência entre contas da mesma empresa'], ['terceiro', 'Empréstimo / financiamento de terceiros'], ['conversivel', 'Dívida conversível (vira participação)']];
const tipoAporteNome = k => (TIPOS_APORTE.find(x => x[0] === k) || [k, k])[1];
const campoAporte = (sel, dest) => `<label id="rowAporte" class="${secaoDe(dest || '') === 'aporte' ? '' : 'hidden'}">Tipo de aporte<select name="aporteTipo">${opts(TIPOS_APORTE, sel || '', 'Escolha o tipo…')}</select></label>`;
const sincAporte = F => { const r = F.querySelector('#rowAporte'); if (r) r.classList.toggle('hidden', secaoDe(F.destino.value) !== 'aporte'); };
const campoAmbito = sel => `<label>A quem pertence<select name="ambito">${opts([['obra', 'Obra: quadra, lotes ou custo geral da obra'], ['empresa', 'Empresa: administrativo, fora das obras']], sel || 'obra')}</select></label>`;
const modoAmbito = F => { if (!F.ambito) return; const em = F.ambito.value === 'empresa'; F.classList.toggle('modo-empresa', em); if (em) { F.quadra.value = '-'; if (F.lote) F.lote.value = '-'; } };
const sincEmp = F => {
  const h = F.querySelector('#empHint'); if (!F.emp || !h) return;
  if (F.ambito && F.ambito.value === 'empresa') { h.className = 'hint-data info'; h.textContent = F.emp.value ? '' : 'Despesa da empresa: escolha o CNPJ que pagou (normalmente a matriz).'; return; }
  const o = obraDoForm(F), eo = empDaObra(o), x = empPorId(eo);
  if (!F.emp.value) { h.className = 'hint-data info'; h.innerHTML = `A obra ${esc(o)} é da ${esc(x.nome)}. Confira na nota em qual CNPJ o gasto foi emitido. <button type="button" class="btn link sm" data-usaemp="${esc(eo)}">Usar ${esc(x.nome)}</button>`; const b = h.querySelector('[data-usaemp]'); b.onclick = () => { F.emp.value = eo; sincEmp(F); }; }
  else if (F.emp.value !== eo) { h.className = 'hint-data'; h.textContent = `Atenção: a obra ${o} é da ${x.nome}, e este lançamento está em ${empPorId(F.emp.value)?.nome}. Ele fica marcado como "CNPJ diferente da obra" para conferência.`; }
  else { h.className = 'hint-data'; h.textContent = ''; }
};
const empDoForm = (F, obj) => { if (!F.emp) return; if (F.emp.value) obj.emp = F.emp.value; };

/* ---------- Leitura de boleto e nota fiscal ---------- */
const BANCOS = { '001': 'Banco do Brasil', '004': 'BNB', '033': 'Santander', '041': 'Banrisul', '070': 'BRB', '077': 'Inter', '104': 'Caixa', '208': 'BTG Pactual', '212': 'Original', '237': 'Bradesco', '260': 'Nubank', '290': 'PagBank', '323': 'Mercado Pago', '336': 'C6', '341': 'Itaú', '389': 'Mercantil', '422': 'Safra', '655': 'Votorantim', '745': 'Citibank', '748': 'Sicredi', '756': 'Sicoob' };
const soDig = s => String(s || '').replace(/\D/g, '');
const mod10 = s => { let t = 0, f = 2; for (let i = s.length - 1; i >= 0; i--) { let v = +s[i] * f; if (v > 9) v = Math.floor(v / 10) + (v % 10); t += v; f = f === 2 ? 1 : 2; } return (10 - (t % 10)) % 10; };
const mod11Bol = s => { let t = 0, f = 2; for (let i = s.length - 1; i >= 0; i--) { t += +s[i] * f; f = f === 9 ? 2 : f + 1; } const r = 11 - (t % 11); return r === 0 || r === 10 || r === 11 ? 1 : r; };
const mod11NFe = s => { let t = 0, f = 2; for (let i = s.length - 1; i >= 0; i--) { t += +s[i] * f; f = f === 9 ? 2 : f + 1; } const r = t % 11; return r < 2 ? 0 : 11 - r; };
function vencFator(fator) {
  if (!fator) return '';
  const h = hoje(); const c = [addDays('1997-10-07', fator), addDays('1997-10-07', fator + 9000)];
  const dist = d => Math.abs(new Date(d) - new Date(h)); return c.sort((a, b) => dist(a) - dist(b))[0];
}
function lerChaveNFe(ch) {
  const d = soDig(ch); if (d.length !== 44 || mod11NFe(d.slice(0, 43)) !== +d[43] || !['55', '65', '57'].includes(d.slice(20, 22))) return null;
  return { tipo: 'nfe', chave: d, fornCnpj: d.slice(6, 20), nNF: String(+d.slice(25, 34)), serie: String(+d.slice(22, 25)), emissao: '20' + d.slice(2, 4) + '-' + d.slice(4, 6) };
}
/* Linha digitável ou código de barras de boleto bancário ou de arrecadação (contas de consumo, tributos). */
function lerCodigoBoleto(txt) {
  const d = soDig(txt);
  const nfe = lerChaveNFe(d); if (nfe) return nfe;
  let cb = '';
  if (d.length === 47) cb = d.slice(0, 4) + d[32] + d.slice(33, 47) + d.slice(4, 9) + d.slice(10, 20) + d.slice(21, 31);
  else if (d.length === 48 && d[0] === '8') cb = d.slice(0, 11) + d.slice(12, 23) + d.slice(24, 35) + d.slice(36, 47);
  else if (d.length === 44) cb = d;
  else return null;
  if (cb[0] === '8') {
    const idv = cb[2]; const valor = ['6', '7'].includes(idv) ? +cb.slice(4, 15) / 100 : 0;
    const dv = ['6', '7'].includes(idv) ? mod10(cb.slice(0, 3) + cb.slice(4)) : mod11NFe(cb.slice(0, 3) + cb.slice(4));
    return { tipo: 'arrecadacao', codigo: cb, valor, valido: +cb[3] === dv || ['8', '9'].includes(idv) };
  }
  if (cb.length !== 44) return null;
  const valido = mod11Bol(cb.slice(0, 4) + cb.slice(5)) === +cb[4];
  return { tipo: 'boleto', codigo: cb, banco: BANCOS[cb.slice(0, 3)] || 'banco ' + cb.slice(0, 3), venc: vencFator(+cb.slice(5, 9)), valor: +cb.slice(9, 19) / 100, valido };
}
const numBR = s => +String(s).replace(/\./g, '').replace(',', '.');
const dataBRiso = s => { const m = String(s).match(/(\d{2})\/(\d{2})\/(\d{4})/); return m ? `${m[3]}-${m[2]}-${m[1]}` : ''; };
function lerXmlNFe(xml) {
  const doc = new DOMParser().parseFromString(xml, 'application/xml');
  const q = (sel, el = doc) => { const x = el.getElementsByTagName(sel)[0]; return x ? x.textContent.trim() : ''; };
  const inf = doc.getElementsByTagName('infNFe')[0]; if (!inf) return null;
  const emit = doc.getElementsByTagName('emit')[0], dest = doc.getElementsByTagName('dest')[0];
  const dups = [...doc.getElementsByTagName('dup')].map(d => ({ d: q('dVenc', d), v: +q('vDup', d) })).filter(x => x.d && x.v > 0);
  const det = [...doc.getElementsByTagName('det')].map(d => ({ prod: q('xProd', d), qtd: +q('qCom', d), unid: q('uCom', d).toLowerCase() }));
  const nome = emit ? (q('xFant', emit) || q('xNome', emit)) : '';
  return { tipo: 'nfe', chave: (inf.getAttribute('Id') || '').replace(/^NFe/, ''), nNF: q('nNF'), emissao: q('dhEmi').slice(0, 10) || q('dEmi'), forn: nome, fornCnpj: emit ? q('CNPJ', emit) || q('CPF', emit) : '', destCnpj: dest ? q('CNPJ', dest) || q('CPF', dest) : '', valor: +q('vNF') || 0, dups, det };
}
function lerTextoDocumento(tx) {
  const r = {}; const T = tx.replace(/ /g, ' ');
  const lin = T.match(/\d{5}[.\s]?\d{5}\s*\d{5}[.\s]?\d{6}\s*\d{5}[.\s]?\d{6}\s*\d\s*\d{14}/) || T.match(/8\d{10}[\s-]?\d\s*\d{11}[\s-]?\d\s*\d{11}[\s-]?\d\s*\d{11}[\s-]?\d/);
  if (lin) Object.assign(r, lerCodigoBoleto(lin[0]) || {});
  const ch = T.match(/(?:\d{4}\s?){11}/g); if (ch) for (const c of ch) { const n = lerChaveNFe(c); if (n) { r.nfe = n; break; } }
  const vt = T.match(/VALOR\s+TOTAL\s+DA\s+NOTA[^\d]{0,60}([\d.]+,\d{2})/i) || T.match(/VALOR\s+(?:L[IÍ]QUIDO|TOTAL)(?:\s+D[AO]S?\s+(?:NOTA|NFS-?E|SERVI[CÇ]OS?))?[^\d]{0,40}([\d.]+,\d{2})/i);
  if (vt) r.valorNota = numBR(vt[1]);
  const nn = T.match(/N[UÚ]MERO\s+DA\s+NFS-?E[^\d]{0,20}(\d{1,10})/i) || T.match(/NFS-?e\s*N[º°o.]*\s*:?\s*(\d{1,10})/i); if (nn) r.nNF = String(+nn[1]);
  const rec = T.match(/RECEBEMOS\s+DE\s+(.+?)\s+OS\s+PRODUTOS/i); if (rec) r.forn = rec[1].trim();
  const ben = T.match(/Benefici[aá]rio\s*:?\s*\n?\s*([^\n]{3,90})/i); if (ben && !/CPF|CNPJ|Ag[eê]ncia/i.test(ben[1].slice(0, 6))) r.benef = ben[1].replace(/\s*(CNPJ|CPF).*$/i, '').trim();
  r.cnpjs = [...new Set((T.match(/\d{2}\.?\d{3}\.?\d{3}\/?\d{4}-?\d{2}/g) || []).map(soDig).filter(x => x.length === 14))];
  const dups = [...T.matchAll(/\b(\d{3})\s+(\d{2}\/\d{2}\/\d{4})\s+(?:R\$\s*)?([\d.]+,\d{2})/g)].map(m => ({ d: dataBRiso(m[2]), v: numBR(m[3]) }));
  if (dups.length) r.dups = dups;
  if (!r.venc) { const vv = T.match(/Vencimento\s*:?\s*\n?\s*(\d{2}\/\d{2}\/\d{4})/i); if (vv) r.vencTexto = dataBRiso(vv[1]); }
  return r;
}
const tituloCase = s => String(s || '').toLowerCase().replace(/(^|\s|\/|-)(\p{L})/gu, (m, a, b) => a + b.toUpperCase()).replace(/\b(Ltda|Me|Epp|Eireli|S\/a|Sa)\b/gi, x => x.toUpperCase());
function fornPorCnpj(c) { const d = soDig(c); if (!d) return ''; const l = (S.fornecedores && S.fornecedores.lista) || {}; for (const n in l) if (soDig(l[n].cnpj) === d) return n; return ''; }
/* Junta o que foi encontrado em um resumo único. */
function dadosLidos(partes) {
  const o = {}; const p = Object.assign({}, ...partes.filter(Boolean));
  const nfe = p.tipo === 'nfe' ? p : p.nfe;
  o.valor = p.dups && p.dups.length > 1 ? sum(p.dups, x => x.v) : (p.tipo === 'nfe' && p.valor) || p.valorNota || p.valor || 0;
  if (p.tipo === 'boleto' || p.tipo === 'arrecadacao') { o.valor = p.valor || o.valor; o.venc = p.venc || p.vencTexto || ''; o.boleto = p; }
  if (p.dups && p.dups.length) { o.dups = p.dups; if (p.dups.length === 1) o.venc = o.venc || p.dups[0].d; }
  if (!o.venc && p.vencTexto) o.venc = p.vencTexto;
  if (nfe) { o.nNF = nfe.nNF; o.fornCnpj = nfe.fornCnpj; o.chave = nfe.chave; }
  if (p.nNF && !o.nNF) o.nNF = p.nNF;
  const empCnpjs = E().empresas.map(x => soDig(x.cnpj)).filter(Boolean);
  const dest = p.destCnpj || (p.cnpjs || []).find(c => empCnpjs.includes(c));
  if (dest) { const e = E().empresas.find(x => soDig(x.cnpj) === soDig(dest)); if (e) o.emp = e.id; o.destCnpj = dest; }
  if (!o.fornCnpj && p.fornCnpj) o.fornCnpj = p.fornCnpj;
  if (!o.fornCnpj && p.cnpjs) o.fornCnpj = p.cnpjs.find(c => !empCnpjs.includes(c)) || '';
  o.forn = fornPorCnpj(o.fornCnpj) || (p.forn ? tituloCase(p.forn) : '') || (p.benef ? tituloCase(p.benef) : '');
  if (p.det && p.det.length === 1 && p.det[0].qtd > 0) { o.qtd = p.det[0].qtd; o.unid = p.det[0].unid; o.prod = p.det[0].prod; }
  o.ehNota = !!(nfe || p.nNF || p.valorNota);
  return o;
}
async function carregarLeitor() {
  if (window.SFGLeitor) return window.SFGLeitor;
  if (!navigator.onLine) throw new Error('A leitura de PDF e foto precisa de internet na primeira vez.');
  await new Promise((res, rej) => { const s = document.createElement('script'); s.src = 'leitor.js?v=__VERSAO__'; s.onload = res; s.onerror = () => rej(new Error('Não consegui carregar o leitor.')); document.head.appendChild(s); });
  return window.SFGLeitor;
}
async function lerArquivoDocumento(file) {
  const nome = (file.name || '').toLowerCase();
  if (nome.endsWith('.xml') || file.type.includes('xml')) { const x = lerXmlNFe(await file.text()); if (!x) throw new Error('Esse XML não parece ser de uma NF-e.'); return x; }
  const L = await carregarLeitor();
  if (file.type === 'application/pdf' || nome.endsWith('.pdf')) { const tx = await L.textoPDF(await file.arrayBuffer()); if (!tx.trim()) throw new Error('O PDF não tem texto (parece uma imagem escaneada). Tente uma foto mais nítida do código de barras.'); return lerTextoDocumento(tx); }
  const cod = await L.codigoImagem(file); if (!cod) throw new Error('Não encontrei código de barras ou QR code na foto. Tente aproximar e deixar o código reto e bem iluminado.');
  const ch = (cod.match(/\d{44}/) || [])[0]; return lerCodigoBoleto(ch && lerChaveNFe(ch) ? ch : cod) || (ch ? lerChaveNFe(ch) : null) || {};
}

/* ---------- Novo lançamento ---------- */
function formNovo(def = {}) {
  const h = hoje();
  const f = folha('Novo lançamento', `<form id="fNovo" class="form">
    <details class="leitor"><summary>${ic('scan', 16)} Ler boleto ou nota fiscal</summary>
      <p class="muted small">Cole a linha digitável do boleto ou a chave da nota, ou escolha o PDF do boleto, o PDF ou o XML da nota, ou uma foto do código de barras. O app preenche o que encontrar e guarda o arquivo como anexo.</p>
      <textarea name="codigo" rows="2" placeholder="Linha digitável, código de barras ou chave de 44 dígitos"></textarea>
      <div class="leitor-acts"><label class="btn sm file">Escolher arquivo ou foto<input type="file" id="lerArq" accept="application/pdf,.xml,text/xml,image/*" multiple hidden></label><button type="button" class="btn sm primary" id="lerOk">Ler</button></div>
      <div id="lerRes" class="lerres"></div></details>
    <input type="hidden" name="fornCnpj">
    <div class="seg full" id="segTipo"><button type="button" data-v="saida" class="${def.tipo !== 'entrada' ? 'on' : ''}">Saída (pagar)</button><button type="button" data-v="entrada" class="${def.tipo === 'entrada' ? 'on' : ''}">Entrada (receber)</button></div>
    <label class="chk"><input type="checkbox" name="realizado" ${def.realizado ? 'checked' : ''}> Já foi pago/recebido (lança direto no extrato)</label>
    <label>Descrição<input name="item" required placeholder="Ex.: Mendonça NF, Parcela terreno, Venda" list="dlItens" value="${esc(def.item || '')}"></label>
    <div class="g2"><label>Fornecedor / cliente<input name="forn" list="dlForn" value="${esc(def.forn || '')}"></label>
      <label>Destino<select name="destino" required>${optsDestino(def.destino || (def.tipo === 'entrada' ? 'Venda' : 'Material'))}</select></label></div>
    ${campoAporte(def.aporteTipo, def.destino)}
    <label id="rowEtapa" class="so-obra ${def.tipo === 'entrada' ? 'hidden' : ''}">Etapa da obra <span class="muted">(opcional, para o orçamento por etapa)</span><select name="etapaObra">${opts(etapasOrc(), def.etapaObra || '', 'Sem etapa')}</select></label>
    <div class="g2"><label>Valor (R$)<input name="valor" inputmode="decimal" required placeholder="0,00" value="${def.valor ? fNum2.format(Math.abs(def.valor)) : ''}"></label>
      <label>O valor informado é<select name="modo"><option value="total">Total (dividir entre lotes e parcelas)</option><option value="cada">De cada lançamento</option></select></label></div>
    ${campoAmbito(def.ambito)}
    ${campoObra(def.quadra && def.quadra !== '-' ? obraDe(def.quadra) : '', def.emp)}
    <input type="hidden" name="quadra" value="${esc(def.quadra || '-')}">
    <div class="so-obra"><div class="lbl">Quadras <span class="muted">(marque uma ou mais; nenhuma = custo geral da obra)</span></div><div class="chips" id="boxQuadras"></div></div>
    <div id="boxLotes" class="so-obra"></div>
    <label>Vencimento${'<input type="date" name="venc" required value="' + (def.venc || h) + '">'}</label>
    <label class="chk"><input type="checkbox" name="recorrente" ${def.recorrente ? 'checked' : ''}> Conta recorrente, sem data final (contador, aluguel, internet…)</label>
    <div id="boxRec" class="g2 hidden"><label>Repete<select name="freq">${opts(FREQ, '1')}</select></label><label>Termina em <span class="muted">(opcional)</span><input type="date" name="fimRec"></label></div>
    <div class="g2" id="rowParc"><label>Parcelas<input name="parcelas" type="number" min="1" max="240" value="1" inputmode="numeric"></label>
      <label>Intervalo<select name="intervalo"><option value="1">Mensal</option><option value="0.5">Quinzenal</option><option value="7d">Semanal</option></select></label></div>
    <label class="chk"><input type="checkbox" name="util" ${prefUtil() ? 'checked' : ''}> Vencimento em fim de semana ou feriado vai para o próximo dia útil</label>
    <div id="boxParc"></div>
    <div id="boxReal" class="g3"><label>Data do pagamento<input type="date" name="pag" value="${def.pag || h}"></label><label>Conta<select name="conta">${opts(E().contas.map(c => c.nome), def.conta)}</select></label><label>Juros / multa (R$) <span class="muted">fora do valor</span><input name="jurosR" inputmode="decimal" placeholder="0,00"></label></div>
    <div class="g2"><label>Quantidade <span class="muted">(opcional, para histórico de preço)</span><input name="qtd" inputmode="decimal" placeholder="Ex.: 200"></label><label>Unidade<input name="unid" list="dlUnid" placeholder="sc, m³, un, kg"></label></div><datalist id="dlUnid"><option value="sc"><option value="m³"><option value="m²"><option value="un"><option value="kg"><option value="t"><option value="milheiro"><option value="m"></datalist>
    <div class="g2"><label class="chk"><input type="checkbox" name="nf"> Tem nota fiscal</label><label>Número da NF<input name="nfNum" inputmode="numeric" placeholder="opcional"></label></div>
    <label>Observação<textarea name="obs" rows="2"></textarea></label>
    <div id="boxAnexos" class="anx"></div>
    <div class="preview" id="prevNovo"></div>
    <div class="acts"><button type="button" class="btn" data-close>Cancelar</button><button class="btn primary" type="submit">Salvar</button></div>
    <datalist id="dlItens">${[...new Set(S.list.slice(-3000).map(t => t.item.replace(/\s*[-–]\s*[A-Z]\d?\s*[\d.]*\s*$/, '').replace(/\s*\(\d+\/\d+\)\s*/, ' ').trim()))].slice(0, 300).map(x => `<option value="${esc(x)}">`).join('')}</datalist>
    <datalist id="dlForn">${[...new Set(S.list.map(t => t.forn).filter(Boolean))].map(x => `<option value="${esc(x)}">`).join('')}</datalist>
  </form>`);
  const F = f.el.querySelector('#fNovo');
  let tipo = def.tipo === 'entrada' ? 'entrada' : 'saida';
  $$('#segTipo button', F).forEach(b => b.onclick = () => { tipo = b.dataset.v; $$('#segTipo button', F).forEach(x => x.classList.toggle('on', x === b)); F.destino.value = tipo === 'entrada' ? 'Venda' : 'Material'; $('#rowEtapa', F).classList.toggle('hidden', tipo === 'entrada'); prev(); });
  const qsMarcadas = () => $$('[name=qsel]:checked', F).map(c => c.value);
  const quadrasBox = () => {
    const ob = F.obra ? F.obra.value : ''; const marc = new Set(qsMarcadas().length ? qsMarcadas() : (F.quadra.value !== '-' ? [F.quadra.value] : []));
    const qs = quadrasDaObra(variasObras() ? ob : '');
    $('#boxQuadras', F).innerHTML = qs.map(q => `<label class="pick"><input type="checkbox" name="qsel" value="${esc(q)}" ${marc.has(q) ? 'checked' : ''}><span>${esc(q)}</span></label>`).join('') || '<span class="muted small">Nenhuma quadra nesta obra.</span>';
  };
  const lotesBox = () => {
    const ant = new Set($$('[name=lotes]:checked', F).map(c => c.value)); if (def.lote && def.quadra && !lotesBox.feito) ant.add(def.quadra + '|' + def.lote); lotesBox.feito = true;
    const qs = qsMarcadas(); F.quadra.value = qs[0] || '-';
    $('#boxLotes', F).innerHTML = qs.filter(q => lotesDe(q).length).map(q => `<div class="lbl">Lotes da quadra ${esc(q)} <span class="muted">(nenhum = custo geral da quadra)</span> <button type="button" class="btn link sm" data-lt="all" data-q="${esc(q)}">Todos</button><button type="button" class="btn link sm" data-lt="none" data-q="${esc(q)}">Limpar</button></div>
      <div class="chips">${lotesDe(q).map(l => `<label class="pick"><input type="checkbox" name="lotes" value="${esc(q + '|' + l)}" ${ant.has(q + '|' + l) ? 'checked' : ''}><span>${esc(l)}</span></label>`).join('')}</div>`).join('');
    $$('[data-lt]', F).forEach(b => b.onclick = () => { $$('[name=lotes]', F).filter(c => c.value.startsWith(b.dataset.q + '|')).forEach(c => c.checked = b.dataset.lt === 'all'); recalcParc(); prev(); });
  };
  const real = () => { $('#boxReal', F).classList.toggle('hidden', !F.realizado.checked); };
  let parc = [];
  const nParc = () => Math.max(1, Math.min(240, parseInt(F.parcelas.value) || 1));
  const pares = () => { if (F.ambito.value === 'empresa') return [['-', '-']]; const qs = qsMarcadas(); if (!qs.length) return [['-', '-']]; const ls = $$('[name=lotes]:checked', F).map(c => c.value.split('|')); return qs.flatMap(q => { const x = ls.filter(p => p[0] === q); return x.length ? x : [[q, '-']]; }); };
  const nLotes = () => pares().length;
  const dataBase = i => { const v0 = F.venc.value, iv = F.intervalo.value; if (!v0) return ''; return iv === '7d' ? addDays(v0, 7 * i) : iv === '0.5' ? (i % 2 ? addDays(addMonths(v0, Math.floor(i / 2)), 15) : addMonths(v0, i / 2)) : addMonths(v0, i); };
  const ajustada = d => F.util.checked && d ? proximoDiaUtil(d) : d;
  const notaData = d => { const m = motivoNaoUtil(d); if (!d || !m) return ''; return F.util.checked ? `${m}: vai para ${dCurtoSem(proximoDiaUtil(d))}` : `cai em ${m}`; };
  const recalcParc = (manterDatas) => {
    const n = nParc(), v = parseBR(F.valor.value) || 0, total = F.modo.value === 'cada' ? v * nLotes() * n : v;
    const novo = Array.from({ length: n }, (_, i) => { const o = parc[i]; return { d: o && o.dEdit && manterDatas !== false ? o.d : dataBase(i), dEdit: !!(o && o.dEdit && manterDatas !== false), v: o && o.vEdit ? o.v : null, vEdit: !!(o && o.vEdit) }; });
    const fixo = sum(novo.filter(p => p.vEdit), p => p.v || 0), livres = novo.filter(p => !p.vEdit);
    const cents = Math.max(0, Math.round((total - fixo) * 100)); const base = livres.length ? Math.floor(cents / livres.length) : 0; let resto = cents - base * livres.length;
    livres.forEach(p => { p.v = (base + (resto-- > 0 ? 1 : 0)) / 100; });
    parc = novo; renderParc();
  };
  const renderParc = () => {
    const box = $('#boxParc', F); const n = parc.length;
    if (n <= 1) { const nt = notaData(F.venc.value); box.innerHTML = nt ? `<p class="hint-data">Vencimento ${dBR(F.venc.value).slice(0, 5)} ${esc(nt)}.</p>` : ''; return; }
    box.innerHTML = `<div class="lbl">Parcelas <span class="muted">altere a data ou o valor de qualquer uma</span> <button type="button" class="btn link sm" id="pReset">Refazer datas</button></div>
      <div class="parc-tbl">${parc.map((p, i) => `<div class="parc-row"><span class="pn">${i + 1}/${n}</span><input type="date" data-pd="${i}" value="${esc(p.d)}" aria-label="Data da parcela ${i + 1}"><input inputmode="decimal" data-pv="${i}" value="${fNum2.format(p.v || 0)}" aria-label="Valor da parcela ${i + 1}"><span class="pnota" id="pn${i}">${esc(notaData(p.d))}</span></div>`).join('')}</div>`;
    $$('[data-pd]', box).forEach(inp => inp.oninput = () => { const i = +inp.dataset.pd; parc[i].d = inp.value; parc[i].dEdit = true; $('#pn' + i, box).textContent = notaData(inp.value); prev(); });
    $$('[data-pv]', box).forEach(inp => inp.onchange = () => { const i = +inp.dataset.pv; const v = parseBR(inp.value); parc[i].v = v > 0 ? v : 0; parc[i].vEdit = true; recalcParc(); prev(); });
    $('#pReset', box).onclick = () => { parc.forEach(p => { p.dEdit = false; p.vEdit = false; }); recalcParc(false); prev(); };
  };
  F.addEventListener('input', e => { const nm = e.target.name; if (['venc', 'parcelas', 'intervalo', 'valor', 'modo', 'lotes'].includes(nm)) recalcParc(); });
  F.addEventListener('change', e => { const nm = e.target.name; if (['intervalo', 'modo', 'lotes'].includes(nm)) recalcParc(); if (nm === 'util') { setPrefUtil(F.util.checked); renderParc(); } });
  const gerar = () => {
    const v = parseBR(F.valor.value); const n = parc.length || 1;
    const L = pares();
    const total = L.length * n;
    if (!(v > 0)) return { erro: 'Informe um valor maior que zero.', total };
    if (parc.some(p => !p.d)) return { erro: 'Preencha a data de todas as parcelas.', total };
    const base = F.item.value.trim();
    const out = []; const grupo = total > 1 ? uid() : ''; let ajustes = 0;
    for (let i = 0; i < n; i++) {
      const venc = ajustada(parc[i].d); if (venc !== parc[i].d) ajustes++;
      const cents = Math.round((parc[i].v || 0) * 100); const bL = Math.floor(cents / L.length); let rL = cents - bL * L.length;
      const valores = L.map(() => (bL + (rL-- > 0 ? 1 : 0)) / 100); let k = 0;
      for (const [q, l] of L) {
        let item = base + (n > 1 ? ` (${i + 1}/${n})` : '') + (l !== '-' ? ` - ${q} ${l}` : (q !== '-' ? ` - ${q}` : ''));
        const valor = (tipo === 'saida' ? -1 : 1) * valores[k++];
        if (!valor) continue;
        const t = { id: uid(), item, valor, juros: 0, ...(F.fornCnpj.value ? { fornCnpj: F.fornCnpj.value } : {}), nf: F.nf.checked || F.nfNum.value.trim() ? 1 : 0, nfNum: F.nfNum.value.trim(), local: q !== '-' && E().quadras.includes(q) ? obraDe(q) : obraDoForm(F), quadra: q, lote: l, destino: F.destino.value, venc, status: 'aberto', forn: F.forn.value.trim(), obs: F.obs.value.trim(), grupo, criado: Date.now() };
        if (F.realizado.checked) { t.status = 'pago'; t.pag = n > 1 ? venc : (F.pag.value || venc); t.conta = F.conta.value; const jr = parseBR(F.jurosR.value) || 0; if (jr > 0) t.juros = (tipo === 'saida' ? -1 : 1) * r2(jr / total); }
        if (venc !== parc[i].d) t.vencOrig = parc[i].d;
        if (F.etapaObra.value && tipo === 'saida' && F.ambito.value !== 'empresa') t.etapaObra = F.etapaObra.value;
        if (F.ambito.value === 'empresa') { t.ambito = 'empresa'; t.quadra = '-'; t.lote = '-'; t.local = ''; }
        if (secaoDe(F.destino.value) === 'aporte' && F.aporteTipo.value) t.aporteTipo = F.aporteTipo.value;
        empDoForm(F, t);
        const qt = parseBR(F.qtd.value); if (qt > 0) { t.qtd = Math.round((F.modo.value === 'cada' ? qt : qt / total) * 1000) / 1000; t.unid = F.unid.value.trim(); }
        if (def.conc && total === 1) t.conc = def.conc;
        out.push(t);
      }
    }
    return { lista: out, total: out.length, ajustes };
  };
  const ehRec = () => F.recorrente.checked;
  const modoRec = () => {
    const r = ehRec(); $('#boxRec', F).classList.toggle('hidden', !r); $('#rowParc', F).classList.toggle('hidden', r); $('#boxParc', F).classList.toggle('hidden', r);
    if (r) { F.realizado.checked = false; F.parcelas.value = 1; recalcParc(); real(); } F.realizado.closest('label').classList.toggle('hidden', r);
  };
  const prev = () => {
    const g = gerar(); const p = $('#prevNovo', F);
    if (!F.item.value.trim()) { p.innerHTML = ''; return; }
    if (ehRec()) { const v = parseBR(F.valor.value); const fr = FREQ.find(x => x[0] === F.freq.value)[1].toLowerCase(); p.innerHTML = v > 0 ? `<b>Conta recorrente ${fr}</b> de <b class="${tipo === 'saida' ? 'neg' : 'pos'}">${brl2(v)}</b>, primeiro vencimento ${dCurtoSem(F.util.checked ? proximoDiaUtil(F.venc.value) : F.venc.value)}${F.fimRec.value ? ', até ' + dBR(F.fimRec.value) : ', sem data final'}.<br><span class="muted">O app mantém os próximos ${HORIZ_REC} meses lançados em Contas a ${tipo === 'saida' ? 'pagar' : 'receber'} e cria os seguintes sozinho. Gerencie em Contas recorrentes.</span>` : ''; return; }
    if (g.erro) { p.innerHTML = `<span class="muted">${g.total} lançamento(s)</span>`; return; }
    const tot = sum(g.lista, t => t.valor);
    if (!g.lista.length) { p.innerHTML = '<span class="muted">Informe valores maiores que zero.</span>'; return; }
    p.innerHTML = `<b>${g.total} lançamento${g.total > 1 ? 's' : ''}</b>, total <b class="${cls(tot)}">${brl2(tot)}</b>${g.ajustes ? `, ${g.ajustes} vencimento(s) passado(s) para dia útil` : ''}<br><span class="muted">Primeiro: ${esc(g.lista[0].item)}, ${brl2(g.lista[0].valor)}, vence ${dCurtoSem(g.lista[0].venc)}${g.total > 1 ? `. Último vence ${dCurtoSem(g.lista[g.lista.length - 1].venc)}` : ''}</span>`;
  };
  $('#boxQuadras', F).addEventListener('change', () => { lotesBox(); sincEmp(F); recalcParc(); prev(); });
  if (F.obra) F.obra.onchange = () => { $$('[name=qsel]', F).forEach(c => c.checked = false); F.quadra.value = '-'; quadrasBox(); lotesBox(); recalcParc(); prev(); sincEmp(F); };
  if (F.obra && def.quadra && def.quadra !== '-') F.obra.value = obraDe(def.quadra);
  quadrasBox();
  if (F.emp) { F.emp.onchange = () => sincEmp(F); sincEmp(F); }
  F.ambito.onchange = () => { modoAmbito(F); if (F.ambito.value === 'empresa') $$('[name=qsel]', F).forEach(c => c.checked = false); lotesBox(); recalcParc(); sincEmp(F); prev(); }; modoAmbito(F);
  F.destino.addEventListener('change', () => sincAporte(F)); sincAporte(F);
  F.item.addEventListener('change', () => aplicarRegrasForm(F, tipo));
  F.realizado.onchange = () => { real(); prev(); };
  F.recorrente.onchange = () => { modoRec(); prev(); }; modoRec();
  F.addEventListener('input', prev); F.addEventListener('change', prev);
  lotesBox(); real(); recalcParc(); prev();
  const anx = caixaAnexos($('#boxAnexos', F), () => F.realizado.checked ? 'comprovante' : 'nota');
  const aplicarLido = (o, arquivos) => {
    const feito = [];
    const tipoAtual = () => $('#segTipo .on', F)?.dataset.v;
    if (tipoAtual() !== 'saida') $('#segTipo [data-v=saida]', F).click();
    if (o.forn) { F.forn.value = o.forn; feito.push('fornecedor'); }
    if (o.fornCnpj) F.fornCnpj.value = o.fornCnpj;
    if (o.nNF) { F.nfNum.value = o.nNF; F.nf.checked = true; feito.push('NF ' + o.nNF); } else if (o.ehNota) F.nf.checked = true;
    if (!F.item.value.trim()) F.item.value = (o.forn || (o.boleto && o.boleto.banco ? 'Boleto ' + o.boleto.banco : 'Conta')) + (o.nNF ? ' NF ' + o.nNF : '');
    if (o.emp && F.emp) { F.emp.value = o.emp; sincEmp(F); feito.push('CNPJ da nota'); }
    if (o.qtd) { F.qtd.value = String(o.qtd).replace('.', ','); F.unid.value = o.unid || ''; feito.push('quantidade'); }
    F.modo.value = 'total';
    if (o.dups && o.dups.length > 1) {
      F.valor.value = fNum2.format(o.valor); F.parcelas.value = o.dups.length; F.venc.value = o.dups[0].d; recalcParc();
      o.dups.forEach((x, i) => { if (parc[i]) { parc[i].d = x.d; parc[i].dEdit = true; parc[i].v = x.v; parc[i].vEdit = true; } }); renderParc();
      feito.push(`${o.dups.length} parcelas com datas e valores da nota`);
    } else {
      if (o.valor) { F.valor.value = fNum2.format(o.valor); feito.push('valor'); }
      if (o.venc) { F.venc.value = o.venc; feito.push('vencimento'); }
      recalcParc();
    }
    for (const a of arquivos || []) anx.adicionar(a, /xml/i.test(a.type + a.name) ? 'nota' : o.boleto ? 'boleto' : 'nota');
    if (arquivos && arquivos.length) feito.push('anexo');
    prev();
    return feito;
  };
  const lerRes = $('#lerRes', F);
  const mostrarLido = (o, feito) => {
    const avisos = [];
    if (o.boleto && o.boleto.valido === false) avisos.push('O dígito verificador do código não confere. Confira se foi digitado certo.');
    if (o.boleto && o.boleto.tipo === 'arrecadacao' && !o.venc) avisos.push('Contas de consumo e tributos não trazem o vencimento no código: informe a data.');
    if (o.chave && !o.valor) avisos.push('A chave da nota não traz o valor: informe o valor ou leia o PDF/XML da nota.');
    if (o.destCnpj && !o.emp && variasEmpresas()) avisos.push('A nota foi emitida para o CNPJ ' + fCnpj(o.destCnpj) + ', que não está cadastrado em Empresas.');
    lerRes.innerHTML = feito.length ? `<p class="ok small">Preenchido: ${esc(feito.join(', '))}. Confira antes de salvar.</p>${avisos.map(a => `<p class="hint-data">${esc(a)}</p>`).join('')}` : `<p class="hint-data">Não encontrei dados para preencher.</p>`;
  };
  $('#lerOk', F).onclick = () => { const c = F.codigo.value.trim(); if (!c) return; const r = lerCodigoBoleto(c); if (!r) { lerRes.innerHTML = '<p class="hint-data">Não reconheci esse número. A linha digitável tem 47 ou 48 dígitos; a chave da nota tem 44.</p>'; return; } const o = dadosLidos([r]); mostrarLido(o, aplicarLido(o)); };
  $('#lerArq', F).onchange = async ev => {
    const arqs = [...ev.target.files]; if (!arqs.length) return; lerRes.innerHTML = '<p class="muted small">Lendo…</p>';
    const partes = [], erros = [];
    for (const a of arqs) { try { partes.push(await lerArquivoDocumento(a)); } catch (e) { erros.push(e.message); } }
    if (F.codigo.value.trim()) partes.push(lerCodigoBoleto(F.codigo.value.trim()));
    const o = dadosLidos(partes); const feito = partes.filter(Boolean).length ? aplicarLido(o, arqs) : [];
    mostrarLido(o, feito); if (erros.length) lerRes.innerHTML += erros.map(e => `<p class="hint-data">${esc(e)}</p>`).join('');
    ev.target.value = '';
  };
  F.onsubmit = async e => {
    e.preventDefault(); if (F.emp && !F.emp.value) { toast('Escolha o CNPJ do lançamento.', 'err'); F.emp.focus(); return; }
    if (ehRec()) {
      const v = parseBR(F.valor.value); if (!(v > 0)) { toast('Informe um valor maior que zero.', 'err'); return; }
      if ($$('[name=lotes]:checked', F).length || qsMarcadas().length > 1) { toast('Conta recorrente vale para uma quadra só, sem lotes. Ajuste a seleção.', 'err'); return; }
      if (F.fimRec.value && F.fimRec.value < F.venc.value) { toast('A data final é anterior ao primeiro vencimento.', 'err'); return; }
      const r = { id: uid(), item: F.item.value.trim(), valor: (tipo === 'saida' ? -1 : 1) * r2(v), destino: F.destino.value, forn: F.forn.value.trim(), local: obraDoForm(F), quadra: F.quadra.value, ambito: F.ambito.value === 'empresa' ? 'empresa' : '', ...(secaoDe(F.destino.value) === 'aporte' && F.aporteTipo.value ? { aporteTipo: F.aporteTipo.value } : {}), emp: F.emp ? F.emp.value : '', obs: F.obs.value.trim(), inicio: F.venc.value, passo: +F.freq.value, fim: F.fimRec.value || '', util: F.util.checked, ativo: true, criado: Date.now(), por: S.user?.email || '' };
      salvarRec(r); f.fechar(); gerarRecorrencias(); toast('Conta recorrente criada'); return;
    }
    const g = gerar(); if (g.erro) { toast(g.erro, 'err'); return; } if (!g.lista.length) { toast('Informe valores maiores que zero.', 'err'); return; }
    if (g.total > 60 && !(await confirmar(`Isso vai criar <b>${g.total}</b> lançamentos. Continuar?`, 'Criar', false))) return;
    if (anx.itens.length) { try { const lst = await anx.salvar(); g.lista.forEach(t => t.anexos = lst.map(x => ({ ...x }))); } catch (err) { toast(err.message, 'err'); return; } }
    salvarTitulos(g.lista); f.fechar(); toast(g.total > 1 ? `${g.total} lançamentos criados` : 'Lançamento salvo');
    const fc = F.fornCnpj.value, fn = F.forn.value.trim(); if (fc && fn && !fornPorCnpj(fc) && !((S.fornecedores && S.fornecedores.lista) || {})[fn]) salvarForn(fn, { cnpj: fCnpj(fc), cat: F.destino.value });
  };
}

/* ---------- Editar ---------- */
function formEditar(t) {
  const lots = t.quadra && t.quadra !== '-' ? lotesDe(t.quadra) : [];
  const f = folha('Editar lançamento', `<form id="fEd" class="form">
    <div class="seg full" id="segTipo"><button type="button" data-v="saida" class="${t.valor < 0 ? 'on' : ''}">Saída</button><button type="button" data-v="entrada" class="${t.valor >= 0 ? 'on' : ''}">Entrada</button></div>
    <label>Descrição<input name="item" required value="${esc(t.item)}"></label>
    <div class="g2"><label>Fornecedor / cliente<input name="forn" value="${esc(t.forn || '')}"></label><label>Destino<select name="destino">${optsDestino(t.destino)}${todosDestinos().includes(t.destino) ? '' : `<option selected>${esc(t.destino)}</option>`}</select></label></div>
    ${campoAporte(t.aporteTipo, t.destino)}
    ${t.valor < 0 ? `<label class="so-obra">Etapa da obra<select name="etapaObra">${opts(etapasOrc().concat(t.etapaObra && !etapasOrc().includes(t.etapaObra) ? [t.etapaObra] : []), t.etapaObra || '', 'Sem etapa')}</select></label>` : ''}
    <div class="g2"><label>Valor (R$)<input name="valor" inputmode="decimal" required value="${fNum2.format(Math.abs(t.valor))}"></label><label>Vencimento<input type="date" name="venc" required value="${esc(t.venc)}"></label></div>
    <label class="chk"><input type="checkbox" name="util" ${prefUtil() ? 'checked' : ''}> Se cair em fim de semana ou feriado, passar para o próximo dia útil</label><p class="hint-data" id="edHint"></p>
    ${campoAmbito(t.ambito)}
    ${campoObra(obraT(t) || E().obras[0], empDef(t) ? t.emp : '')}
    <div class="g2 so-obra"><label>Quadra<select name="quadra">${opts(quadrasOpts(), t.quadra || '-')}${(t.quadra && t.quadra !== '-' && !E().quadras.includes(t.quadra)) ? `<option selected>${esc(t.quadra)}</option>` : ''}</select></label>
      <label>Lote<select name="lote">${opts([['-', 'Sem lote (-)'], ...lots], t.lote || '-')}${(t.lote && t.lote !== '-' && !lots.includes(t.lote)) ? `<option selected>${esc(t.lote)}</option>` : ''}</select></label></div>
    ${pago(t) ? `<div class="g3"><label>Pago em<input type="date" name="pag" value="${esc(t.pag || t.venc)}"></label><label>Conta<select name="conta">${opts(E().contas.map(c => c.nome), t.conta)}</select></label><label>Juros/multa (R$)<input name="juros" inputmode="decimal" value="${t.juros ? fNum2.format(Math.abs(t.juros)) : ''}"></label></div>` : ''}
    <div class="g2"><label>Quantidade<input name="qtd" inputmode="decimal" value="${t.qtd ? String(t.qtd).replace('.', ',') : ''}"></label><label>Unidade<input name="unid" value="${esc(t.unid || '')}"></label></div>
    <div class="g2"><label class="chk"><input type="checkbox" name="nf" ${t.nf ? 'checked' : ''}> Tem nota fiscal</label><label>Número da NF<input name="nfNum" inputmode="numeric" value="${esc(t.nfNum || '')}"></label></div>
    <label>Observação<textarea name="obs" rows="2">${esc(t.obs || '')}</textarea></label>
    <div class="acts"><button type="button" class="btn" data-close>Cancelar</button><button class="btn primary" type="submit">Salvar</button></div></form>`);
  const F = f.el.querySelector('#fEd'); let tipo = t.valor < 0 ? 'saida' : 'entrada';
  $$('#segTipo button', F).forEach(b => b.onclick = () => { tipo = b.dataset.v; $$('#segTipo button', F).forEach(x => x.classList.toggle('on', x === b)); });
  const edHint = () => { const m = motivoNaoUtil(F.venc.value); $('#edHint', F).textContent = m ? (F.util.checked ? `Cai em ${m}: será salvo para ${dCurtoSem(proximoDiaUtil(F.venc.value))}.` : `Cai em ${m}.`) : ''; };
  F.venc.oninput = edHint; F.util.onchange = () => { setPrefUtil(F.util.checked); edHint(); }; edHint();
  F.quadra.onchange = () => { const q = F.quadra.value; if (F.obra && q !== '-') F.obra.value = obraDe(q); sincEmp(F); F.lote.innerHTML = opts([['-', 'Sem lote (-)'], ...(q === '-' ? [] : lotesDe(q))], '-'); };
  if (F.obra) F.obra.onchange = () => { if (F.quadra.value !== '-' && obraDe(F.quadra.value) !== F.obra.value) { F.quadra.value = '-'; F.quadra.onchange(); } sincEmp(F); };
  if (F.emp) { F.emp.onchange = () => sincEmp(F); sincEmp(F); }
  F.ambito.onchange = () => { modoAmbito(F); sincEmp(F); }; modoAmbito(F);
  F.destino.addEventListener('change', () => sincAporte(F));
  F.onsubmit = e => {
    e.preventDefault(); const v = parseBR(F.valor.value); if (!(v >= 0)) { toast('Valor inválido', 'err'); return; }
    const sg = tipo === 'saida' ? -1 : 1;
    const n = { ...t, item: F.item.value.trim(), forn: F.forn.value.trim(), destino: F.destino.value, valor: sg * v, venc: F.util.checked ? proximoDiaUtil(F.venc.value) : F.venc.value, quadra: F.quadra.value, local: F.quadra.value !== '-' && E().quadras.includes(F.quadra.value) ? obraDe(F.quadra.value) : (F.obra ? F.obra.value : (t.local || E().obras[0])), lote: F.lote.value, nf: F.nf.checked || F.nfNum.value.trim() ? 1 : 0, nfNum: F.nfNum.value.trim(), obs: F.obs.value.trim() };
    empDoForm(F, n);
    if (F.etapaObra) { if (F.etapaObra.value) n.etapaObra = F.etapaObra.value; else delete n.etapaObra; }
    if (F.ambito.value === 'empresa') { n.ambito = 'empresa'; n.quadra = '-'; n.lote = '-'; n.local = ''; delete n.etapaObra; } else delete n.ambito;
    if (secaoDe(n.destino) === 'aporte' && F.aporteTipo.value) n.aporteTipo = F.aporteTipo.value; else delete n.aporteTipo;
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
    <label>Comprovante <span class="muted">(opcional, foto ou PDF)</span><input type="file" name="comp" accept="image/*,application/pdf"></label>
    <div class="acts"><button type="button" class="btn" data-close>Cancelar</button><button class="btn primary" type="submit">Confirmar baixa</button></div></form>`);
  const F = f.el.querySelector('#fBx');
  F.onsubmit = async e => {
    e.preventDefault(); const j = parseBR(F.juros.value) || 0;
    let comp = null; if (F.comp.files[0]) { try { comp = await salvarAnexo(F.comp.files[0]); } catch (err) { toast(err.message, 'err'); return; } }
    const out = ts.map(t => {
      const sg = t.valor < 0 ? -1 : 1; const n = { ...t, status: 'pago', pag: F.pag.value, conta: F.conta.value };
      if (um) { const v = parseBR(F.valor.value); if (v >= 0) n.valor = sg * v; n.juros = j > 0 ? sg * j : 0; }
      else n.juros = j > 0 ? sg * r2(j * Math.abs(t.valor) / Math.abs(tot || 1)) : 0;
      return comp ? comAnexos(n, [...anexosDe(n), { id: comp, tipo: 'comprovante' }]) : n;
    });
    salvarTitulos(out); f.fechar(); listaSel.clear(); toast(um ? 'Baixa registrada' : `${out.length} baixas registradas`); refreshView();
  };
}

ACOES.histT = async d => { try { const hs = (await carregarHistCompleto()).filter(x => x.t === d.id).sort((a, b) => b.em - a.em); folha('Histórico completo', hs.length ? `<ol class="hist-ol">${hs.map(h => `<li>${histTxt(h)}</li>`).join('')}</ol>` : '<p class="muted">As mudanças anteriores a esta versão do app ficam só no resumo do lançamento.</p>'); } catch (e) { toast('Não consegui carregar o histórico: ' + e.message, 'err'); } };
/* ---------- Detalhe ---------- */
function detalhe(t) {
  const venc = !pago(t) && t.venc < hoje();
  const grupo = t.grupo ? S.list.filter(x => x.grupo === t.grupo) : [];
  const f = folha(pago(t) ? (t.valor < 0 ? 'Pagamento realizado' : 'Recebimento realizado') : (t.valor < 0 ? 'Conta a pagar' : 'Conta a receber'), `
    <div class="resumo"><b>${esc(t.item)}</b><div class="big ${cls(t.valor)}">${brl2(t.valor)}</div>
      ${pago(t) ? `<span class="chip ok">Pago em ${dBR(t.pag)} · ${esc(t.conta || '')}</span>` : venc ? `<span class="chip bad">Vencida em ${dBR(t.venc)}</span>` : `<span class="chip warn">Vence em ${dBR(t.venc)}</span>`}</div>
    <dl class="dl">
      <dt>Destino</dt><dd>${esc(t.destino)}${t.aporteTipo ? ' · ' + esc(tipoAporteNome(t.aporteTipo)) : ''}${t.conversao ? ` · convertida em participação em ${dBR(t.conversao)}` : ''}${t.quitada ? ` · devolvida (não convertida), vencimento ${dBR(t.quitada)}` : ''}</dd><dt>Pertence a</dt><dd>${ehEmpresa(t) ? 'Empresa (administrativo)' : `${esc(obraT(t))} · quadra ${esc(t.quadra || '-')} · lote ${esc(t.lote || '-')}`}${variasEmpresas() ? ' · ' + esc(empPorId(empT(t))?.nome || '') : ''}</dd>
      <dt>Vencimento</dt><dd>${dBR(t.venc)}</dd>${t.juros ? `<dt>Juros / multa</dt><dd class="neg">${brl2(t.juros)}</dd>` : ''}
      ${t.forn ? `<dt>Fornecedor</dt><dd>${esc(t.forn)}</dd>` : ''}<dt>Nota fiscal</dt><dd>${t.nf ? (t.nfNum ? 'Nº ' + esc(t.nfNum) : 'Sim') : 'Não'}</dd>
      ${t.obs ? `<dt>Observação</dt><dd>${esc(t.obs)}</dd>` : ''}${grupo.length > 1 ? `<dt>Grupo</dt><dd>${grupo.length} lançamentos (${grupo.filter(pago).length} pagos)</dd>` : ''}
    </dl>
    <div id="anexoBox" class="anx"><div class="lbl">Anexos</div>
      <div class="anx-lista">${anexosDe(t).map((a, i) => `<div class="anx-row"><button type="button" class="btn sm" data-vax="${i}">${ic('clip', 14)} ${esc(tipoAnexoNome(a.tipo))}</button><span class="muted small">${a.link ? 'no Google Drive' : 'aguardando envio ao Drive'}</span><button type="button" class="x rw" data-rax="${i}" aria-label="Remover anexo">×</button></div>`).join('') || '<p class="muted small">Nenhum anexo.</p>'}</div>
      <div class="anx-add rw"><select id="tipoAx" aria-label="Tipo do anexo">${opts(TIPOS_ANEXO, pago(t) ? 'comprovante' : 'nota')}</select><label class="btn sm file">${ic('clip', 14)} Adicionar foto ou PDF<input type="file" id="addAnexo" accept="image/*,application/pdf" hidden></label></div></div>
    ${(t.hist || []).length ? `<details class="hist"><summary>Histórico de alterações</summary><ol>${t.hist.slice().reverse().map(h => `<li>${histTxt(h)}</li>`).join('')}</ol><button class="btn link sm" data-act="histT" data-id="${esc(t.id)}">Ver histórico completo</button></details>` : ''}
    <div class="acts wrap">
      ${pago(t) ? `<button class="btn" id="dEst">Estornar baixa</button>` : `<button class="btn primary" id="dBx">${t.valor < 0 ? 'Registrar pagamento' : 'Registrar recebimento'}</button>`}
      ${!pago(t) && t.valor > 0 ? `<button class="btn" data-act="pix" data-id="${esc(t.id)}">Cobrar por PIX</button>` : ''}${t.aporteTipo === 'conversivel' && t.valor > 0 && !t.conversao && !t.quitada ? `<button class="btn" data-act="converter" data-id="${esc(t.id)}">Converter em participação</button><button class="btn" data-act="quitarConv" data-id="${esc(t.id)}">Devolver com juros</button>` : ''}<button class="btn" id="dEd">Editar</button><button class="btn" id="dDup">Duplicar</button><button class="btn danger" id="dDel">Excluir</button>
      ${grupo.filter(x => !pago(x)).length > 1 ? `<button class="btn danger ghost" id="dDelG">Excluir ${grupo.filter(x => !pago(x)).length} em aberto do grupo</button>` : ''}
    </div>`);
  const on = (id, fn) => { const b = f.el.querySelector('#' + id); if (b) b.onclick = fn; };
  on('dBx', () => { f.fechar(); formBaixa([t.id]); });
  on('dEst', async () => { if (await confirmar('Estornar a baixa? O lançamento volta para "em aberto" e sai do extrato.', 'Estornar')) { const n = { ...t, status: 'aberto', juros: 0 }; delete n.pag; delete n.conta; salvarTitulos([n]); f.fechar(); toast('Baixa estornada'); } });
  on('dEd', () => { f.fechar(); formEditar(t); });
  on('dDup', () => { f.fechar(); formNovo({ tipo: t.valor < 0 ? 'saida' : 'entrada', item: t.item.replace(/\s*-\s*\S+\s+\S+$/, ''), forn: t.forn, destino: t.destino, valor: t.valor, quadra: t.quadra, lote: t.lote, venc: addMonths(t.venc, 1) }); });
  on('dDel', async () => { if (await confirmar(`Excluir <b>${esc(t.item)}</b>?`, 'Excluir')) { excluirTitulos([t.id]); f.fechar(); toast('Lançamento excluído'); } });
  on('dDelG', async () => { const ab = grupo.filter(x => !pago(x)); if (await confirmar(`Excluir os <b>${ab.length}</b> lançamentos em aberto deste grupo? Os já pagos ficam.`, 'Excluir')) { excluirTitulos(ab.map(x => x.id)); f.fechar(); toast(ab.length + ' lançamentos excluídos'); } });
  $$('[data-vax]', f.el).forEach(b => b.onclick = () => { const a = anexosDe(t)[+b.dataset.vax]; verAnexo(a.id, a.link); });
  $$('[data-rax]', f.el).forEach(b => b.onclick = async () => { const ax = anexosDe(t), a = ax[+b.dataset.rax]; if (await confirmar(`Tirar o anexo <b>${esc(tipoAnexoNome(a.tipo))}</b> deste lançamento? ${a.link ? 'O arquivo continua no seu Google Drive.' : ''}`, 'Remover')) { salvarTitulos([comAnexos(t, ax.filter((_, i) => i !== +b.dataset.rax))]); f.fechar(); toast('Anexo removido'); } });
  const inp = f.el.querySelector('#addAnexo');
  if (inp) inp.onchange = async () => { try { const aid = await salvarAnexo(inp.files[0]); salvarTitulos([comAnexos(t, [...anexosDe(t), { id: aid, tipo: f.el.querySelector('#tipoAx').value }])]); f.fechar(); toast('Anexo salvo'); } catch (e) { toast(e.message, 'err'); } };
}

/* ---------- Anexos ---------- */
const TIPOS_ANEXO = [['nota', 'Nota fiscal'], ['boleto', 'Boleto'], ['comprovante', 'Comprovante'], ['contrato', 'Contrato'], ['outro', 'Outro']];
const tipoAnexoNome = k => (TIPOS_ANEXO.find(x => x[0] === k) || ['', 'Anexo'])[1];
const anexosDe = t => Array.isArray(t.anexos) ? t.anexos : (t.anexo || t.anexoLink ? [{ id: t.anexo || '', tipo: '', link: t.anexoLink || '', drive: t.anexoDrive || '' }] : []);
const comAnexos = (t, lista) => { const n = { ...t, anexos: lista }; delete n.anexo; delete n.anexoLink; delete n.anexoDrive; return n; };
const tipoPorNome = (nome, padrao) => /boleto/i.test(nome) ? 'boleto' : /comprov|recibo|pix|ted/i.test(nome) ? 'comprovante' : /contrato/i.test(nome) ? 'contrato' : /nf|nota|danfe|nfs/i.test(nome) ? 'nota' : padrao;
/* Seletor de arquivos com tipo, usado no lançamento novo. */
function caixaAnexos(box, padrao) {
  const itens = [];
  const desenhar = () => {
    box.innerHTML = `<div class="lbl">Anexos <span class="muted">nota, boleto, comprovante…</span></div>
      <div class="anx-lista">${itens.map((x, i) => `<div class="anx-row"><span class="anx-nome">${ic('clip', 14)} ${esc(x.file.name || 'arquivo')}</span><select data-ati="${i}" aria-label="Tipo do anexo">${opts(TIPOS_ANEXO, x.tipo)}</select><button type="button" class="x" data-atx="${i}" aria-label="Remover anexo">×</button></div>`).join('')}</div>
      <label class="btn sm file">${ic('clip', 14)} ${itens.length ? 'Adicionar outro arquivo' : 'Adicionar foto ou PDF'}<input type="file" accept="image/*,application/pdf" multiple hidden></label>`;
    $$('[data-ati]', box).forEach(s => s.onchange = () => { itens[+s.dataset.ati].tipo = s.value; });
    $$('[data-atx]', box).forEach(b => b.onclick = () => { itens.splice(+b.dataset.atx, 1); desenhar(); });
    const inp = box.querySelector('input[type=file]'); inp.onchange = () => { for (const f of inp.files) itens.push({ file: f, tipo: tipoPorNome(f.name || '', typeof padrao === 'function' ? padrao() : padrao) }); desenhar(); };
  };
  desenhar();
  return { itens, adicionar(file, tipo) { itens.push({ file, tipo }); desenhar(); }, async salvar() { const out = []; for (const x of itens) out.push({ id: await salvarAnexo(x.file), tipo: x.tipo }); return out; } };
}
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
  setTimeout(() => sincronizarDrive(false), 1500);
  return id;
}
async function verAnexo(id, link) {
  if (!id && !link) return;
  if (link) { window.open(link, '_blank', 'noopener'); return; }
  try {
    const s = await getDoc(doc(db, 'anexos', id)); if (!s.exists()) { toast('Anexo ainda não disponível neste aparelho (precisa de internet).', 'err'); return; }
    const a = s.data(); if (a.link) { window.open(a.link, '_blank', 'noopener'); return; } const url = 'data:' + a.mime + ';base64,' + a.b64;
    folha('Anexo', a.mime.startsWith('image/') ? `<img src="${url}" class="anexo-img" alt="Anexo">` : `<p><a class="btn" href="${url}" download="${esc(a.nome || 'anexo.pdf')}">Abrir PDF</a></p>`, { largo: true });
  } catch (e) { toast('Não foi possível abrir o anexo: ' + (e.code || e.message), 'err'); }
}

/* =========================================================
   Google Drive: anexos ficam numa pasta da conta do dono
   ========================================================= */
const DRIVE_SCOPE = 'https://www.googleapis.com/auth/drive.file';
const CAL_SCOPE = 'https://www.googleapis.com/auth/calendar.app.created';
const DRV = { token: null, exp: 0, scope: '', carregando: null, enviando: false, erro: '' };
try { const s = JSON.parse(sessionStorage.getItem('sfgDrive') || 'null'); if (s && s.e > Date.now()) { DRV.token = s.t; DRV.exp = s.e; DRV.scope = s.s || DRIVE_SCOPE; } } catch (e) { /* sem sessão */ }
const driveConfigurado = () => !!firebaseConfig.driveClientId;
const agendaAtiva = () => !!(S.agenda && S.agenda.ativo);
const tokenTem = sc => !!DRV.token && Date.now() < DRV.exp - 60000 && (DRV.scope || '').split(' ').includes(sc);
const driveConectado = () => tokenTem(DRIVE_SCOPE);
function carregarGIS() {
  if (window.google && google.accounts && google.accounts.oauth2) return Promise.resolve();
  if (DRV.carregando) return DRV.carregando;
  DRV.carregando = new Promise((res, rej) => { const s = document.createElement('script'); s.src = 'https://accounts.google.com/gsi/client'; s.async = true; s.onload = res; s.onerror = () => { DRV.carregando = null; rej(new Error('Sem acesso ao Google agora. Verifique a internet.')); }; document.head.appendChild(s); });
  return DRV.carregando;
}
async function driveToken(interativo, precisa = DRIVE_SCOPE) {
  if (tokenTem(precisa)) return DRV.token;
  if (!interativo || !driveConfigurado()) return null;
  await carregarGIS();
  const escopos = [DRIVE_SCOPE, ...(agendaAtiva() || precisa === CAL_SCOPE ? [CAL_SCOPE] : []), ...((S.docsdrive && S.docsdrive.ler) || precisa === META_SCOPE ? [META_SCOPE] : [])].join(' ');
  return new Promise((res, rej) => {
    const c = google.accounts.oauth2.initTokenClient({
      client_id: firebaseConfig.driveClientId, scope: escopos, include_granted_scopes: true, hint: S.user?.email || '', prompt: '',
      callback: r => { if (r.error) { rej(new Error('O Google não autorizou: ' + r.error)); return; } DRV.token = r.access_token; DRV.exp = Date.now() + (+r.expires_in || 3600) * 1000; DRV.scope = r.scope || escopos; try { sessionStorage.setItem('sfgDrive', JSON.stringify({ t: DRV.token, e: DRV.exp, s: DRV.scope })); } catch (e) { /* ok */ } if (!tokenTem(precisa)) { rej(new Error(precisa === CAL_SCOPE ? 'A permissão para a Agenda não foi concedida. Marque a opção da Agenda na tela do Google.' : precisa === META_SCOPE ? 'A permissão para ver os arquivos do Drive não foi concedida.' : 'A permissão do Drive não foi concedida.')); return; } setTimeout(() => sincronizarAgenda(false), 500); res(DRV.token); },
      error_callback: e => rej(new Error(e && e.type === 'popup_closed' ? 'Autorização do Drive cancelada.' : 'Não consegui abrir a autorização do Google. Permita pop-ups para este site.'))
    });
    c.requestAccessToken();
  });
}
async function driveReq(url, opts = {}) {
  const tk = await driveToken(false); if (!tk) throw new Error('drive-desconectado');
  const r = await fetch(url, { ...opts, headers: { ...(opts.headers || {}), Authorization: 'Bearer ' + tk } });
  if (r.status === 401) { DRV.token = null; try { sessionStorage.removeItem('sfgDrive'); } catch (e) { /* ok */ } throw new Error('drive-desconectado'); }
  if (r.status === 404) return null;
  if (!r.ok) throw new Error('O Drive respondeu ' + r.status);
  return r.json();
}
const novaPasta = (nome, pai) => driveReq('https://www.googleapis.com/drive/v3/files?fields=id,webViewLink', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ name: nome, mimeType: 'application/vnd.google-apps.folder', parents: pai ? [pai] : [] }) });
async function pastaExiste(id) { if (!id) return false; const r = await driveReq(`https://www.googleapis.com/drive/v3/files/${id}?fields=id,trashed`); return !!(r && !r.trashed); }
async function pastaDoMes(ym) {
  const cfg = JSON.parse(JSON.stringify(S.drive || {})); cfg.anos = cfg.anos || {}; cfg.meses = cfg.meses || {};
  let mudou = false;
  if (!cfg.raiz || !(await pastaExiste(cfg.raiz.anexos))) {
    const sfg = await novaPasta('SFG Financeiro'); const an = await novaPasta('Anexos', sfg.id);
    cfg.raiz = { sfg: sfg.id, anexos: an.id, link: sfg.webViewLink }; cfg.anos = {}; cfg.meses = {}; mudou = true;
  }
  const ano = ym.slice(0, 4);
  if (!cfg.anos[ano] || (mudou ? false : !(await pastaExiste(cfg.anos[ano])))) { cfg.anos[ano] = (await novaPasta(ano, cfg.raiz.anexos)).id; cfg.meses = Object.fromEntries(Object.entries(cfg.meses).filter(([k]) => !k.startsWith(ano))); mudou = true; }
  if (!cfg.meses[ym]) { cfg.meses[ym] = (await novaPasta(ym.slice(5, 7) + ' ' + MESES[+ym.slice(5, 7) - 1], cfg.anos[ano])).id; mudou = true; }
  if (mudou) { S.drive = cfg; gravar(doc(db, 'config', 'drive'), cfg).catch(onErroEscrita); }
  return cfg.meses[ym];
}
async function enviarArquivoDrive(blob, nome, pasta) {
  const b = 'sfg' + uid();
  const corpo = new Blob([`--${b}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n${JSON.stringify({ name: nome, parents: [pasta] })}\r\n--${b}\r\nContent-Type: ${blob.type || 'application/octet-stream'}\r\n\r\n`, blob, `\r\n--${b}--`]);
  const r = await driveReq('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id,webViewLink', { method: 'POST', headers: { 'Content-Type': 'multipart/related; boundary=' + b }, body: corpo });
  if (!r || !r.id) throw new Error('O Drive não confirmou o envio.');
  return r;
}
function nomeAnexo(ts, mime, tipo) {
  const t = ts[0]; const lotes = ts.map(x => x.lote).filter(l => l && l !== '-');
  const onde = t.quadra && t.quadra !== '-' ? ` - ${t.quadra}${lotes.length ? ' ' + (lotes.length > 4 ? lotes.length + ' lotes' : lotes.join('-')) : ''}` : '';
  const base = `${t.venc || hoje()} ${ts.length > 1 ? baseItem(t) : t.item.replace(/\s*-\s*\S+\s+\S+$/, '')}${t.nfNum ? ' NF ' + t.nfNum : ''}${onde}`;
  return (base + (tipo ? ' (' + tipoAnexoNome(tipo).toLowerCase() + ')' : '')).replace(/[\/\\:*?"<>|#%]+/g, '-').replace(/\s+/g, ' ').trim().slice(0, 150) + (mime === 'application/pdf' ? '.pdf' : '.jpg');
}
/* Outros módulos com arquivos a enviar ao Drive (ex.: fotos do diário de obra). */
const DRIVE_EXTRA = [];
const anexosPendentes = () => { const m = new Map(); for (const t of S.list) for (const a of anexosDe(t)) if (a.id && !a.link) { const e = m.get(a.id) || m.set(a.id, { ids: [], tipo: a.tipo }).get(a.id); e.ids.push(t.id); } return m; };
async function sincronizarDrive(interativo) {
  if (!driveConfigurado() || DRV.enviando || somenteLeitura() || !navigator.onLine) return;
  const pend = anexosPendentes(); const extras = sum(DRIVE_EXTRA, x => x.pendentes()); if (!pend.size && !extras) return;
  let tk; try { tk = await driveToken(interativo); } catch (e) { if (interativo) toast(e.message, 'err'); return; }
  if (!tk) return;
  DRV.enviando = true; DRV.erro = ''; let ok = 0;
  try {
    const upd = new Map(); const cur = id => upd.get(id) || S.byId.get(id);
    const aplicar = (aid, ids, link, drive) => { const out = ids.map(cur).filter(Boolean).map(t => comAnexos(t, anexosDe(t).map(x => x.id === aid ? { ...x, link, drive } : x))); out.forEach(t => upd.set(t.id, t)); salvarTitulos(out, { sistema: true }); };
    for (const [aid, { ids, tipo }] of pend) {
      const ts = ids.map(cur).filter(Boolean); if (!ts.length) continue;
      const s = await getDoc(doc(db, 'anexos', aid)); if (!s.exists()) continue;
      const a = s.data();
      if (a.link) { aplicar(aid, ids, a.link, a.drive); ok++; continue; }
      if (!a.b64) continue;
      const bin = atob(a.b64); const u8 = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) u8[i] = bin.charCodeAt(i);
      const pasta = await pastaDoMes((ts[0].venc || hoje()).slice(0, 7));
      const r = await enviarArquivoDrive(new Blob([u8], { type: a.mime }), nomeAnexo(ts, a.mime, tipo), pasta);
      aplicar(aid, ids, r.webViewLink, r.id);
      gravar(doc(db, 'anexos', aid), { drive: r.id, link: r.webViewLink, mime: a.mime, nome: a.nome || '', enviado: Date.now() }).catch(onErroEscrita);
      ok++; refreshView();
    }
    for (const x of DRIVE_EXTRA) ok += await x.enviar();
    if (ok) toast(`${ok} anexo(s) enviados ao Google Drive`);
  } catch (e) { DRV.erro = e.message === 'drive-desconectado' ? 'A autorização do Drive expirou. Toque em Conectar.' : e.message; if (interativo) toast(DRV.erro, 'err'); }
  finally { DRV.enviando = false; refreshView(); }
}
addEventListener('online', () => setTimeout(() => sincronizarDrive(false), 3000));
ACOES.drive = () => sincronizarDrive(true).then(() => { if (!anexosPendentes().size && driveConectado()) toast('Drive conectado. Nenhum anexo pendente.'); });
/* ---------- Avisos de vencimento pela Google Agenda ---------- */
const AG = { enviando: false, erro: '' };
const TZ = (() => { try { return Intl.DateTimeFormat().resolvedOptions().timeZone || 'America/Recife'; } catch (e) { return 'America/Recife'; } })();
async function calReq(url, opts = {}) {
  const tk = await driveToken(false, CAL_SCOPE); if (!tk) throw new Error('agenda-desconectada');
  const r = await fetch('https://www.googleapis.com/calendar/v3/' + url, { method: opts.method || 'GET', headers: { Authorization: 'Bearer ' + tk, ...(opts.json ? { 'Content-Type': 'application/json' } : {}) }, body: opts.json ? JSON.stringify(opts.json) : undefined });
  if (r.status === 401) { DRV.token = null; try { sessionStorage.removeItem('sfgDrive'); } catch (e) { /* ok */ } throw new Error('agenda-desconectada'); }
  if (r.status === 404 || r.status === 410) return null;
  if (r.status === 204) return {};
  if (!r.ok) { let m = ''; try { m = JSON.stringify(await r.json()); } catch (e) { /* ok */ } if (r.status === 403 && /accessNotConfigured|SERVICE_DISABLED|has not been used/i.test(m)) throw new Error('api-desativada'); throw new Error('A Agenda respondeu ' + r.status); }
  return r.json();
}
function eventosAgenda() {
  const h = hoje(), lim = addDays(h, 60), dias = new Map();
  for (const t of S.list) { if (pago(t) || !t.venc) continue; const d = t.venc < h ? h : t.venc; if (d > lim) continue; const g = dias.get(d) || dias.set(d, { pag: [], rec: [], venc: [] }).get(d); (t.venc < h ? g.venc : t.valor < 0 ? g.pag : g.rec).push(t); }
  const cfg = S.agenda || {}; const hora = cfg.hora || '08:00';
  const lin = es => es.map(e => { const x = entTxt(e); return `• ${x[1]}${x[2] && x[2] !== '-' ? ' (' + x[2] + ')' : ''}: ${brl2(x[3])}`; }).join('\n');
  const out = new Map();
  for (const [d, g] of dias) {
    const tot = a => Math.abs(sum(a, t => t.valor)); const partes = [];
    if (g.venc.length) partes.push(`${g.venc.length} vencida(s) ${brl(tot(g.venc))}`);
    if (g.pag.length) partes.push(`pagar ${brl(tot(g.pag))} (${g.pag.length})`);
    if (g.rec.length) partes.push(`receber ${brl(tot(g.rec))} (${g.rec.length})`);
    const desc = [g.venc.length ? 'VENCIDAS EM ABERTO\n' + lin(agrupar(g.venc, 'venc')) : '', g.pag.length ? 'A PAGAR\n' + lin(agrupar(g.pag, 'venc')) : '', g.rec.length ? 'A RECEBER\n' + lin(agrupar(g.rec, 'venc')) : '', 'Abrir o SFG Financeiro: ' + location.origin + location.pathname + '#/pagar'].filter(Boolean).join('\n\n');
    const [hh, mm] = hora.split(':').map(Number); const fimH = `${String(hh).padStart(2, '0')}:${String(mm + 15).padStart(2, '0')}`;
    const ev = { summary: 'SFG: ' + partes.join(' · ').replace(/ /g, ' '), description: desc.replace(/ /g, ' '), start: { dateTime: `${d}T${hora}:00`, timeZone: TZ }, end: { dateTime: `${d}T${fimH}:00`, timeZone: TZ },
      reminders: { useDefault: false, overrides: [{ method: 'popup', minutes: 0 }, ...(cfg.vespera !== false ? [{ method: 'popup', minutes: 24 * 60 }] : [])] }, colorId: g.venc.length ? '11' : '9', transparency: 'transparent' };
    out.set(d, ev);
  }
  return out;
}
async function sincronizarAgenda(interativo) {
  if (!driveConfigurado() || !agendaAtiva() || AG.enviando || somenteLeitura() || !navigator.onLine) return;
  let tk; try { tk = await driveToken(interativo, CAL_SCOPE); } catch (e) { AG.erro = e.message; if (interativo) toast(e.message, 'err'); refreshView(); return; }
  if (!tk) return;
  AG.enviando = true; AG.erro = ''; let n = 0;
  const cfg = clone(S.agenda || {}); cfg.eventos = cfg.eventos || {};
  try {
    if (!cfg.calId || !(await calReq('calendars/' + encodeURIComponent(cfg.calId)))) {
      const c = await calReq('calendars', { method: 'POST', json: { summary: 'SFG Vencimentos', description: 'Vencimentos do SFG Financeiro. Criado e atualizado pelo app.', timeZone: TZ } });
      cfg.calId = c.id; cfg.eventos = {};
    }
    const base = 'calendars/' + encodeURIComponent(cfg.calId) + '/events';
    const evs = eventosAgenda(), h = hoje();
    for (const [d, ev] of evs) {
      const hs = String(hashTxt(JSON.stringify(ev))); const at = cfg.eventos[d];
      if (at && at.h === hs) continue;
      let r = at ? await calReq(base + '/' + encodeURIComponent(at.id), { method: 'PUT', json: ev }) : null;
      if (!r) r = await calReq(base, { method: 'POST', json: ev });
      cfg.eventos[d] = { id: r.id, h: hs }; n++;
    }
    for (const d of Object.keys(cfg.eventos)) {
      if (evs.has(d)) continue;
      if (d >= h) { await calReq(base + '/' + encodeURIComponent(cfg.eventos[d].id), { method: 'DELETE' }); n++; }
      delete cfg.eventos[d];
    }
    cfg.ultima = Date.now(); S.agenda = cfg; gravar(doc(db, 'config', 'agenda'), cfg).catch(onErroEscrita);
    if (interativo) toast(n ? `Agenda atualizada (${n} dia(s) alterado(s))` : 'Agenda já estava em dia');
  } catch (e) {
    AG.erro = e.message === 'api-desativada' ? 'A Google Calendar API está desativada no projeto do Google Cloud. Ative em console.cloud.google.com/apis/library/calendar-json.googleapis.com (mesmo projeto do Drive) e tente de novo.' : e.message === 'agenda-desconectada' ? 'A autorização do Google expirou. Toque em Atualizar agenda.' : e.message;
    if (cfg.calId && cfg.calId !== (S.agenda || {}).calId) { S.agenda = { ...(S.agenda || {}), calId: cfg.calId, eventos: cfg.eventos }; gravar(doc(db, 'config', 'agenda'), S.agenda).catch(onErroEscrita); }
    if (interativo) toast(AG.erro, 'err');
  } finally { AG.enviando = false; refreshView(); }
}
let _agT = 0; const agendarAgenda = () => { clearTimeout(_agT); _agT = setTimeout(() => { if (agendaAtiva() && tokenTem(CAL_SCOPE)) sincronizarAgenda(false); }, 15000); };
ACOES.agAtivar = () => { S.agenda = { ...(S.agenda || {}), ativo: true, hora: S.agenda?.hora || '08:00', vespera: S.agenda?.vespera !== false }; gravar(doc(db, 'config', 'agenda'), S.agenda).catch(onErroEscrita); sincronizarAgenda(true); };
ACOES.agSync = () => sincronizarAgenda(true);
ACOES.agDesl = async () => { if (await confirmar('Parar de atualizar a agenda "SFG Vencimentos"? Os eventos que já estão lá continuam; você pode apagar a agenda no Google Agenda.', 'Desligar', false)) { S.agenda = { ...(S.agenda || {}), ativo: false }; gravar(doc(db, 'config', 'agenda'), S.agenda).catch(onErroEscrita); refreshView(); } };
ACOES.agHora = () => {
  const c = S.agenda || {};
  const f = folha('Horário dos avisos', `<form class="form" id="fx"><label>Avisar no dia do vencimento às<select name="hora">${opts(['06:00', '07:00', '08:00', '09:00', '10:00', '12:00', '14:00', '18:00'], c.hora || '08:00')}</select></label><label class="chk"><input type="checkbox" name="vespera" ${c.vespera !== false ? 'checked' : ''}> Avisar também na véspera, no mesmo horário</label><div class="acts"><button type="button" class="btn" data-close>Cancelar</button><button class="btn primary">Salvar</button></div></form>`);
  f.el.querySelector('#fx').onsubmit = e => { e.preventDefault(); S.agenda = { ...c, hora: e.target.hora.value, vespera: e.target.vespera.checked }; gravar(doc(db, 'config', 'agenda'), S.agenda).catch(onErroEscrita); f.fechar(); sincronizarAgenda(true); };
};
function cardAgenda() {
  if (!driveConfigurado()) return '';
  const c = S.agenda || {}; const at = agendaAtiva();
  return `<div class="card"><h2>Avisos de vencimento no celular</h2>
    <p class="muted small">O app cria no seu Google Agenda um calendário chamado <b>SFG Vencimentos</b>, com um evento em cada dia que tem conta a pagar ou a receber (próximos 60 dias) e as contas vencidas no dia de hoje. A notificação chega pelo Google Agenda no celular, mesmo com o SFG fechado.</p>
    ${at ? `<div class="kv"><span>Situação</span><b>${AG.enviando ? 'Atualizando…' : 'Ligado'}</b></div><div class="kv"><span>Aviso</span><b>${esc(c.hora || '08:00')}${c.vespera !== false ? ', no dia e na véspera' : ', no dia'}</b><button class="btn sm rw" data-act="agHora">Alterar</button></div><div class="kv"><span>Última atualização</span><b>${c.ultima ? quando(c.ultima) : 'ainda não'}</b></div>` : ''}
    ${AG.erro ? `<p class="erro small">${esc(AG.erro)}</p>` : ''}
    <div class="acts left rw">${at ? `<button class="btn primary" data-act="agSync">Atualizar agenda agora</button><button class="btn" data-act="agDesl">Desligar</button>` : `<button class="btn primary" data-act="agAtivar">Ligar avisos na Google Agenda</button>`}</div>
    <p class="muted small">No iPhone: em Ajustes, Calendário, Contas, adicione a conta Google ${esc(S.user?.email || '')} e ligue Calendários (ou use o app Google Agenda). Depois de ligado, a agenda se atualiza sozinha sempre que o app estiver conectado ao Google; se ficar mais de um dia sem atualizar, o Início mostra um lembrete.</p></div>`;
}
function cardDrive() {
  if (!driveConfigurado()) return `<div class="card"><h2>Google Drive</h2><p class="muted">Os anexos vão para uma pasta "SFG Financeiro" no seu Google Drive assim que a credencial do Google for configurada.</p></div>`;
  const pend = anexosPendentes().size; const lk = S.drive && S.drive.raiz && S.drive.raiz.link;
  return `<div class="card"><h2>Google Drive</h2>
    <div class="kv"><span>Situação</span><b>${DRV.enviando ? 'Enviando…' : driveConectado() ? 'Conectado nesta sessão' : 'Desconectado'}</b></div>
    <div class="kv"><span>Aguardando envio</span><b>${pend} anexo(s)</b></div>
    ${DRV.erro ? `<p class="erro small">${esc(DRV.erro)}</p>` : ''}
    <p class="muted small">Fotos e PDFs ficam primeiro no app (funciona sem internet) e depois vão para a pasta SFG Financeiro / Anexos, separada por ano e mês. O app só enxerga os arquivos que ele mesmo criou no seu Drive.</p>
    <div class="acts left"><button class="btn primary rw" data-act="drive">${driveConectado() ? 'Enviar pendentes agora' : 'Conectar e enviar'}</button>${lk ? `<a class="btn" href="${esc(lk)}" target="_blank" rel="noopener">Abrir pasta no Drive</a>` : ''}</div></div>`;
}

/* =========================================================
   Componentes de lista
   ========================================================= */
const listaSel = new Set();
function linhaTitulo(t, o = {}) {
  const h = hoje(); const venc = !pago(t) && t.venc < h; const hj = !pago(t) && t.venc === h;
  const d = o.data === 'pag' ? (t.pag || t.venc) : t.venc;
  const meta = `<span class="loc">${esc(ehEmpresa(t) ? 'Empresa' : t.quadra && t.quadra !== '-' ? t.quadra + (t.lote && t.lote !== '-' ? ' ' + t.lote : '') : 'Geral da obra')}</span>` + [t.destino, t.forn, t.nf ? (t.nfNum ? 'NF ' + t.nfNum : 'com NF') : '', anexosDe(t).length ? (anexosDe(t).length > 1 ? anexosDe(t).length + ' anexos' : 'anexo') : '', t.rec ? 'recorrente' : ''].filter(Boolean).map(x => `<span>${esc(x)}</span>`).join('') + (variasEmpresas() ? (empDiverge(t) ? `<span class="chip warn" title="CNPJ diferente do CNPJ da obra">${esc(empPorId(t.emp).nome)} ≠ obra</span>` : !empDef(t) ? '<span class="chip">sem CNPJ</span>' : '') : '');
  return `<div class="tr ${o.sel ? 'selmode' : ''} ${listaSel.has(t.id) ? 'sel' : ''}" data-act="${o.sel ? 'tsel' : 'abrir'}" data-id="${t.id}">
    ${o.sel ? `<span class="cb">${listaSel.has(t.id) ? '✓' : ''}</span>` : ''}
    ${o.semData ? '' : `<div class="td-date"><b>${d ? d.slice(8, 10) + '/' + d.slice(5, 7) : '—'}</b><small>${d ? d.slice(0, 4) : ''}</small></div>`}
    <div class="td-main"><div class="td-item">${esc(t.item)}</div><div class="td-meta">${meta}</div></div>
    <div class="td-val"><b class="${cls(t.valor)}">${brl2(t.valor)}</b>${t.juros ? `<small class="neg">juros ${brl2(t.juros)}</small>` : ''}${o.saldo !== undefined ? `<small class="muted">saldo ${brl(o.saldo)}</small>` : ''}
      ${o.semData ? '' : !pago(t) ? (venc ? '<span class="chip bad">vencida</span>' : hj ? '<span class="chip warn">hoje</span>' : '') : (o.mostrarPago ? '<span class="chip ok">pago</span>' : '')}</div>
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
    ${o.semData ? '' : `<div class="td-date"><b>${d ? d.slice(8, 10) + '/' + d.slice(5, 7) : '—'}</b><small>${d ? d.slice(0, 4) : ''}</small></div>`}
    <div class="td-main"><div class="td-item">${esc(baseItem(t))} <span class="chip">${g.itens.length} lotes</span></div><div class="td-meta"><span class="loc">${esc(t.quadra)}</span><span>lotes ${esc(ls.length > 8 ? ls.slice(0, 8).join(', ') + '…' : ls.join(', '))}</span><span>${esc(t.destino)}</span></div></div>
    <div class="td-val"><b class="${cls(tot)}">${brl2(tot)}</b>${jur ? `<small class="neg">juros ${brl2(jur)}</small>` : ''}<small class="muted">${g.itens.length} × ${brl2(t.valor)}</small>
      ${o.semData ? '' : !pago(t) ? (venc ? '<span class="chip bad">vencida</span>' : hj ? '<span class="chip warn">hoje</span>' : '') : (o.mostrarPago ? '<span class="chip ok">pago</span>' : '')}</div>
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
function grafico(id, cfg) { if (!Chart) { carregarLib('chart').then(() => grafico(id, cfg)).catch(e => toast(e.message, 'err')); return; } if (charts[id]) { charts[id].destroy(); delete charts[id]; } const c = document.getElementById(id); if (c) charts[id] = new Chart(c, cfg); }
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

/* ---------- Linha d'água do caixa (assinatura do Início) ---------- */
function linhaDagua(pj, horizonte) {
  const h = hoje(), fim = addDays(h, horizonte);
  const ds = pj.dias.filter(x => x.d <= fim); if (!ds.length) return '';
  const idx = d => Math.max(0, Math.round((new Date(d + 'T12:00') - new Date(h + 'T12:00')) / 864e5));
  const vals = ds.map(x => x.v); const vmin = Math.min(0, ...vals), vmax = Math.max(0, ...vals);
  const pad = (vmax - vmin) * 0.1 || 1; const lo = vmin - pad, hi = vmax + pad;
  const X = i => i / horizonte * 1000, Y = v => 200 - (v - lo) / (hi - lo) * 200, y0 = Y(0);
  let d = `M0,${Y(ds[0].v).toFixed(1)}`;
  for (let k = 1; k < ds.length; k++) d += ` H${X(idx(ds[k].d)).toFixed(1)} V${Y(ds[k].v).toFixed(1)}`;
  d += ' H1000';
  const area = d + ` V${y0.toFixed(1)} H0 Z`;
  let mn = ds[0]; for (const x of ds) if (x.v < mn.v) mn = x;
  const mx = X(idx(mn.d)) / 10, my = Y(mn.v) / 2;
  const meses = []; let m = addMonths(h.slice(0, 7) + '-01', 1);
  while (m <= fim) { if (X(idx(m)) / 10 > 7) meses.push(`<span class="agua-mes" style="left:${X(idx(m)) / 10}%">${MES3[+m.slice(5, 7) - 1]}</span>`); m = addMonths(m, 1); }
  return `<div class="agua" role="img" aria-label="Saldo projetado dos próximos ${horizonte} dias. Menor saldo ${brl(mn.v)} em ${dBR(mn.d)}.">
    <svg viewBox="0 0 1000 200" preserveAspectRatio="none" aria-hidden="true">
      <defs><clipPath id="aguaCima"><rect x="0" y="0" width="1000" height="${y0.toFixed(1)}"/></clipPath><clipPath id="aguaBaixo"><rect x="0" y="${y0.toFixed(1)}" width="1000" height="${(200 - y0).toFixed(1)}"/></clipPath></defs>
      <path d="${area}" class="agua-pos" clip-path="url(#aguaCima)"/><path d="${area}" class="agua-neg" clip-path="url(#aguaBaixo)"/>
      <line x1="0" x2="1000" y1="${y0.toFixed(1)}" y2="${y0.toFixed(1)}" class="agua-zero" vector-effect="non-scaling-stroke"/>
      <path d="${d}" class="agua-linha" vector-effect="non-scaling-stroke"/>
    </svg>
    <span class="agua-zero-lab" style="top:${y0 / 2}%">zero</span>
    <span class="agua-ponto" style="left:${mx}%;top:${my}%"></span>
    <span class="agua-nota ${mx > 52 ? 'esq' : ''}" style="left:${mx}%;top:${my}%">Menor saldo <b>${brl(mn.v)}</b> em ${dBR(mn.d).slice(0, 5)}</span>
    <div class="agua-meses"><span class="agua-mes" style="left:0">hoje</span>${meses.join('')}</div>
  </div>`;
}

VIEWS.inicio = {
  titulo: 'Início',
  refresh(el) {
    const h = hoje(), L = S.list, ab = L.filter(t => !pago(t));
    if (!L.length && !S.estrutura) { el.innerHTML = `<div class="vazio"><img src="${LOGO_COMPLETO}" alt="" width="180"><h2>Comece importando o histórico</h2><p>Traga a planilha SFG Projeto para o sistema e todos os lançamentos, quadras e o orçamento aparecem aqui.</p><a class="btn primary" href="#/dados">Importar planilha</a></div>`; return; }
    const sd = saldos(), pj = projecao();
    const venc = ab.filter(t => t.valor < 0 && t.venc < h), p7 = ab.filter(t => t.valor < 0 && t.venc >= h && t.venc <= addDays(h, 7)), r30 = ab.filter(t => t.valor > 0 && t.venc <= addDays(h, 30));
    const neg = pj.dias.find(x => x.v < 0);
    const dataLonga = new Date().toLocaleDateString('pt-BR', { weekday: 'long', day: 'numeric', month: 'long' });
    const prox = ab.filter(t => t.venc <= addDays(h, 10)).sort((a, b) => a.venc.localeCompare(b.venc) || a.valor - b.valor);
    const dias = []; for (const t of prox) { const k = t.venc < h ? 'vencidas' : t.venc; let g = dias.find(x => x.k === k); if (!g) { g = { k, ts: [] }; dias.push(g); } g.ts.push(t); }
    const nomeDia = k => k === 'vencidas' ? 'Vencidas' : k === h ? 'Hoje' : k === addDays(h, 1) ? 'Amanhã' : new Date(k + 'T12:00').toLocaleDateString('pt-BR', { weekday: 'long', day: 'numeric', month: 'short' });
    let restam = 12;
    const agenda = dias.map(g => { if (restam <= 0) return ''; const es = agrupar(g.ts, 'venc').slice(0, restam); restam -= es.length; const tot = sum(g.ts, t => t.valor);
      return `<div class="ag-dia"><h3 class="${g.k === 'vencidas' ? 'neg' : ''}"><span>${nomeDia(g.k)}</span><span class="fig ${cls(tot)}">${brl(tot)}</span></h3>${es.map(e => linhaEntrada(e, { baixa: true, semData: true })).join('')}</div>`; }).join('');
    const qs = E().quadras.map(resumoQuadra).sort((a, b) => b.proj - a.proj); const qmax = Math.max(1, ...qs.map(q => Math.abs(q.proj)));
    const agVelha = agendaAtiva() && !somenteLeitura() && (!S.agenda.ultima || Date.now() - S.agenda.ultima > 864e5);
    el.innerHTML = `${painelAlertas()}<section class="extrato-hero">
      <div class="eh-topo">
        <div class="eh-saldo"><p class="eh-lab">Saldo em caixa, ${esc(dataLonga)}</p><p class="eh-val fig ${sd.total < 0 ? 'neg' : ''}">${brl(sd.total)}</p>
          <p class="eh-contas">${sd.contas.map(c => `<span>${esc(c.nome)} <b class="fig">${brl(c.saldo)}</b></span>`).join('')}</p></div>
        <div class="eh-fatos">
          <a href="#/pagar"><span class="k">Vencido a pagar</span><span class="v fig ${venc.length ? 'neg' : ''}">${brl(sum(venc, t => t.valor))}</span><small>${venc.length ? venc.length + ' conta(s) em atraso' : 'nada em atraso'}</small></a>
          <a href="#/pagar"><span class="k">Sai nos próximos 7 dias</span><span class="v fig">${brl(sum(p7, t => t.valor))}</span><small>${p7.length} lançamento(s)</small></a>
          <a href="#/receber"><span class="k">Entra nos próximos 30 dias</span><span class="v fig pos">${brl(sum(r30, t => t.valor))}</span><small>${r30.length} recebimento(s)</small></a>
        </div>
      </div>
      <div class="eh-agua"><div class="eh-agua-cab"><h2>Caixa nos próximos 120 dias</h2>${neg ? `<p class="aviso">Fica negativo em ${dBR(neg.d).slice(0, 5)} e chega a <b class="fig">${brl(pj.min.v)}</b> em ${dBR(pj.min.d).slice(0, 5)}. Planeje um aporte ou antecipe recebíveis.</p>` : '<p class="muted">O saldo projetado não fica negativo nesse período.</p>'}<a class="btn link" href="#/fluxo">Abrir fluxo de caixa</a></div>${linhaDagua(pj, 120)}</div>
    </section>
    <div class="dupla">
      <section class="painel"><div class="painel-cab"><h2>Vencimentos</h2><button class="btn sm" data-act="ics">${ic('calendario', 16)} Lembretes no calendário</button></div>${agenda || '<p class="muted pad">Nada vencendo nos próximos dias.</p>'}<a class="btn link" href="#/pagar">Ver todas as contas a pagar</a></section>
      <section class="painel"><div class="painel-cab"><h2>Resultado projetado por quadra</h2></div>${qs.map(q => `<a class="qbar" href="#/obras/${encodeURIComponent(q.q)}"><b>${esc(q.q)}</b><span class="bar"><i style="width:${Math.max(0, q.proj) / qmax * 100}%"></i></span><span class="fig ${cls(q.proj)}">${short(q.proj)}</span></a>`).join('')}
        <p class="painel-rodape">Total projetado <b class="fig">${short(sum(qs, q => q.proj))}</b>, previsto ${short(sum(qs, q => q.prevRes))}</p></section>
    </div>`;
  }
};

/* Filtros dependem dos cadastros: se eles chegam ou mudam depois de a tela abrir, a barra de filtros é refeita. */
const sigLista = () => JSON.stringify([variasEmpresas(), E().obras, E().quadras, E().contas.map(c => c.nome), todosDestinos(), S.list.some(ehEmpresa)]);
function viewLista(modo) {
  const F = { q: '', per: modo === 'extrato' ? 'mes' : 'todos', st: 'aberto', emp: '', obra: '', quadra: '', dest: '', locs: [], dests: [], jur: 'com', conta: '', ini: '', fim: '' };
  let selMode = false, limite = 200, agrup = true;
  const per = modo === 'extrato'
    ? [['mes', 'Este mês'], ['mesant', 'Mês passado'], ['30d', 'Últimos 30 dias'], ['ano', 'Este ano'], ['todos', 'Tudo'], ['custom', 'Período…']]
    : [['todos', 'Todas as datas'], ['vencidas', 'Vencidas'], ['hoje', 'Vencem hoje'], ['7', 'Próximos 7 dias'], ['30', 'Próximos 30 dias'], ['mes', 'Este mês'], ['proxmes', 'Próximo mês'], ['custom', 'Período…']];
  return {
    titulo: modo === 'pagar' ? 'Contas a pagar' : modo === 'receber' ? 'Contas a receber' : 'Extrato',
    montar(el) {
      listaSel.clear(); selMode = false; limite = 200; el._sigLista = sigLista();
      if (filtroPendente) { Object.assign(F, filtroPendente); filtroPendente = null; }
      if (F.quadra) { F.locs = ['Q:' + F.quadra]; F.quadra = ''; } if (F.dest) { F.dests = [F.dest]; F.dest = ''; }
      el.innerHTML = `<div class="filtros card">
        <input type="search" id="fq" placeholder="Buscar descrição, fornecedor…" value="${esc(F.q)}">
        ${modo !== 'extrato' ? `<select id="fst">${opts([['aberto', 'Em aberto'], ['pago', modo === 'pagar' ? 'Pagas' : 'Recebidas'], ['todos', 'Todas']], F.st)}</select>` : `<select id="fconta">${opts(E().contas.map(c => c.nome), F.conta, 'Todas as contas')}</select>`}
        <select id="fper">${opts(per, F.per)}</select>
        <span id="fcustom" class="${F.per === 'custom' ? '' : 'hidden'}"><input type="date" id="fini" value="${F.ini}"> até <input type="date" id="ffim" value="${F.fim}"></span>
        ${variasEmpresas() ? `<select id="femp">${opts([...empOpts(), ['__sem', 'Sem CNPJ definido'], ['__div', 'CNPJ diferente da obra']], F.emp, 'Todas as empresas')}</select>` : ''}
        ${variasObras() || S.list.some(ehEmpresa) ? `<select id="fobra">${opts([...E().obras, ['__emp', 'Empresa (fora das obras)']], F.obra, 'Todas as obras')}</select>` : ''}
        ${multiSelLocais('flocs', F.locs, F.obra)}
        ${multiSel('fdests', 'Todos os destinos', ['op', 'nonop', 'rec', 'aporte'].map(s => [SECAO_NOME[s], (E().destinos[s] || []).map(d => [d, d])]).filter(g => g[1].length), F.dests)}
        <select id="fjur" title="Como somar juros e multas">${opts([['com', 'Valores com juros'], ['sem', 'Sem juros (só o principal)'], ['so', 'Só juros e multas']], F.jur)}</select>
        <label class="tgl"><input type="checkbox" id="fgrp" ${agrup ? 'checked' : ''}> Agrupar lotes</label>
        <button class="btn" id="bSel">Selecionar</button>
      </div><div id="lsaldo"></div><div class="card"><div id="ltot" class="ltot"></div><div id="lrows" class="tlist"></div><div id="lmore"></div></div><div id="selbar" class="selbar hidden"></div>`;
      const bind = (id, k, ev = 'change') => $('#' + id, el)?.addEventListener(ev, e => { F[k] = e.target.value; limite = 200; if (k === 'per') $('#fcustom', el).classList.toggle('hidden', F.per !== 'custom'); this.refresh(el); });
      bind('fq', 'q', 'input'); bind('fst', 'st'); bind('fconta', 'conta'); bind('fper', 'per'); bind('fini', 'ini'); bind('ffim', 'fim'); bind('femp', 'emp'); bind('fobra', 'obra'); bind('fjur', 'jur'); ligarMultiSel(el, 'flocs', v => { F.locs = v; limite = 200; this.refresh(el); }); ligarMultiSel(el, 'fdests', v => { F.dests = v; limite = 200; this.refresh(el); });
      $('#fgrp', el).onchange = e => { agrup = e.target.checked; limite = 200; this.refresh(el); };
      const bs = $('#bSel', el); if (bs) bs.onclick = () => { selMode = !selMode; listaSel.clear(); bs.classList.toggle('on', selMode); bs.textContent = selMode ? 'Cancelar seleção' : 'Selecionar'; this.refresh(el); };
    },
    sair() { listaSel.clear(); },
    refresh(el) {
      if (el._sigLista !== undefined && el._sigLista !== sigLista()) { this.montar(el); }
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
        if (F.emp === '__sem') { if (empDef(t)) return false; } else if (F.emp === '__div') { if (!empDiverge(t)) return false; } else if (F.emp && empT(t) !== F.emp) return false; if (F.obra === '__emp') { if (!ehEmpresa(t)) return false; } else if (F.obra && obraT(t) !== F.obra) return false; if (F.locs.length && !casaLocais(t, F.locs)) return false; if (F.dests.length && !F.dests.includes(t.destino)) return false;
        if (F.conta && t.conta !== F.conta) return false;
        if (F.jur === 'so' && !(pago(t) && t.juros)) return false;
        if (q && !(t.item.toLowerCase().includes(q) || (t.forn || '').toLowerCase().includes(q) || (t.obs || '').toLowerCase().includes(q) || (t.nfNum || '').includes(q))) return false;
        return okPer(campo === 'pag' ? (t.pag || t.venc) : t.venc);
      });
      const desc = campo === 'pag';
      rows.sort((a, b) => { const x = (campo === 'pag' ? (a.pag || a.venc) : a.venc) || '', y = (campo === 'pag' ? (b.pag || b.venc) : b.venc) || ''; return desc ? y.localeCompare(x) || (b.criado || 0) - (a.criado || 0) : x.localeCompare(y) || (a.criado || 0) - (b.criado || 0); });
      // saldo corrido no extrato
      let saldoMap = null;
      if (modo === 'extrato' && !q && !F.emp && !F.obra && !F.locs.length && !F.dests.length) {
        const sd = saldos(); let s = F.conta ? (E().contas.find(c => c.nome === F.conta)?.inicial || 0) : sum(E().contas, c => +c.inicial || 0);
        const all = S.list.filter(t => pago(t) && (!F.conta || t.conta === F.conta)).sort((a, b) => (a.pag || a.venc || '').localeCompare(b.pag || b.venc || '') || (a.criado || 0) - (b.criado || 0));
        saldoMap = new Map(); for (const t of all) { s += t.valor + (t.juros || 0); saldoMap.set(t.id, s); }
        $('#lsaldo', el).innerHTML = `<div class="grid kpis">${kpi('Saldo total', brl(sd.total), 'todas as contas', cls(sd.total))}${sd.contas.map(c => kpi(esc(c.nome), brl(c.saldo), c.inicial ? 'saldo inicial ' + brl(c.inicial) : '', cls(c.saldo))).join('')}</div>`;
      } else if (modo === 'extrato') $('#lsaldo', el).innerHTML = '';
      const jurT = x => pago(x) ? (x.juros || 0) : 0;
      const valT = x => F.jur === 'sem' ? x.valor : F.jur === 'so' ? jurT(x) : x.valor + jurT(x);
      const tot = sum(rows, valT), jurTot = sum(rows, jurT);
      const ent = sum(rows, t => F.jur === 'so' ? Math.max(0, jurT(t)) : Math.max(0, t.valor)), sai = sum(rows, t => F.jur === 'so' ? Math.min(0, jurT(t)) : Math.min(0, t.valor) + (F.jur === 'sem' ? 0 : Math.min(0, jurT(t))));
      $('#ltot', el).innerHTML = `<span><b>${rows.length.toLocaleString('pt-BR')}</b> lançamento(s)</span>${modo === 'extrato' ? `<span>Entradas <b class="pos">${brl(ent)}</b></span><span>Saídas <b class="neg">${brl(sai)}</b></span>` : ''}<span>Total ${F.jur === 'sem' ? 'sem juros' : F.jur === 'so' ? 'de juros e multas' : ''} <b class="${cls(tot)}">${brl2(tot)}</b></span>${F.jur === 'com' && Math.abs(jurTot) > 0.004 ? `<span class="muted">principal ${brl2(tot - jurTot)} + juros/multas <b class="neg">${brl2(jurTot)}</b></span>` : F.jur === 'sem' && Math.abs(jurTot) > 0.004 ? `<span class="muted">juros/multas fora do total: ${brl2(jurTot)}</span>` : ''}${selMode ? `<button class="btn link sm" id="bSelAll">${modo === 'extrato' || F.st === 'pago' ? 'Selecionar todos' : 'Selecionar todos em aberto'}</button>` : ''}`;
      const sa = $('#bSelAll', el); if (sa) sa.onclick = () => { rows.filter(t => modo === 'extrato' || F.st === 'pago' || !pago(t)).forEach(t => listaSel.add(t.id)); this.refresh(el); };
      let html = '', mesAtual = null;
      const dOf = x => ((campo === 'pag' ? (x.pag || x.venc) : x.venc) || '');
      const subMes = {}; for (const x of rows) { const m = dOf(x).slice(0, 7); subMes[m] = (subMes[m] || 0) + valT(x); }
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
        sb.innerHTML = `<span><b>${sel.length}</b> selecionado(s) · <b class="${cls(sum(sel, t => t.valor))}">${brl2(sum(sel, t => t.valor))}</b></span><span class="acts"><button class="btn danger ghost" id="sbDel">Excluir</button>${variasEmpresas() ? '<button class="btn" id="sbEmp">Definir CNPJ</button>' : ''}${sel.some(t => !pago(t)) ? '<button class="btn primary" id="sbBx">Dar baixa</button>' : ''}</span>`;
        const se = $('#sbEmp', sb); if (se) se.onclick = () => {
          const f = folha('Definir CNPJ', `<form class="form" id="fx"><p class="muted">${sel.length} lançamento(s) selecionado(s).</p><label>CNPJ da nota / pagamento<select name="emp" required>${opts(empOpts(), '', 'Escolha o CNPJ…')}</select></label><div class="acts"><button type="button" class="btn" data-close>Cancelar</button><button class="btn primary">Aplicar</button></div></form>`);
          f.el.querySelector('#fx').onsubmit = ev => { ev.preventDefault(); const emp = ev.target.emp.value; salvarTitulos(sel.map(t => ({ ...t, emp, alterado: Date.now() }))); f.fechar(); listaSel.clear(); toast(`CNPJ definido em ${sel.length} lançamento(s)`); };
        };
        const sbx = $('#sbBx', sb); if (sbx) sbx.onclick = () => formBaixa(sel.filter(t => !pago(t)).map(t => t.id));
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
  titulo: a => a.length === 2 ? `Quadra ${a[0]}, lote ${a[1]}` : a.length === 1 ? 'Quadra ' + a[0] : 'Quadras e lotes',
  refresh(el, a) { if (a.length === 2) return obraLote(el, a[0], a[1]); if (a.length === 1) return obraQuadra(el, a[0]); return obrasResumo(el); }
};
let obraVista = '';
ACOES.obraVista = d => { obraVista = d.o || ''; refreshView(); };
function obrasResumo(el) {
  if (obraVista && !E().obras.includes(obraVista)) obraVista = '';
  const qsV = quadrasDaObra(obraVista);
  const rows = qsV.map(resumoQuadra); const T = k => sum(rows, r => r[k]);
  const lotesQ = q => lotesDe(q).map(l => infoVendaLote(q, l));
  const todos = qsV.flatMap(lotesQ);
  const seletor = variasObras() ? `<nav class="seg obras-seg" aria-label="Obra">${[['', 'Todas as obras'], ...E().obras.map(o => [o, o])].map(([v, l]) => `<button class="${v === obraVista ? 'on' : ''}" data-act="obraVista" data-o="${esc(v)}">${esc(l)}</button>`).join('')}</nav>` : '';
  const n = { rec: todos.filter(x => x.etapa === 'recebido').length, disp: todos.filter(x => x.etapa === 'disp').length }; n.and = todos.length - n.rec - n.disp;
  el.innerHTML = `${seletor}<section class="obras-cab"><div><p class="eh-lab">${esc(obraVista || (variasObras() ? 'Todas as obras' : E().local))}</p><p class="eh-val fig">${T('n')} lotes em ${qsV.length} quadra${qsV.length === 1 ? '' : 's'}</p></div>
    <div class="eh-fatos"><div><span class="k">Resultado projetado</span><span class="v fig ${cls(T('proj'))}">${short(T('proj'))}</span><small>previsto ${short(T('prevRes'))}</small></div><div><span class="k">A receber de vendas</span><span class="v fig pos">${short(T('receber'))}</span><small>${n.and} lotes vendidos a receber</small></div><div><span class="k">A pagar nas obras</span><span class="v fig">${short(T('pagar'))}</span><small>já pago ${short(T('custo'))}</small></div></div></section>
  <section class="painel mapa"><div class="painel-cab"><h2>Mapa de lotes</h2><p class="legenda"><span><i class="lt rec"></i>Repasse recebido (${n.rec})</span><span><i class="lt and"></i>Vendido, a receber (${n.and})</span><span><i class="lt disp"></i>Disponível (${n.disp})</span></p></div>
    ${rows.map(r => { const ls = lotesQ(r.q); const ov = r.prevCusto ? (r.custo + r.pagar) / r.prevCusto : null; return `<div class="mq">
      <a class="mq-nome" href="#/obras/${encodeURIComponent(r.q)}"><b class="fig">${esc(r.q)}</b><small>${r.n} lotes</small></a>
      <div class="mq-lotes">${ls.map(x => `<a class="lt ${x.etapa === 'recebido' ? 'rec' : x.etapa === 'disp' ? 'disp' : 'and'}" href="#/obras/${encodeURIComponent(r.q)}/${encodeURIComponent(x.l)}" title="Lote ${esc(x.l)}: ${esc(etapaNome(x.etapa))}${x.v && x.v.cliente ? ', ' + esc(x.v.cliente) : ''}" aria-label="Lote ${esc(x.l)}, ${esc(etapaNome(x.etapa))}"></a>`).join('')}</div>
      <div class="mq-num"><span class="fig ${cls(r.proj)}">${short(r.proj)}</span><small class="${ov > 1 ? 'neg' : ''}">${ov === null ? 'sem orçamento' : 'custo em ' + fPct.format(ov) + ' do orçado'}</small></div></div>`; }).join('') || '<p class="muted">Nenhuma quadra cadastrada. Cadastre em Cadastros.</p>'}
  </section>
  <section class="painel"><div class="painel-cab"><h2>Detalhe financeiro por quadra</h2></div><div class="tbl-wrap"><table class="tbl"><thead><tr><th class="l">Quadra</th><th>Lotes</th><th>Custo pago</th><th>A pagar</th><th>Receita</th><th>A receber</th><th>Resultado projetado</th><th>Previsto</th><th>Custo vs orçado</th><th>Checklist</th></tr></thead><tbody>
  ${rows.map(r => `<tr class="click" data-act="ir" data-href="#/obras/${encodeURIComponent(r.q)}"><td class="l"><b>${esc(r.q)}</b></td><td>${r.n}</td><td class="neg">${brl(r.custo)}</td><td>${brl(r.pagar)}</td><td class="pos">${brl(r.receita)}</td><td class="pos">${brl(r.receber)}</td><td class="${cls(r.proj)}"><b>${brl(r.proj)}</b></td><td>${r.prevCusto ? brl(r.prevRes) : '—'}</td><td>${pctBar(r.prevCusto ? (r.custo + r.pagar) / r.prevCusto : null, true)}</td><td>${pctBar(r.cl)}</td></tr>`).join('')}
  </tbody></table></div></section>`;
}
function obraQuadra(el, q) {
  const s = resumoQuadra(q);
  const real = byDest(S.list.filter(t => t.quadra === q && pago(t))), pag = byDest(S.list.filter(t => t.quadra === q && !pago(t) && t.valor < 0)), p = previstoLote(q);
  const sec = secoesViab(); const labs = [...sec.op, ...sec.nonop].filter(k => Math.abs(p[k] || 0) > .5 || Math.abs(real[k] || 0) > .5 || Math.abs(pag[k] || 0) > .5);
  el.innerHTML = `<div class="grid kpis">${kpi('Lotes', s.n, '', '')}${kpi('Custo realizado', brl(s.custo), s.n ? short(s.custo / s.n) + ' por lote' : '', 'neg')}${kpi('A pagar', brl(s.pagar), '', 'neg')}${kpi('Receita realizada', brl(s.receita), '', 'pos')}${kpi('A receber', brl(s.receber), '', 'pos')}${kpi('Resultado projetado', brl(s.proj), (admQuadra(q) > 0.5 ? 'após administração rateada ' + brl(s.proj - admQuadra(q)) + ' · ' : '') + 'previsto ' + brl(s.prevRes), cls(s.proj))}${kpi('Checklist', s.cl === null ? '—' : fPct.format(s.cl), 'documentos', '', '#/checklist/' + encodeURIComponent(q))}</div>
  <div class="grid two">
    <div class="card"><h2>Previsto x realizado <span class="hint">quadra inteira</span> <button class="btn sm" id="bOrc">Editar orçamento</button></h2>
      <div class="legend"><span style="--c:var(--pend)">Previsto</span><span style="--c:var(--bal)">Realizado</span><span style="--c:var(--steel)">A pagar</span></div><div class="chartbox tall"><canvas id="cQ"></canvas></div></div>
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
  const soft = css('--steel');
  grafico('cQ', { type: 'bar', data: { labels: labs, datasets: [
    { label: 'Previsto', data: labs.map(k => -(p[k] || 0) * s.n), backgroundColor: css('--pend'), borderRadius: 3, stack: 'p' },
    { label: 'Realizado', data: labs.map(k => -Math.min(0, real[k] || 0)), backgroundColor: css('--bal'), borderRadius: 3, stack: 'r' },
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
  const adm = admLote(q, l);
  el.innerHTML = `<div class="grid kpis">${kpi('Custo realizado', brl(custo), (adm > 0.5 ? '+ ' + brl(adm) + ' de administração rateada · ' : '') + 'previsto ' + brl(iP.op + iP.non), 'neg')}${kpi('Lucro líquido projetado', brl(iC.ll), 'previsto ' + brl(iP.ll), cls(iC.ll))}${kpi('ROI projetado', pct(iC.roi), 'previsto ' + pct(iP.roi), '')}${kpi('Checklist', ckPct(q, l) === null ? '—' : fPct.format(ckPct(q, l)), 'documentos', '', '#/checklist/' + encodeURIComponent(q))}</div>
  <div class="grid lg"><div class="card"><h2>Viabilidade <span class="hint">projetado = realizado + em aberto</span></h2><div class="tbl-wrap"><table class="tbl">${html}</tbody></table></div></div>
  <div class="card"><h2>Indicadores</h2><div class="tbl-wrap"><table class="tbl"><thead><tr><th class="l">Indicador</th><th>Previsto</th><th>Realizado</th><th>Projetado</th></tr></thead><tbody>${ind.map(([n, k, fm]) => `<tr><td class="l">${n}</td><td>${fm(iP[k])}</td><td class="${typeof iR[k] === 'number' ? cls(iR[k]) : ''}">${fm(iR[k])}</td><td class="${typeof iC[k] === 'number' ? cls(iC[k]) : ''}"><b>${fm(iC[k])}</b></td></tr>`).join('')}</tbody></table></div>
  <p class="muted small">Realizado considera só o que já foi pago/recebido. Enquanto a venda não entra, o resultado realizado fica negativo.</p></div></div>
  ${cardDadosLote(q, l)}
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
    const ok = (x, l, i) => estadoDoc(x, l, i).feito;
    const dd = S.docsdrive || {};
    el.innerHTML = `<div class="filtros card"><select id="ckq">${opts(qs, q, 'Todas as quadras')}</select><span class="muted">Toque no lote para ver e enviar os documentos; toque numa bolinha para marcar ou desmarcar à mão.</span></div>
    ${cardDocsDrive(dd)}
    <div class="grid two"><div class="card"><h2>Por documento</h2><div class="tbl-wrap"><table class="tbl"><tbody>${docs.map((d, i) => { const n = lotes.filter(([x, l]) => ok(x, l, i)).length; const nv = lotes.filter(([x, l]) => estadoDoc(x, l, i).venc).length; return `<tr><td class="l">${esc(d)}${nv ? ` <span class="chip bad">${nv} vencido(s)</span>` : ''}</td><td>${n}/${lotes.length}</td><td>${pctBar(lotes.length ? n / lotes.length : null)}</td></tr>`; }).join('')}</tbody></table></div></div>
    <div class="card"><h2>Por quadra</h2><div class="tbl-wrap"><table class="tbl"><thead><tr><th class="l">Quadra</th><th>Lotes</th><th>Completos</th><th>Andamento</th></tr></thead><tbody>${qs.map(x => { const ls = lotesDe(x); const full = ls.filter(l => docs.every((_, i) => ok(x, l, i))).length; return `<tr class="click" data-act="ir" data-href="#/checklist/${encodeURIComponent(x)}"><td class="l"><b>${esc(x)}</b></td><td>${ls.length}</td><td>${full}</td><td>${pctBar(ckPct(x))}</td></tr>`; }).join('')}</tbody></table></div></div></div>
    <div class="card"><h2>Matriz</h2><div class="tbl-wrap matriz"><table class="tbl cl"><thead><tr><th class="l">Quadra</th><th class="l">Lote</th>${docs.map(d => `<th><span>${esc(d)}</span></th>`).join('')}</tr></thead><tbody>
    ${lotes.map(([x, l]) => `<tr><td class="l">${esc(x)}</td><td class="l"><button class="btn link sm" data-act="ckLote" data-k="${esc(x + '|' + l)}">${esc(l)}</button></td>${docs.map((d, i) => { const e = estadoDoc(x, l, i); return `<td><button class="dot ${e.feito ? 'y' : ''} ${e.arq ? 'arq' : ''} ${e.venc ? 'venc' : ''}" data-act="${e.arq ? 'ckLote' : 'ck'}" data-k="${esc(x + '|' + l)}" data-i="${i}" title="${esc(d + (e.arq ? ' · arquivo de ' + dBR(e.arq.d) + (e.venc ? ' (vencido)' : '') : e.manual ? ' · marcado à mão' : ''))}" aria-label="${esc(d)}"></button></td>`; }).join('')}</tr>`).join('')}</tbody></table></div>
      <p class="legenda-ck muted small"><span><i class="dot y"></i> marcado à mão</span><span><i class="dot y arq"></i> arquivo no Drive</span><span><i class="dot venc arq"></i> arquivo vencido</span><span><i class="dot"></i> pendente</span></p></div>`;
    if (dd.ler && tokenTem(META_SCOPE) && (!dd.ultima || Date.now() - dd.ultima > 10 * 60000) && !DD.lendo) setTimeout(() => lerDocsDrive(false), 300);
    $('#ckq', el).onchange = e => { location.hash = '#/checklist' + (e.target.value ? '/' + encodeURIComponent(e.target.value) : ''); };
  }
};
ACOES.ck = d => {
  const c = S.checklist = S.checklist || { docs: DOCS_PADRAO, lotes: {} }; c.lotes = c.lotes || {};
  const docs = c.docs || DOCS_PADRAO; let s = (c.lotes[d.k] || '').padEnd(docs.length, '0').split(''); const i = +d.i;
  s[i] = s[i] === '1' ? '0' : '1'; c.lotes[d.k] = s.join('');
  gravar(doc(db, 'config', 'checklist'), { docs, lotes: { [d.k]: c.lotes[d.k] } }, { merge: true }).catch(onErroEscrita);
  const b = document.querySelector(`[data-act="ck"][data-k="${CSS.escape(d.k)}"][data-i="${i}"]`); if (b) b.classList.toggle('y', s[i] === '1');
  if (d.depois) d.depois();
};

/* ---------- Documentos dos lotes no Google Drive ---------- */
const META_SCOPE = 'https://www.googleapis.com/auth/drive.metadata.readonly';
const DD = { lendo: false, erro: '', prog: '' };
const nomeQuadraPasta = q => 'Quadra ' + q, nomeLotePasta = l => 'Lote ' + l;
const semPrefixo = (s, pre) => normTxt(s).replace(new RegExp('^(' + pre + ')\\s+'), '');
function docDoNome(nome) {
  const n = ' ' + normTxt(nome.replace(/\.[a-z0-9]{2,4}$/i, '')) + ' '; let melhor = '';
  for (const d of docsCk()) { const k = normTxt(d); if (k && n.includes(' ' + k + ' ') && k.length > normTxt(melhor).length) melhor = d; }
  return melhor;
}
function dataDoNome(nome, fallback) {
  let m = nome.match(/(20\d{2})[-_. ](\d{2})[-_. ](\d{2})/); if (m) return `${m[1]}-${m[2]}-${m[3]}`;
  m = nome.match(/(\d{2})[-_. ](\d{2})[-_. ](20\d{2})/); if (m) return `${m[3]}-${m[2]}-${m[1]}`;
  return (fallback || '').slice(0, 10);
}
async function listarPasta(id) {
  const out = []; let tok = '';
  do {
    const q = encodeURIComponent(`'${id}' in parents and trashed=false`);
    const r = await driveReq(`https://www.googleapis.com/drive/v3/files?q=${q}&fields=nextPageToken,files(id,name,mimeType,createdTime,modifiedTime)&pageSize=1000${tok ? '&pageToken=' + tok : ''}`);
    if (!r) break; out.push(...(r.files || [])); tok = r.nextPageToken || '';
  } while (tok);
  return out;
}
async function raizSFG() {
  const cfg = JSON.parse(JSON.stringify(S.drive || {}));
  if (cfg.raiz && await pastaExiste(cfg.raiz.sfg)) return cfg.raiz.sfg;
  const sfg = await novaPasta('SFG Financeiro'); const an = await novaPasta('Anexos', sfg.id);
  cfg.raiz = { sfg: sfg.id, anexos: an.id, link: sfg.webViewLink }; cfg.anos = {}; cfg.meses = {}; S.drive = cfg; gravar(doc(db, 'config', 'drive'), cfg).catch(onErroEscrita);
  return sfg.id;
}
function salvarDocsDrive(cfg) { const c = { ...cfg }; delete c.arq; delete c.outros; S.docsdrive = c; gravar(doc(db, 'config', 'docsdrive'), c).catch(onErroEscrita); }
const blocoDoc = chave => { const [n, a] = chave.split('|'); return n === 'O' ? 'O_' + idBloco(a) : 'Q_' + idBloco(a); };
/* Grava os arquivos encontrados agrupados por quadra (e por obra). */
function salvarArqDocs(arq, outros, parcial) {
  const por = {};
  for (const k in arq) { const b = blocoDoc(k); (por[b] = por[b] || { arq: {}, outros: {} }).arq[k] = arq[k]; }
  for (const k in (outros || {})) { const b = blocoDoc(k); (por[b] = por[b] || { arq: {}, outros: {} }).outros[k] = outros[k]; }
  if (!parcial) for (const b of ((S.docsArq && S.docsArq.blocos) || [])) if (!por[b]) por[b] = { arq: {}, outros: {} };
  for (const b in por) gravar(doc(db, 'docsArq', b), por[b], parcial ? { merge: true } : undefined).catch(onErroEscrita);
  const atual = parcial ? { arq: { ...((S.docsArq || {}).arq || {}), ...arq }, outros: { ...((S.docsArq || {}).outros || {}), ...(outros || {}) } } : { arq, outros: outros || {} };
  S.docsArq = { ...atual, blocos: [...new Set([...(((S.docsArq || {}).blocos) || []), ...Object.keys(por)])] };
}
/* Garante a pasta Documentos / Obra / Quadra / Lote e devolve o id. */
async function pastaDocs(q, l) {
  const cfg = clone(S.docsdrive || {}); cfg.p = cfg.p || {}; let mudou = false;
  if (!cfg.raiz || !(await pastaExiste(cfg.raiz))) { const r = await novaPasta('Documentos', await raizSFG()); cfg.raiz = r.id; cfg.link = r.webViewLink; cfg.p = {}; mudou = true; }
  const o = obraDe(q); const ko = 'O|' + o, kq = 'Q|' + q, kl = 'L|' + q + '|' + l;
  if (!cfg.p[ko]) { cfg.p[ko] = (await novaPasta(o, cfg.raiz)).id; mudou = true; }
  if (!cfg.p[kq]) { cfg.p[kq] = (await novaPasta(nomeQuadraPasta(q), cfg.p[ko])).id; mudou = true; }
  if (l && !cfg.p[kl]) { cfg.p[kl] = (await novaPasta(nomeLotePasta(l), cfg.p[kq])).id; mudou = true; }
  if (mudou) salvarDocsDrive(cfg);
  return l ? cfg.p[kl] : cfg.p[kq];
}
async function criarPastasDocs() {
  if (!(await driveToken(true).catch(e => { toast(e.message, 'err'); return null; }))) return;
  const lotes = E().quadras.flatMap(q => lotesDe(q).map(l => [q, l])); let i = 0;
  DD.lendo = true; DD.erro = '';
  try { for (const [q, l] of lotes) { DD.prog = `Criando pastas ${++i}/${lotes.length}`; if (i % 10 === 1) refreshView(); await pastaDocs(q, l); } toast('Pastas criadas no Drive'); }
  catch (e) { DD.erro = e.message === 'drive-desconectado' ? 'A autorização do Google expirou. Toque de novo.' : e.message; toast(DD.erro, 'err'); }
  finally { DD.lendo = false; DD.prog = ''; refreshView(); }
}
/* Percorre a pasta Documentos e associa cada arquivo ao documento, lote, quadra ou obra. */
async function lerDocsDrive(interativo) {
  if (DD.lendo || !driveConfigurado() || !navigator.onLine) return;
  const cfg0 = S.docsdrive || {}; if (!cfg0.raiz) { if (interativo) toast('Crie as pastas primeiro.', 'err'); return; }
  let tk; try { tk = await driveToken(interativo, cfg0.ler ? META_SCOPE : DRIVE_SCOPE); } catch (e) { if (interativo) toast(e.message, 'err'); return; }
  if (!tk) return;
  DD.lendo = true; DD.erro = ''; DD.prog = 'Lendo pastas do Drive…'; refreshView();
  const arq = {}, outros = {}; let nArq = 0;
  const guarda = (chave, f) => { const d = docDoNome(f.name); if (!d) { (outros[chave] = outros[chave] || []).push(f.name); return; } const x = { i: f.id, n: f.name, d: dataDoNome(f.name, f.modifiedTime || f.createdTime) }; const at = (arq[chave] = arq[chave] || {})[d]; if (!at || (x.d || '') > (at.d || '')) arq[chave][d] = { ...x, c: (at ? at.c || 1 : 0) + 1 }; else at.c = (at.c || 1) + 1; nArq++; };
  const pasta = 'application/vnd.google-apps.folder';
  const arquivosEm = async (id, chave, docPasta) => {
    for (const f of await listarPasta(id)) {
      if (f.mimeType === pasta) { const d = docDoNome(f.name); if (d && !docPasta) for (const g of await listarPasta(f.id)) { if (g.mimeType !== pasta) guarda(chave, { ...g, name: d + ' ' + g.name }); } continue; }
      guarda(chave, docPasta ? { ...f, name: docPasta + ' ' + f.name } : f);
    }
  };
  try {
    const cfg = clone(cfg0); cfg.p = cfg.p || {};
    for (const fo of await listarPasta(cfg.raiz)) {
      if (fo.mimeType !== pasta) continue; const o = E().obras.find(x => normTxt(x) === normTxt(fo.name)); if (!o) continue;
      cfg.p['O|' + o] = fo.id;
      for (const fq of await listarPasta(fo.id)) {
        if (fq.mimeType !== pasta) { guarda('O|' + o, fq); continue; }
        const q = quadrasDaObra(o).find(x => normTxt(x) === semPrefixo(fq.name, 'quadra|qd|q')); if (!q) { if (docDoNome(fq.name)) await arquivosEm(fq.id, 'O|' + o, docDoNome(fq.name)); continue; }
        cfg.p['Q|' + q] = fq.id; DD.prog = 'Lendo quadra ' + q + '…'; refreshView();
        for (const fl of await listarPasta(fq.id)) {
          if (fl.mimeType !== pasta) { guarda('Q|' + q, fl); continue; }
          const l = lotesDe(q).find(x => normTxt(x) === semPrefixo(fl.name, 'lote|lt|l'));
          if (!l) { if (docDoNome(fl.name)) await arquivosEm(fl.id, 'Q|' + q, docDoNome(fl.name)); continue; }
          cfg.p['L|' + q + '|' + l] = fl.id; await arquivosEm(fl.id, 'L|' + q + '|' + l);
        }
      }
    }
    cfg.ultima = Date.now(); cfg.n = nArq; salvarDocsDrive(cfg); salvarArqDocs(arq, outros);
    if (interativo) toast(`${nArq} arquivo(s) encontrados no Drive`);
  } catch (e) { DD.erro = e.message === 'drive-desconectado' ? 'A autorização do Google expirou. Toque em Atualizar do Drive.' : e.message; if (interativo) toast(DD.erro, 'err'); }
  finally { DD.lendo = false; DD.prog = ''; refreshView(); }
}
function cardDocsDrive(dd) {
  if (!driveConfigurado()) return '';
  return `<div class="card docs-drive"><h2>Documentos no Google Drive ${dd.link ? `<a class="btn sm" href="${esc(dd.link)}" target="_blank" rel="noopener">Abrir pasta</a>` : ''}</h2>
    <p class="muted small">Pasta <b>SFG Financeiro / Documentos / obra / Quadra / Lote</b>. Coloque o nome do documento no nome do arquivo (ex.: "ART lote 19.pdf", "Certidão de ônus 2026-10-15.pdf"). O app marca o checklist sozinho e usa sempre o arquivo mais recente. Um arquivo na pasta da quadra ou da obra vale para todos os lotes dela.</p>
    ${dd.raiz ? `<div class="kv"><span>Última leitura</span><b>${dd.ultima ? quando(dd.ultima) + ' · ' + (dd.n || 0) + ' arquivo(s)' : 'ainda não'}</b></div>` : ''}
    ${DD.prog ? `<p class="muted small">${esc(DD.prog)}</p>` : ''}${DD.erro ? `<p class="erro small">${esc(DD.erro)}</p>` : ''}
    <div class="acts left rw">${dd.raiz ? `<button class="btn primary" data-act="ddLer" ${DD.lendo ? 'disabled' : ''}>Atualizar do Drive</button><button class="btn" data-act="ddCriar" ${DD.lendo ? 'disabled' : ''}>Criar pastas que faltam</button>` : `<button class="btn primary" data-act="ddCriar" ${DD.lendo ? 'disabled' : ''}>Criar as pastas no Drive</button>`}
      ${dd.raiz ? (dd.ler ? `<button class="btn ghost" data-act="ddLerOff">Só ver arquivos enviados pelo app</button>` : `<button class="btn" data-act="ddLerOn">Ver também arquivos que eu colocar direto no Drive</button>`) : ''}</div>
    ${dd.raiz && !dd.ler ? '<p class="muted small">Hoje o app só enxerga os arquivos que ele mesmo enviou. Para ler o que você arrastar direto para as pastas, ligue a opção acima: o Google vai pedir permissão para o app ver os nomes dos arquivos do seu Drive (só nomes e datas, sem abrir o conteúdo).</p>' : ''}</div>`;
}
ACOES.ddCriar = () => criarPastasDocs();
ACOES.ddLer = () => lerDocsDrive(true);
ACOES.ddLerOn = async () => { const c = { ...(S.docsdrive || {}), ler: true }; salvarDocsDrive(c); try { await driveToken(true, META_SCOPE); lerDocsDrive(true); } catch (e) { toast(e.message, 'err'); } };
ACOES.ddLerOff = () => { salvarDocsDrive({ ...(S.docsdrive || {}), ler: false }); refreshView(); };
ACOES.docVal = d => {
  const nd = docsCk()[+d.i]; const v = ((S.checklist && S.checklist.validade) || {})[nd] || '';
  const f = folha('Documento: ' + nd, `<form class="form" id="fx"><label>Validade em dias <span class="muted">(vazio = não vence)</span><input name="v" type="number" min="0" inputmode="numeric" value="${esc(v)}" placeholder="Ex.: 30"></label><p class="muted small">A validade conta a partir da data no nome do arquivo (2026-10-15 ou 15-10-2026) ou, sem data no nome, do dia em que ele foi colocado no Drive.</p><div class="acts"><button type="button" class="btn" data-close>Cancelar</button><button class="btn primary">Salvar</button></div></form>`);
  f.el.querySelector('#fx').onsubmit = e => { e.preventDefault(); const n = +e.target.v.value || 0; const c = S.checklist = S.checklist || { docs: DOCS_PADRAO, lotes: {} }; c.validade = { ...(c.validade || {}) }; if (n) c.validade[nd] = n; else delete c.validade[nd]; gravar(doc(db, 'config', 'checklist'), { docs: docsCk(), validade: c.validade }, { merge: true }).catch(onErroEscrita); f.fechar(); refreshView(); };
};
/* Documentos de um lote: situação de cada um, abrir, enviar arquivo, marcar à mão. */
ACOES.ckLote = d => {
  const [q, l] = d.k.split('|'); const docs = docsCk(); const dd = S.docsdrive || {};
  const pl = dd.p && dd.p['L|' + q + '|' + l]; const outros = ((((S.docsArq || {}).outros) || dd.outros || {})['L|' + q + '|' + l] || []);
  const linha = (doc, i) => { const e = estadoDoc(q, l, i); return `<div class="ckdoc ${e.feito ? 'ok' : ''} ${e.venc ? 'venc' : ''}"><span class="dot ${e.feito ? 'y' : ''} ${e.arq ? 'arq' : ''} ${e.venc ? 'venc' : ''}"></span><div class="ckdoc-txt"><b>${esc(doc)}</b><small>${e.arq ? `${esc(e.arq.n)} · ${dBR(e.arq.d)}${e.dias ? (e.venc ? ` · <span class="neg">vencido em ${dBR(e.vale)}</span>` : ` · vale até ${dBR(e.vale)}`) : ''}${e.arq.c > 1 ? ` · ${e.arq.c} versões` : ''}` : e.manual ? 'marcado à mão, sem arquivo' : 'pendente'}</small></div>
    <span class="ckdoc-acts">${e.arq ? `<a class="btn sm" href="https://drive.google.com/file/d/${esc(e.arq.i)}/view" target="_blank" rel="noopener">Abrir</a>` : ''}<label class="btn sm file rw">Enviar<input type="file" data-up="${i}" accept="image/*,application/pdf" hidden></label><button type="button" class="btn sm ghost rw" data-man="${i}">${e.manual ? 'Desmarcar' : 'Marcar'}</button></span></div>`; };
  const f = folha(`Documentos · Quadra ${q}, lote ${l}`, `<div class="ckdocs">${docs.map(linha).join('')}</div>
    ${outros.length ? `<p class="muted small">Arquivos na pasta do lote sem documento reconhecido: ${esc(outros.join(', '))}. Coloque o nome do documento no nome do arquivo.</p>` : ''}
    <div class="acts left"><a class="btn" href="#/etapas/${encodeURIComponent(q)}/${encodeURIComponent(l)}" data-close>Ver fluxo e prazos do lote</a>${pl ? `<a class="btn" href="https://drive.google.com/drive/folders/${esc(pl)}" target="_blank" rel="noopener">Abrir pasta do lote no Drive</a>` : ''}</div>`, { largo: true });
  $$('[data-man]', f.el).forEach(b => b.onclick = () => { ACOES.ck({ k: d.k, i: b.dataset.man }); f.fechar(); ACOES.ckLote(d); });
  $$('[data-up]', f.el).forEach(inp => inp.onchange = async () => {
    const file = inp.files[0]; if (!file) return; const doc = docs[+inp.dataset.up];
    if (!navigator.onLine) { toast('Enviar documentos precisa de internet.', 'err'); return; }
    try {
      if (!(await driveToken(true))) return; toast('Enviando ' + doc + '…');
      const pasta = await pastaDocs(q, l); const ext = (file.name.match(/\.[a-z0-9]{2,4}$/i) || [file.type === 'application/pdf' ? '.pdf' : '.jpg'])[0];
      const nome = `${doc} - Quadra ${q} lote ${l} - ${hoje()}${ext}`.replace(/[\/\\:*?"<>|#%]+/g, '-');
      const r = await enviarArquivoDrive(file, nome, pasta);
      const k = 'L|' + q + '|' + l; const daLote = { ...((((S.docsArq || {}).arq) || {})[k] || {}) }; const at = daLote[doc];
      daLote[doc] = { i: r.id, n: nome, d: hoje(), c: (at ? at.c || 1 : 0) + 1 }; salvarArqDocs({ [k]: daLote }, null, true);
      const cfg = clone(S.docsdrive || {}); cfg.n = (cfg.n || 0) + 1; salvarDocsDrive(cfg);
      toast(doc + ' enviado'); f.fechar(); ACOES.ckLote(d);
    } catch (e) { toast(e.message === 'drive-desconectado' ? 'A autorização do Google expirou. Tente de novo.' : e.message, 'err'); }
  });
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
      <div class="card"><h2>Obras <button class="btn sm rw" data-act="cadObra">Nova obra</button></h2><div class="tbl-wrap"><table class="tbl"><thead><tr><th class="l">Obra</th>${variasEmpresas() ? '<th class="l">Empresa</th>' : ''}<th>Quadras</th><th>Lanç.</th><th></th></tr></thead><tbody>
        ${e.obras.map(o => `<tr><td class="l"><b>${esc(o)}</b>${obraLocal(o) ? `<br><span class="muted small">${esc(obraLocal(o))}</span>` : ''}</td>${variasEmpresas() ? `<td class="l">${esc(empRotulo(empPorId(empDaObra(o)), true))}</td>` : ''}<td>${quadrasDaObra(o).length}</td><td>${S.list.filter(x => obraT(x) === o).length}</td><td><button class="btn sm rw" data-act="cadObra" data-o="${esc(o)}">Editar</button></td></tr>`).join('')}</tbody></table></div></div>
      <div class="card"><h2>Contas bancárias <button class="btn sm" data-act="cadConta">Adicionar</button><button class="btn sm" data-act="transferir">Transferir</button></h2><div class="tbl-wrap"><table class="tbl"><thead><tr><th class="l">Conta</th><th>Saldo inicial</th><th>Saldo atual</th><th></th></tr></thead><tbody>
        ${saldos().contas.map(c => { const cd = e.contas.find(x => x.nome === c.nome) || {}; return { ...c, cd }; }).map(c => `<tr><td class="l"><b>${esc(c.nome)}</b>${c.cd.num ? `<br><span class="muted small">${esc([c.cd.banco, c.cd.ag ? 'ag. ' + c.cd.ag : '', 'cc ' + c.cd.num].filter(Boolean).join(' · '))}</span>` : `<br><span class="small" style="color:var(--warn)">sem número da conta</span>`}${variasEmpresas() ? `<br><span class="muted small">${esc(empRotulo(empPorId(empConta(c.nome)), true))}</span>` : ''}</td><td>${brl2(c.inicial)}</td><td class="${cls(c.saldo)}">${brl2(c.saldo)}</td><td>${e.contas.some(x => x.nome === c.nome) ? `<button class="btn sm" data-act="cadConta" data-nome="${esc(c.nome)}">Editar</button>` : ''}</td></tr>`).join('')}</tbody></table></div></div>
    </div>
    <div class="card"><h2>Empresas (CNPJ) <button class="btn sm rw" data-act="cadEmp">Nova filial</button></h2>
      <p class="muted small">Cadastre a matriz e cada filial aberta para um projeto no RET. Cada obra pertence a um CNPJ, e os lançamentos da obra vão para ele.</p>
      <div class="tbl-wrap"><table class="tbl"><thead><tr><th class="l">Empresa</th><th class="l">CNPJ</th><th class="l">Tipo</th><th class="l">Obras</th><th>Lançamentos</th><th></th></tr></thead><tbody>
      ${e.empresas.map(x => `<tr><td class="l"><b>${esc(x.nome)}</b></td><td class="l fig">${esc(fCnpj(x.cnpj)) || '<span class="muted">não informado</span>'}</td><td class="l"><span class="chip ${x.tipo === 'matriz' ? 'ok' : ''}">${x.tipo === 'matriz' ? 'Matriz' : 'Filial'}${x.ret ? ' · RET' : ''}</span></td><td class="l wrap">${esc(e.obras.filter(o => empDaObra(o) === x.id).join(', ')) || '<span class="muted">nenhuma</span>'}</td><td>${S.list.filter(t => empT(t) === x.id).length}</td><td><button class="btn sm rw" data-act="cadEmp" data-id="${esc(x.id)}">Editar</button></td></tr>`).join('')}</tbody></table></div></div>
    <div class="card"><h2>Quadras e lotes <button class="btn sm" data-act="cadQuadra">Adicionar quadra</button></h2><div class="tbl-wrap"><table class="tbl"><thead><tr><th class="l">Quadra</th>${variasObras() ? '<th class="l">Obra</th>' : ''}<th>Lotes</th><th class="l">Numeração</th><th>Lançamentos</th><th></th></tr></thead><tbody>
      ${e.quadras.map(q => { const ls = lotesDe(q); return `<tr><td class="l"><b>${esc(q)}</b></td>${variasObras() ? `<td class="l">${esc(obraDe(q))}</td>` : ''}<td>${ls.length}</td><td class="l wrap muted small">${esc(ls.slice(0, 30).join(', '))}${ls.length > 30 ? '…' : ''}</td><td>${usados('quadra', q)}</td><td><button class="btn sm" data-act="cadQuadra" data-q="${esc(q)}">Editar</button></td></tr>`; }).join('')}</tbody></table></div></div>
    <div class="card"><h2>Destinos (categorias)</h2><div class="grid two">${['op', 'nonop', 'rec', 'aporte'].map(s => `<div><div class="lbl">${SECAO_NOME[s]} <button class="btn link sm" data-act="cadDest" data-s="${s}">+ adicionar</button></div><div class="chips">${(e.destinos[s] || []).map(d => `<span class="chip">${esc(d)} <button class="x" data-act="rmDest" data-s="${s}" data-d="${esc(d)}" aria-label="Remover">×</button></span>`).join('')}</div></div>`).join('')}</div></div>
    <div class="card"><h2>Feriados <button class="btn sm rw" data-act="cadFeriado">Adicionar feriado local</button></h2>
      <p class="muted small">Vencimentos que caem nestes dias (ou em fins de semana) vão para o próximo dia útil. Os feriados nacionais, inclusive Carnaval, Sexta-feira Santa e Corpus Christi, já são calculados todo ano. Inclua os feriados de Juazeiro e da Bahia que afetam os seus pagamentos.</p>
      <div class="chips">${((e.feriados) || []).map((f, i) => `<span class="chip">${esc(f.d.length === 5 ? f.d.slice(3) + '/' + f.d.slice(0, 2) + ' (todo ano)' : dBR(f.d))} ${esc(f.nome || '')} <button class="x rw" data-act="rmFeriado" data-i="${i}" aria-label="Remover">×</button></span>`).join('') || '<span class="muted small">Nenhum feriado local cadastrado.</span>'}</div>
      <details style="margin-top:10px"><summary class="muted small">Ver feriados nacionais de ${new Date().getFullYear()}</summary><div class="chips" style="margin-top:8px">${[...feriadosNacionais(new Date().getFullYear()).entries()].sort().map(([d, n]) => `<span class="chip">${dBR(d).slice(0, 5)} ${esc(n)}</span>`).join('')}</div></details></div>
    <div class="card"><h2>Documentos do checklist <button class="btn sm" data-act="cadDoc">Adicionar</button></h2><p class="muted small">Toque num documento para definir a validade em dias (ex.: certidão de ônus, 30 dias). O app avisa quando o arquivo mais recente passar da validade.</p><div class="chips">${docsCk().map((d, i) => { const v = ((S.checklist && S.checklist.validade) || {})[d]; return `<button class="chip" data-act="docVal" data-i="${i}">${esc(d)}${v ? ` · ${v} dias` : ''}</button>`; }).join('')}</div></div>`;
  }
};
const clone = o => JSON.parse(JSON.stringify(o));
ACOES.cadFeriado = () => {
  const f = folha('Feriado local', `<form class="form" id="fx"><div class="g2"><label>Data<input type="date" name="d" required></label><label>Nome<input name="nome" placeholder="Ex.: Independência da Bahia"></label></div><label class="chk"><input type="checkbox" name="anual" checked> Repete todo ano</label><div class="acts"><button type="button" class="btn" data-close>Cancelar</button><button class="btn primary">Adicionar</button></div></form>`);
  f.el.querySelector('#fx').onsubmit = ev => { ev.preventDefault(); const F = ev.target; const n = clone(E()); n.feriados = [...n.feriados, { d: F.anual.checked ? F.d.value.slice(5) : F.d.value, nome: F.nome.value.trim() }]; n.feriados.sort((a, b) => a.d.slice(-5).localeCompare(b.d.slice(-5))); salvarEstrutura(n); f.fechar(); toast('Feriado adicionado'); };
};
ACOES.rmFeriado = d => { const n = clone(E()); n.feriados = n.feriados.filter((_, i) => i !== +d.i); salvarEstrutura(n); };
ACOES.tema = d => aplicarTema(d.v);
ACOES.temaCiclo = () => { const o = ['auto', 'claro', 'escuro']; const m = o[(o.indexOf(temaPref()) + 1) % 3]; aplicarTema(m); toast('Tema: ' + TEMAS.find(x => x[0] === m)[1]); };
ACOES.cadEmp = d => {
  const x = d.id ? empPorId(d.id) : null; const e = E();
  const nT = x ? S.list.filter(t => empT(t) === x.id).length : 0, nO = x ? e.obras.filter(o => empDaObra(o) === x.id).length : 0;
  const pode = x && x.tipo !== 'matriz' && !nT && !nO && !e.contas.some(c => c.emp === x.id);
  const f = folha(x ? x.nome : 'Nova filial', `<form class="form" id="fx">
    <label>Nome<input name="nome" value="${esc(x?.nome || '')}" required placeholder="Ex.: SFG Araripina SPE"></label>
    <div class="g2"><label>CNPJ<input name="cnpj" inputmode="numeric" value="${esc(fCnpj(x?.cnpj || ''))}" placeholder="00.000.000/0000-00"></label>
    <label>Tipo<select name="tipo">${opts([['filial', 'Filial'], ['matriz', 'Matriz']], x?.tipo || 'filial')}</select></label></div>
    <label class="chk"><input type="checkbox" name="ret" ${x ? (x.ret ? 'checked' : '') : 'checked'}> Optante pelo RET (patrimônio de afetação)</label>
    <div class="g2"><label>Alíquota do RET<select name="retAliq">${opts([['4', '4% (regra geral)'], ['1', '1% (MCMV Faixa 1)']], String(x?.retAliq || 4))}</select></label><label>Chave PIX para cobranças<input name="pix" value="${esc(x?.pix || '')}" placeholder="CNPJ, e-mail, telefone ou chave aleatória"></label></div>
    <div class="g2"><label>Nome no PIX<input name="pixNome" value="${esc(x?.pixNome || x?.nome || '')}" maxlength="25"></label><label>Cidade no PIX<input name="pixCidade" value="${esc(x?.pixCidade || 'Juazeiro')}" maxlength="15"></label></div>
    <label>Obras desta empresa<div class="chips">${e.obras.map(o => `<label class="chip"><input type="checkbox" name="obras" value="${esc(o)}" ${x && empDaObra(o) === x.id ? 'checked' : ''}> ${esc(o)}</label>`).join('')}</div></label>
    <p class="muted small" id="cnpjAviso"></p>
    ${x && nT ? `<p class="muted small">${nT} lançamento(s) nesta empresa.</p>` : ''}
    <div class="acts">${pode ? '<button type="button" class="btn danger ghost" id="rm">Remover</button>' : ''}<button type="button" class="btn" data-close>Cancelar</button><button class="btn primary">Salvar</button></div></form>`);
  const F = f.el.querySelector('#fx');
  const aviso = () => { const v = F.cnpj.value.replace(/\D/g, ''); $('#cnpjAviso', F).textContent = v && !cnpjValido(v) ? 'Atenção: esse CNPJ parece inválido (confira os dígitos).' : ''; };
  F.cnpj.oninput = aviso; F.cnpj.onblur = () => { F.cnpj.value = fCnpj(F.cnpj.value.replace(/\D/g, '')) || F.cnpj.value; aviso(); };
  F.onsubmit = ev => {
    ev.preventDefault(); const n = clone(E());
    const cnpj = F.cnpj.value.replace(/\D/g, ''); const nome = F.nome.value.trim(); if (!nome) return;
    if (cnpj && n.empresas.some(y => y.cnpj === cnpj && (!x || y.id !== x.id))) { toast('Esse CNPJ já está cadastrado', 'err'); return; }
    const id = x ? x.id : uid();
    const obj = { ...(x || {}), id, nome, cnpj, tipo: F.tipo.value, ret: F.ret.checked, retAliq: +F.retAliq.value || 4, pix: F.pix.value.trim(), pixNome: F.pixNome.value.trim(), pixCidade: F.pixCidade.value.trim() };
    if (x) n.empresas = n.empresas.map(y => y.id === id ? obj : y); else n.empresas.push(obj);
    if (obj.tipo === 'matriz') n.empresas.forEach(y => { if (y.id !== id && y.tipo === 'matriz') y.tipo = 'filial'; });
    if (!n.empresas.some(y => y.tipo === 'matriz')) n.empresas[0].tipo = 'matriz';
    const marcadas = $$('[name=obras]:checked', F).map(c => c.value);
    for (const o of n.obras) { if (marcadas.includes(o)) n.obraEmpresa[o] = id; else if (n.obraEmpresa[o] === id) delete n.obraEmpresa[o]; }
    salvarEstrutura(n); f.fechar(); toast(x ? 'Empresa salva' : 'Filial cadastrada');
  };
  const rm = f.el.querySelector('#rm'); if (rm) rm.onclick = () => { const n = clone(E()); n.empresas = n.empresas.filter(y => y.id !== x.id); salvarEstrutura(n); f.fechar(); };
};
ACOES.cadObra = d => {
  const o = d.o; const e = E();
  const nQ = o ? quadrasDaObra(o).length : 0, nT = o ? S.list.filter(x => obraT(x) === o).length : 0;
  const pode = o && e.obras.length > 1 && !nQ && !nT; const oi = (o && e.obraInfo[o]) || {};
  const f = folha(o ? 'Obra ' + o : 'Nova obra', `<form class="form" id="fx"><label>Nome da obra (local)<input name="v" value="${esc(o || '')}" required placeholder="Ex.: Araripina"></label>
    ${variasEmpresas() ? `<label>Empresa (CNPJ) da obra<select name="emp">${opts(empOpts(), o ? empDaObra(o) : matrizId())}</select></label>` : ''}
    <div class="g3"><label>Cidade<input name="cidade" value="${esc(oi.cidade || '')}" placeholder="Ex.: Petrolina"></label><label>UF<input name="uf" maxlength="2" value="${esc(oi.uf || '')}" placeholder="PE" style="text-transform:uppercase"></label><label>CEP<input name="cep" inputmode="numeric" value="${esc(oi.cep || '')}"></label></div>
    <label>Endereço / localização do empreendimento<input name="endereco" value="${esc(oi.endereco || '')}" placeholder="Ex.: Loteamento Nova Juazeiro, Rod. BR-407 km 3"></label>
    <div class="g2"><label>Construtor responsável<input name="construtor" list="dlConstr" value="${esc(oi.construtor || '')}" placeholder="Empresa ou empreiteiro"></label><label>Matrícula da gleba / loteamento<input name="matricula" value="${esc(oi.matricula || '')}"></label></div>
    <div class="g2"><label>Cartório de registro<input name="cartorio" value="${esc(oi.cartorio || '')}"></label><label>Mapa (link ou coordenadas)<input name="mapa" value="${esc(oi.mapa || '')}" placeholder="-9.39, -40.50 ou link do Google Maps"></label></div>
    <datalist id="dlConstr">${construtores().map(x => `<option value="${esc(x)}">`).join('')}</datalist>
    ${o ? `<p class="muted small">${nQ} quadra(s) e ${nT} lançamento(s) nesta obra. Ao renomear, os lançamentos passam a usar o novo nome.${e.obras.length > 1 && !pode ? ' Para remover, a obra não pode ter quadras nem lançamentos.' : ''}</p>` : '<p class="muted small">Depois de criar, cadastre as quadras e lotes dela em Quadras e lotes, escolhendo esta obra.</p>'}
    <div class="acts">${pode ? '<button type="button" class="btn danger ghost" id="rm">Remover obra</button>' : ''}<button type="button" class="btn" data-close>Cancelar</button><button class="btn primary">Salvar</button></div></form>`);
  const F = f.el.querySelector('#fx');
  F.onsubmit = async ev => {
    ev.preventDefault(); const nome = F.v.value.trim(); const n = clone(E());
    if (!nome) return; if (nome !== o && n.obras.some(x => x.toLowerCase() === nome.toLowerCase())) { toast('Essa obra já existe', 'err'); return; }
    if (!o) n.obras.push(nome);
    else if (nome !== o) {
      n.obras = n.obras.map(x => x === o ? nome : x);
      for (const q in n.quadraObra) if (n.quadraObra[q] === o) n.quadraObra[q] = nome;
      const mud = S.list.filter(x => obraT(x) === o && x.local !== nome).map(x => ({ ...x, local: nome, alterado: Date.now() }));
      if (mud.length) salvarTitulos(mud);
    }
    if (o && nome !== o && n.obraEmpresa[o]) { n.obraEmpresa[nome] = n.obraEmpresa[o]; delete n.obraEmpresa[o]; }
    if (F.emp) n.obraEmpresa[nome] = F.emp.value;
    n.obraInfo = { ...(n.obraInfo || {}) }; if (o && nome !== o) delete n.obraInfo[o];
    n.obraInfo[nome] = { cidade: F.cidade.value.trim(), uf: F.uf.value.trim().toUpperCase(), cep: F.cep.value.trim(), endereco: F.endereco.value.trim(), construtor: F.construtor.value.trim(), matricula: F.matricula.value.trim(), cartorio: F.cartorio.value.trim(), mapa: F.mapa.value.trim() };
    n.local = n.obras[0]; salvarEstrutura(n); f.fechar(); toast(o ? 'Obra salva' : 'Obra criada');
  };
  const rm = f.el.querySelector('#rm'); if (rm) rm.onclick = () => { const n = clone(E()); n.obras = n.obras.filter(x => x !== o); n.local = n.obras[0]; salvarEstrutura(n); f.fechar(); };
};
ACOES.cadLocal = () => { const f = folha('Local da obra', `<form class="form" id="fx"><label>Nome<input name="v" value="${esc(E().local)}" required></label><div class="acts"><button type="button" class="btn" data-close>Cancelar</button><button class="btn primary">Salvar</button></div></form>`); f.el.querySelector('#fx').onsubmit = e => { e.preventDefault(); const n = clone(E()); n.local = e.target.v.value.trim(); salvarEstrutura(n); f.fechar(); }; };
ACOES.cadConta = d => {
  const c = E().contas.find(x => x.nome === d.nome);
  const f = folha(c ? 'Editar conta' : 'Nova conta bancária', `<form class="form" id="fx"><label>Nome<input name="nome" value="${esc(c?.nome || '')}" required ${c ? 'readonly' : ''}></label><label>Saldo inicial (R$)<input name="ini" inputmode="decimal" value="${c ? fNum2.format(c.inicial || 0) : '0,00'}"></label>
    <div class="g3"><label>Banco<input name="banco" value="${esc(c?.banco || '')}" placeholder="Ex.: BTG Pactual (208)"></label><label>Agência<input name="ag" inputmode="numeric" value="${esc(c?.ag || '')}"></label><label>Conta<input name="num" value="${esc(c?.num || '')}" placeholder="00000-0"></label></div>
    <label>Tipo<select name="tipoC">${opts([['corrente', 'Conta corrente'], ['aplicacao', 'Aplicação / investimento'], ['caixa', 'Dinheiro em caixa']], c?.tipoC || 'corrente')}</select></label>
    <label>Empresa (CNPJ) dona da conta<select name="emp">${opts(empOpts(), c ? empConta(c.nome) : matrizId())}</select></label>
    <div class="acts">${c && !usados('conta', c.nome) ? '<button type="button" class="btn danger ghost" id="rm">Remover</button>' : ''}<button type="button" class="btn" data-close>Cancelar</button><button class="btn primary">Salvar</button></div></form>`);
  const F = f.el.querySelector('#fx');
  F.onsubmit = e => { e.preventDefault(); const n = clone(E()); const nome = F.nome.value.trim(); const ini = parseBR(F.ini.value) || 0; const i = n.contas.findIndex(x => x.nome === nome); if (i >= 0) n.contas[i].inicial = ini; else n.contas.push({ nome, inicial: ini }); const cc = n.contas[i >= 0 ? i : n.contas.length - 1]; if (F.emp) cc.emp = F.emp.value; cc.banco = F.banco.value.trim(); cc.ag = F.ag.value.trim(); cc.num = F.num.value.trim(); cc.tipoC = F.tipoC.value; salvarEstrutura(n); f.fechar(); };
  const rm = f.el.querySelector('#rm'); if (rm) rm.onclick = () => { const n = clone(E()); n.contas = n.contas.filter(x => x.nome !== c.nome); salvarEstrutura(n); f.fechar(); };
};
ACOES.cadQuadra = d => {
  const q = d.q; const ls = q ? lotesDe(q) : [];
  const f = folha(q ? 'Quadra ' + q : 'Nova quadra', `<form class="form" id="fx"><label>Nome da quadra<input name="q" value="${esc(q || '')}" required ${q ? 'readonly' : ''} placeholder="Ex.: K4"></label>
    ${variasObras() ? `<label>Obra<select name="obra">${opts(E().obras, q ? obraDe(q) : (d.o || obraVista || E().obras[0]))}</select></label>` : ''}
    <label>Lotes <span class="muted">— use vírgulas e intervalos: 1-45, 19.1, 19.2</span><textarea name="l" rows="4">${esc(ls.join(', '))}</textarea></label>
    <div class="acts">${q && !usados('quadra', q) ? '<button type="button" class="btn danger ghost" id="rm">Remover quadra</button>' : ''}<button type="button" class="btn" data-close>Cancelar</button><button class="btn primary">Salvar</button></div></form>`);
  const F = f.el.querySelector('#fx');
  F.onsubmit = e => { e.preventDefault(); const n = clone(E()); const nome = F.q.value.trim(); if (!q && n.quadras.includes(nome)) { toast(variasObras() ? `A quadra ${nome} já existe na obra ${obraDe(nome)}. Use outro nome, por exemplo ${nome}-${(F.obra?.value || '').slice(0, 3).toUpperCase()}.` : 'Essa quadra já existe', 'err'); return; }
    if (F.obra) n.quadraObra[nome] = F.obra.value; if (!n.quadras.includes(nome)) n.quadras.push(nome); n.lotes[nome] = expandirLotes(F.l.value); salvarEstrutura(n); f.fechar(); toast('Quadra salva'); };
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
    <div class="g2"><label>Corretor <span class="muted">(vazio = venda direta)</span><input name="corretor" value="${esc(v.corretor || '')}" list="dlCorretor" placeholder="Nome do corretor"></label><label>Imobiliária / CRECI<input name="imob" value="${esc(v.imob || '')}" list="dlImob"></label></div>
    <div class="g2"><label>Correspondente bancário<input name="corresp" value="${esc(v.corresp || '')}" list="dlCorr"></label><label>Agência do financiamento<input name="agencia" value="${esc(v.agencia || '')}" list="dlAg" placeholder="Ex.: Caixa ag. 0812 Petrolina"></label></div>
    <label>Observação<textarea name="obs" rows="2">${esc(v.obs || '')}</textarea></label>
    ${['corresp:dlCorr', 'corretor:dlCorretor', 'imob:dlImob', 'agencia:dlAg'].map(s => { const [c, id] = s.split(':'); return `<datalist id="${id}">${[...new Set(Object.values((S.vendas && S.vendas.lotes) || {}).map(o => o[c]).filter(Boolean))].map(x => `<option value="${esc(x)}">`).join('')}</datalist>`; }).join('')}
    ${hist.length ? `<div class="lbl">Histórico</div><div class="tlist small">${hist.map(hh => `<div class="kv"><span>${dBR(hh.data)}</span><b>${esc(etapaNome(hh.etapa))}</b></div>`).join('')}</div>` : ''}
    <div class="acts left rw"><button type="button" class="btn" id="vPl">${v.plano ? 'Plano de pagamento' : 'Criar plano de pagamento'}</button><button type="button" class="btn" id="vPo">${v.portal ? 'Link do cliente' : 'Gerar link para o cliente'}</button></div>
    <div class="acts"><a class="btn" href="#/obras/${encodeURIComponent(q)}/${encodeURIComponent(l)}" data-close>Ver lançamentos do lote</a><button type="button" class="btn" data-close>Cancelar</button><button class="btn primary rw">Salvar</button></div></form>`);
  const F = f.el.querySelector('#fV');
  const guardar = () => {
    const novo = { ...v, etapa: F.etapa.value, data: F.data.value, cliente: F.cliente.value.trim(), tel: F.tel.value.trim(), corresp: F.corresp.value.trim(), corretor: F.corretor.value.trim(), imob: F.imob.value.trim(), agencia: F.agencia.value.trim(), obs: F.obs.value.trim(), hist: v.hist || [] };
    if (novo.etapa !== v.etapa) novo.hist = [...novo.hist, { etapa: novo.etapa, data: novo.data }];
    S.vendas = S.vendas || { lotes: {} }; S.vendas.lotes = { ...(S.vendas.lotes || {}), [k]: novo };
    gravar(doc(db, 'config', 'vendas'), { lotes: { [k]: novo } }, { merge: true }).catch(onErroEscrita);
  };
  F.querySelector('#vPl').onclick = () => { if (!somenteLeitura()) guardar(); f.fechar(); formPlano(k); };
  F.querySelector('#vPo').onclick = () => { if (!somenteLeitura()) guardar(); f.fechar(); ACOES.portalLote({ k }); };
  F.onsubmit = e => { e.preventDefault(); guardar(); f.fechar(); toast('Lote ' + q + ' · ' + l + ' atualizado'); refreshView(); };
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
    return { acct: tag(txt, 'ACCTID'), linhas, saldo: bal ? parseFloat(bal[1].replace(',', '.')) : null, saldoData: dta ? dta[1].slice(0, 4) + '-' + dta[1].slice(4, 6) + '-' + dta[1].slice(6, 8) : '' };
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
  if (!linhas.length) return null;
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
    const rej = new Set(l.rej || []); const okT = t => !usado.has(t.id) && !rej.has(t.id);
    let c = livres.find(t => okT(t) && pago(t) && (t.conta || '') === conta && Math.abs(val(t) - l.v) < 0.01 && dias(t.pag || t.venc, l.d) <= 5);
    if (c) { usado.add(c.id); return { ...l, st: 'auto', ts: [c] }; }
    for (const g of grupos.values()) if (g.length > 1 && pago(g[0]) && (g[0].conta || '') === conta && !g.some(t => usado.has(t.id) || rej.has(t.id)) && Math.abs(sum(g, val) - l.v) < 0.01 && dias(g[0].pag || g[0].venc, l.d) <= 5) { g.forEach(t => usado.add(t.id)); return { ...l, st: 'auto', ts: g }; }
    c = livres.find(t => okT(t) && !pago(t) && Math.abs(t.valor - l.v) < 0.01 && dias(t.venc, l.d) <= 25);
    if (c) { usado.add(c.id); return { ...l, st: 'sug', ts: [c] }; }
    for (const g of grupos.values()) if (g.length > 1 && !pago(g[0]) && !g.some(t => usado.has(t.id) || rej.has(t.id)) && Math.abs(sum(g, t => t.valor) - l.v) < 0.01 && dias(g[0].venc, l.d) <= 25) { g.forEach(t => usado.add(t.id)); return { ...l, st: 'sug', ts: g }; }
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
      <label class="tgl"><input type="checkbox" id="cTodas" ${concTodas ? 'checked' : ''}> Mostrar já conciliadas e ignoradas</label>
      ${a && a.linhas.length ? '<button class="btn ghost rw" data-act="cLimpar">Limpar extrato</button>' : ''}</div>
    ${!a ? `<div class="card empty"><h2>Nenhum extrato importado para ${esc(concConta)}</h2><p class="muted">No app ou internet banking do banco, exporte o extrato do período em <b>OFX</b> (ou CSV) e importe aqui. O sistema casa cada linha com os lançamentos, inclusive rateios por lote.</p></div>` : `
    <div class="grid kpis k4">${kpi('Saldo no banco', a.ex.saldo != null ? brl2(a.ex.saldo) : '—', a.ex.saldoData ? 'em ' + dBR(a.ex.saldoData) : 'o arquivo não trouxe saldo', '')}${kpi('Saldo no sistema', brl2(sd ? sd.saldo : 0), esc(concConta), '')}${kpi('Conciliadas', ls.filter(l => l.st === 'ok').length + ' de ' + ls.filter(l => l.st !== 'ign').length, '', 'pos')}${kpi('Precisam de você', pend.length, ls.filter(l => l.st === 'sug').length + ' sugestão(ões) · ' + ls.filter(l => l.st === 'novo').length + ' sem lançamento', pend.length ? 'neg' : '')}</div>
    ${a.ex.saldo != null && sd && Math.abs(a.ex.saldo - sd.saldo) > 0.01 && !pend.length ? `<div class="alert">Tudo conciliado, mas o saldo do banco difere do sistema em ${brl2(a.ex.saldo - sd.saldo)}. Pode haver lançamentos anteriores ao período do extrato ou saldo inicial da conta a ajustar em Cadastros.</div>` : ''}
    <div class="card"><div class="tbl-wrap"><table class="tbl ftbl"><thead><tr><th class="l">Data</th><th class="l">No banco</th><th>Valor</th><th class="l">Situação</th><th class="l">No sistema</th><th></th></tr></thead><tbody>
    ${vis.map(l => { const [n, c] = ST_CONC[l.st]; const t0 = l.ts[0]; return `<tr><td class="l">${dBR(l.d)}</td><td class="l wrap">${esc(l.desc)}</td><td class="${cls(l.v)}"><b>${brl2(l.v)}</b></td><td class="l st"><span class="chip ${c}">${n}</span></td>
      <td class="l wrap">${t0 ? `<a href="#" data-act="abrir" data-id="${t0.id}">${esc(l.ts.length > 1 ? baseItem(t0) + ' · ' + l.ts.length + ' lotes' : t0.item)}</a><br><span class="muted small">${pago(t0) ? 'Pago ' + dBR(t0.pag) : 'Em aberto · vence ' + dBR(t0.venc)}</span>` : l.st === 'novo' ? '<span class="muted">Nenhum lançamento com este valor e data</span>' : ''}</td>
      <td class="rw conc-acts">${l.st === 'auto' ? `<button class="btn sm" data-act="cOk" data-id="${l.id}">Confirmar</button>${botoesConcExtra(l)}` : l.st === 'sug' ? `<button class="btn sm primary" data-act="cBx" data-id="${l.id}">Dar baixa</button>${botoesConcExtra(l)}` : l.st === 'novo' ? `<button class="btn sm" data-act="cNovo" data-id="${l.id}">Criar lançamento</button> <button class="btn sm" data-act="cEsc" data-id="${l.id}">Escolher lançamento</button> <button class="btn sm link" data-act="cIgn" data-id="${l.id}">Ignorar</button>` : l.st === 'ok' ? `<button class="btn sm link" data-act="cDes" data-id="${l.id}">Desfazer</button>` : `<button class="btn sm link" data-act="cIgn" data-id="${l.id}" data-v="0">Reativar</button>`}</td></tr>`; }).join('') || '<tr><td colspan="6" class="l muted">Nada pendente. Tudo conciliado.</td></tr>'}
    </tbody></table></div></div>`}`;
    $('#cc', el).onchange = e => { concConta = e.target.value; this.refresh(el); };
    $('#cTodas', el).onchange = e => { concTodas = e.target.checked; this.refresh(el); };
    const fi = $('#cfile', el); if (fi) fi.onchange = async e => {
      const file = e.target.files[0]; if (!file) return;
      try {
        const buf = new Uint8Array(await file.arrayBuffer());
        let txt = new TextDecoder('utf-8').decode(buf); if (txt.includes('\uFFFD')) txt = new TextDecoder('windows-1252').decode(buf);
        const r = lerExtratoBanco(txt); if (!r.linhas.length) { toast('Não encontrei lançamentos nesse arquivo.', 'err'); return; }
        if (r.acct) { const dig = soDig(r.acct); const achou = E().contas.find(c => c.num && dig && (soDig(c.num).endsWith(dig.slice(-5)) || dig.endsWith(soDig(c.num).slice(-5)))); if (achou && achou.nome !== concConta && await confirmar(`Este extrato é da conta ${esc(r.acct)}, cadastrada como <b>${esc(achou.nome)}</b>. Importar nela em vez de ${esc(concConta)}?`, 'Usar ' + achou.nome, false)) { concConta = achou.nome; } }
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
ACOES.cNovo = d => { const l = linhaConc(d.id); if (!l) return; formNovo({ emp: variasEmpresas() ? empConta(concConta) : '', tipo: l.v < 0 ? 'saida' : 'entrada', realizado: true, item: l.desc, valor: Math.abs(l.v), venc: l.d, pag: l.d, conta: concConta, conc: l.id, destino: l.v < 0 ? 'Administrativo' : 'Aporte', quadra: '-' }); };
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
    el.innerHTML = `<div class="card"><div class="legend"><span style="--c:var(--bal)">Físico (medição na obra)</span><span style="--c:var(--steel)">Financeiro (gasto ÷ orçado)</span><span style="--c:#1F7A57">Repasse previsto</span></div>
    <div class="tbl-wrap"><table class="tbl"><thead><tr><th class="l">Quadra</th><th>Lotes</th><th class="l">Etapa</th><th class="l">Físico × financeiro</th><th>Comprometido</th><th class="l"><div class="gant" style="--n:${meses.length}">${meses.map(m => `<span>${MES3[+m.slice(5, 7) - 1]}</span>`).join('')}</div></th><th></th></tr></thead><tbody>
    ${rows.map(r => `<tr><td class="l"><b>${esc(r.q)}</b></td><td>${r.n}</td>
      <td class="l"><span class="chip ${r.entregue ? 'ok' : ''}">${esc(r.o.etapa || (r.entregue ? 'Entregue' : 'Não informada'))}</span>${r.alerta ? '<br><span class="chip warn" style="margin-top:4px">Gasto à frente da obra</span>' : ''}</td>
      <td class="l"><div class="dual"><span class="pbar" style="justify-content:flex-start"><span class="bar"><i style="width:${r.o.fis !== null ? r.o.fis : (r.entregue ? 100 : 0)}%;background:var(--bal)"></i></span><span>${r.o.fis !== null ? fPct.format(r.o.fis / 100) : r.entregue ? '100%' : '—'}</span></span>
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
const listaEmp = emp => emp ? S.list.filter(t => empT(t) === emp) : S.list;
const contasEmp = emp => E().contas.filter(c => !emp || empConta(c.nome) === emp);
function dadosRelatorio(m, emp) {
  const ini = m + '-01', fim = addDays(addMonths(ini, 1), -1), h = hoje();
  const L = listaEmp(emp);
  const pagos = L.filter(pago); const base = sum(contasEmp(emp), c => +c.inicial || 0);
  const ate = d => base + sum(pagos.filter(t => (t.pag || t.venc) <= d), t => t.valor + (t.juros || 0));
  const mes = pagos.filter(t => { const d = t.pag || t.venc; return d >= ini && d <= fim; });
  const porDest = {}; for (const t of mes) if (t.valor < 0 || t.juros) { if (t.valor < 0) porDest[t.destino] = (porDest[t.destino] || 0) + t.valor; if (t.juros) porDest['Juros e Multa'] = (porDest['Juros e Multa'] || 0) + t.juros; }
  const ab = L.filter(t => !pago(t)); const pj = emp ? projecao(L, ate(h)) : projecao();
  const fluxo = [...Array(6)].map((_, i) => { const mi = addMonths(h.slice(0, 7) + '-01', i), mf = addDays(addMonths(mi, 1), -1); let v = pj.base; for (const x of pj.dias) if (x.d <= mf) v = x.v; return [mesNome(mi.slice(0, 7)), v]; });
  return {
    m, ini, fim, saldoIni: ate(addDays(ini, -1)), saldoFim: ate(fim),
    entradas: sum(mes.filter(t => t.valor > 0 && t.destino !== 'Aporte'), t => t.valor), aportes: sum(mes.filter(t => t.destino === 'Aporte'), t => t.valor),
    saidas: sum(mes, t => Math.min(0, t.valor) + Math.min(0, t.juros || 0)),
    porDest: Object.entries(porDest).sort((a, b) => a[1] - b[1]),
    quadras: E().quadras.filter(q => !emp || empDaObra(obraDe(q)) === emp).map(resumoQuadra),
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
    if (relEmp && !empPorId(relEmp)) relEmp = '';
    const r = dadosRelatorio(relMes, relEmp);
    el.innerHTML = `<div class="filtros card"><label class="tgl">Mês <select id="rM">${opts(ms.map(m => [m, mesNome(m)]), relMes)}</select></label>${variasEmpresas() ? `<label class="tgl">Empresa <select id="rE">${opts(empOpts(), relEmp, 'Grupo consolidado')}</select></label>` : ''}<button class="btn primary" id="rPdf">Baixar PDF</button><span class="muted small">O PDF traz o resumo do mês, despesas por destino, resultado por quadra, contas vencidas, próximos 30 dias e o fluxo projetado.</span></div>
    <div class="grid kpis k4">${kpi('Saldo inicial', brl(r.saldoIni), dBR(addDays(r.ini, -1)), cls(r.saldoIni))}${kpi('Entradas', brl(r.entradas), r.aportes ? 'aportes ' + brl(r.aportes) : '', 'pos')}${kpi('Saídas', brl(r.saidas), 'inclui juros', 'neg')}${kpi('Saldo final', brl(r.saldoFim), dBR(r.fim), cls(r.saldoFim))}</div>
    <div class="grid two"><div class="card"><h2>Saídas por destino</h2><table class="tbl"><tbody>${r.porDest.map(([d, v]) => `<tr><td class="l">${esc(d)}</td><td class="neg">${brl(v)}</td></tr>`).join('') || '<tr><td class="l muted">Sem saídas no mês</td></tr>'}</tbody></table></div>
    <div class="card"><h2>Fluxo projetado</h2><table class="tbl"><tbody>${r.fluxo.map(([n, v]) => `<tr><td class="l">${esc(n)}</td><td class="${cls(v)}">${brl(v)}</td></tr>`).join('')}</tbody></table></div></div>`;
    $('#rM', el).onchange = e => { relMes = e.target.value; this.refresh(el); };
    const rE = $('#rE', el); if (rE) rE.onchange = e => { relEmp = e.target.value; this.refresh(el); };
    $('#rPdf', el).onclick = () => gerarPDF(relMes, relEmp).catch(e => { console.error(e); toast('Erro ao gerar PDF: ' + e.message, 'err'); });
  }
};
async function imgParaPNG(url, w) {
  const img = await new Promise((res, rej) => { const i = new Image(); i.onload = () => res(i); i.onerror = rej; i.src = url; });
  const r = (img.naturalHeight || 508) / (img.naturalWidth || 908); const c = document.createElement('canvas'); c.width = w; c.height = Math.round(w * r);
  c.getContext('2d').drawImage(img, 0, 0, c.width, c.height); return { data: c.toDataURL('image/png'), r };
}
let relEmp = '';
const nomeEmpPDF = emp => { const x = emp && empPorId(emp); return x ? x.nome + (x.cnpj ? ' · CNPJ ' + fCnpj(x.cnpj) : '') : (variasEmpresas() ? 'Grupo consolidado' : 'SFG Incorporadora'); };
async function gerarPDF(m, emp) {
  toast('Gerando PDF…'); await carregarLib('pdf');
  const r = dadosRelatorio(m, emp);
  const P = s => String(s).replace(/\u2212/g, '-').replace(/[\u202f\u00a0]/g, ' ');
  const pdf = new jsPDF({ unit: 'pt', format: 'a4' }); const W = pdf.internal.pageSize.getWidth();
  const NAVY = [11, 32, 70];
  try { const lg = await imgParaPNG(LOGO_COMPLETO, 480); pdf.addImage(lg.data, 'PNG', 36, 16, 96, 96 * lg.r); } catch (e) { /* sem logo */ }
  pdf.setDrawColor(...NAVY); pdf.setLineWidth(1.2); pdf.line(36, 92, W - 36, 92);
  pdf.setTextColor(...NAVY); pdf.setFont('helvetica', 'bold'); pdf.setFontSize(17); pdf.text('Relatório financeiro', W - 36, 40, { align: 'right' });
  pdf.setFont('helvetica', 'normal'); pdf.setFontSize(11); pdf.text(P(`${mesNome(m)} · ${nomeEmpPDF(emp)}`), W - 36, 58, { align: 'right' });
  pdf.setFontSize(8.5); pdf.setTextColor(91, 107, 130); pdf.text(P('SFG Incorporadora · gerado em ' + new Date().toLocaleString('pt-BR', { dateStyle: 'short', timeStyle: 'short' })), W - 36, 74, { align: 'right' });
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
  await baixarArquivo(pdf.output('blob'), `SFG Relatório ${m}${emp ? ' ' + (empPorId(emp)?.nome || '') : ''}.pdf`);
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
    // Link de acompanhamento do cliente: qualquer pessoa com o código abre só aquele lote (sem listar os outros).
    match /portal/{id} { allow get: if true; allow list: if leitor(); allow write: if editor(); }
    match /{col}/{id} {
      allow read: if leitor();
      allow write: if editor() && !(col == 'config' && id == 'acesso');
    }
  }
}`;
let filtroPendente = null, empPer = 'ano';
ACOES.verEmp = d => { const conf = d.id.startsWith('__'); filtroPendente = { emp: d.id, ...(d.modo === 'extrato' ? { per: conf || empPer !== 'ano' ? 'todos' : 'ano' } : { st: conf ? 'todos' : 'aberto', per: 'todos' }) }; location.hash = '#/' + d.modo; };
ACOES.empPer = d => { empPer = d.v; refreshView(); };
function resumoEmpresa(id) {
  const h = hoje(), ano = h.slice(0, 4);
  const ts = S.list.filter(t => empT(t) === id);
  const noPer = t => empPer !== 'ano' || (dataRef(t) || '').startsWith(ano);
  const real = ts.filter(t => pago(t) && noPer(t)), ab = ts.filter(t => !pago(t));
  const ap = t => t.destino === 'Aporte';
  const r = {
    id, n: ts.length,
    recebido: sum(real.filter(t => !ap(t) && t.valor > 0), t => t.valor),
    pago: sum(real.filter(t => !ap(t) && t.valor < 0), t => t.valor + Math.min(t.juros || 0, 0)),
    aportes: sum(real.filter(ap), t => t.valor),
    receber: sum(ab.filter(t => !ap(t) && t.valor > 0), t => t.valor),
    pagar: sum(ab.filter(t => !ap(t) && t.valor < 0), t => t.valor),
    vencido: sum(ab.filter(t => t.valor < 0 && t.venc < h), t => t.valor),
    contas: saldos().contas.filter(c => empConta(c.nome) === id)
  };
  r.div = ts.filter(empDiverge).length;
  r.resultado = r.recebido + r.pago; r.projetado = r.resultado + r.receber + r.pagar; r.saldo = sum(r.contas, c => c.saldo);
  return r;
}
VIEWS.empresas = {
  titulo: 'Empresas',
  refresh(el) {
    const e = E(); const rs = e.empresas.map(x => ({ x, ...resumoEmpresa(x.id) }));
    const T = k => sum(rs, r => r[k]);
    const linha = (lab, v, extra = '') => `<div class="kv"><span>${lab}</span><b class="fig ${cls(v)}">${brl2(v)}</b>${extra}</div>`;
    const nSem = S.list.filter(t => !empDef(t)).length, nDiv = S.list.filter(empDiverge).length;
    const conf = e.empresas.length > 1 && (nSem || nDiv) ? `<div class="card conf-cnpj"><h2>Conferência de CNPJ</h2>
      ${nDiv ? `<div class="kv"><span><b>${nDiv}</b> lançamento(s) em CNPJ diferente do CNPJ da obra</span><button class="btn sm" data-act="verEmp" data-id="__div" data-modo="pagar">A pagar</button><button class="btn sm" data-act="verEmp" data-id="__div" data-modo="extrato">Extrato</button></div>` : ''}
      ${nSem ? `<div class="kv"><span><b>${nSem}</b> lançamento(s) sem CNPJ definido (lançados antes das filiais). Abra, selecione e use Definir CNPJ.</span><button class="btn sm" data-act="verEmp" data-id="__sem" data-modo="pagar">A pagar</button><button class="btn sm" data-act="verEmp" data-id="__sem" data-modo="receber">A receber</button><button class="btn sm" data-act="verEmp" data-id="__sem" data-modo="extrato">Extrato</button></div>` : ''}
      <p class="muted small">Enquanto o CNPJ não é definido, o lançamento conta na empresa da obra.</p></div>` : '';
    el.innerHTML = `${conf}${e.empresas.length < 2 ? `<div class="alerta-caixa"><span>Só a matriz está cadastrada. Cadastre as filiais do RET em Cadastros, em Empresas (CNPJ), e escolha a empresa de cada obra.</span><a class="btn sm" href="#/cadastros">Ir para Cadastros</a></div>` : ''}
      <div class="seg">${[['ano', 'Realizado em ' + hoje().slice(0, 4)], ['todos', 'Realizado desde o início']].map(([v, l]) => `<button class="${v === empPer ? 'on' : ''}" data-act="empPer" data-v="${v}">${l}</button>`).join('')}</div>
      <div class="grid two emp-grid">${rs.map(r => `<div class="card emp-card">
        <h2><span>${esc(r.x.nome)}<small class="muted fig">${esc(fCnpj(r.x.cnpj)) || 'CNPJ não informado'}</small></span><span class="chip ${r.x.tipo === 'matriz' ? 'ok' : ''}">${r.x.tipo === 'matriz' ? 'Matriz' : 'Filial'}${r.x.ret ? ' · RET' : ''}</span></h2>
        <p class="muted small">${esc(e.obras.filter(o => empDaObra(o) === r.id).join(', ')) || 'Nenhuma obra vinculada'}${r.contas.length ? ' · contas: ' + esc(r.contas.map(c => c.nome).join(', ')) : ''}</p>
        ${linha('Saldo nas contas', r.saldo)}
        ${linha('Recebido', r.recebido)}${linha('Pago', r.pago)}
        ${linha('Resultado realizado', r.resultado)}
        ${r.aportes ? linha('Aportes / transferências', r.aportes) : ''}
        ${linha('A receber', r.receber)}${linha('A pagar', r.pagar, r.vencido ? `<small class="neg">${brl2(r.vencido)} vencido</small>` : '')}
        ${linha('Resultado projetado', r.projetado)}
        ${r.div ? `<p class="small" style="color:var(--warn)">${r.div} lançamento(s) desta empresa são de obra de outro CNPJ.</p>` : ''}
        <div class="acts" style="justify-content:flex-start"><button class="btn sm" data-act="verEmp" data-id="${esc(r.id)}" data-modo="pagar">Contas a pagar</button><button class="btn sm" data-act="verEmp" data-id="${esc(r.id)}" data-modo="receber">A receber</button><button class="btn sm" data-act="verEmp" data-id="${esc(r.id)}" data-modo="extrato">Extrato</button></div>
      </div>`).join('')}</div>
      ${rs.length > 1 ? `<div class="card"><h2>Consolidado</h2><div class="tbl-wrap"><table class="tbl"><thead><tr><th class="l">Empresa</th><th>Saldo contas</th><th>Recebido</th><th>Pago</th><th>Resultado</th><th>A receber</th><th>A pagar</th><th>Projetado</th></tr></thead><tbody>
        ${rs.map(r => `<tr><td class="l"><b>${esc(r.x.nome)}</b><br><span class="muted small">${esc(fCnpj(r.x.cnpj))}</span></td><td class="${cls(r.saldo)}">${brl(r.saldo)}</td><td class="pos">${brl(r.recebido)}</td><td class="neg">${brl(r.pago)}</td><td class="${cls(r.resultado)}"><b>${brl(r.resultado)}</b></td><td>${brl(r.receber)}</td><td>${brl(r.pagar)}</td><td class="${cls(r.projetado)}"><b>${brl(r.projetado)}</b></td></tr>`).join('')}
        <tr class="tot"><td class="l"><b>Total do grupo</b></td><td class="${cls(T('saldo'))}">${brl(T('saldo'))}</td><td>${brl(T('recebido'))}</td><td>${brl(T('pago'))}</td><td><b>${brl(T('resultado'))}</b></td><td>${brl(T('receber'))}</td><td>${brl(T('pagar'))}</td><td><b>${brl(T('projetado'))}</b></td></tr></tbody></table></div>
        <p class="muted small">Aportes e transferências entre CNPJs (destino Aporte) ficam fora do resultado, para não contar a mesma receita duas vezes.</p></div>` : ''}`;
  }
};
/* ---------- Histórico financeiro: por destino (Fluxo Bancário) e por quadra (Cronograma) ---------- */
let hfAba = 'destino', hfIni = '', hfFim = '', hfProj = false, hfEmp = '', hfObra = '';
const mesCurto = m => MES3[+m.slice(5, 7) - 1] + ' ' + m.slice(2, 4);
function mesesEntre(a, b) { const out = []; let m = a; for (let k = 0; m <= b && k < 240; k++) { out.push(m); m = addMonths(m + '-01', 1).slice(0, 7); } return out; }
function dadosHistorico() {
  const h = hoje(), mesAtual = h.slice(0, 7);
  let L = listaEmp(hfEmp); if (hfObra) L = L.filter(t => obraT(t) === hfObra);
  const real = L.filter(pago), ab = hfProj ? L.filter(t => !pago(t)) : [];
  const mesDe = t => pago(t) ? (t.pag || t.venc || '').slice(0, 7) : (t.venc < h ? mesAtual : t.venc.slice(0, 7));
  const meses0 = real.map(mesDe).filter(m => /^\d{4}-\d{2}$/.test(m)).sort();
  if (!hfIni) hfIni = meses0[0] || mesAtual; if (!hfFim) hfFim = mesAtual;
  let fim = hfFim; if (hfProj) { const ult = ab.map(mesDe).sort().pop(); const lim = addMonths(mesAtual + '-01', 12).slice(0, 7); if (ult && ult > fim) fim = ult > lim ? lim : ult; }
  const meses = mesesEntre(hfIni, fim);
  const base = sum(hfObra ? [] : contasEmp(hfEmp), c => +c.inicial || 0) + sum(real.filter(t => mesDe(t) < hfIni), t => t.valor + (t.juros || 0));
  const doPer = [...real, ...ab].filter(t => { const m = mesDe(t); return m >= hfIni && m <= fim; });
  return { meses, base, doPer, mesDe, mesAtual };
}
function tabelaHistorico() {
  const { meses, base, doPer, mesDe, mesAtual } = dadosHistorico();
  const idx = Object.fromEntries(meses.map((m, i) => [m, i])); const N = meses.length;
  const proj = m => hfProj && m > mesAtual;
  if (hfAba === 'destino') {
    const por = {}; const add = (k, i, v) => { (por[k] = por[k] || Array(N).fill(0))[i] += v; };
    for (const t of doPer) { const i = idx[mesDe(t)]; if (i === undefined) continue; add(t.destino || 'Sem destino', i, t.valor || 0); if (t.juros && pago(t)) add('Juros e Multa', i, t.juros); }
    const sec = d => d === 'Juros e Multa' ? 'nonop' : secaoDe(d);
    const ent = Object.keys(por).filter(d => ['rec', 'aporte'].includes(sec(d))).sort((a, b) => (sec(b) === 'aporte') - (sec(a) === 'aporte') || a.localeCompare(b));
    const sai = Object.keys(por).filter(d => !ent.includes(d)).sort((a, b) => a.localeCompare(b));
    const soma = ks => meses.map((_, i) => sum(ks, k => por[k][i]));
    const tE = soma(ent), tS = soma(sai), res = meses.map((_, i) => tE[i] + tS[i]);
    let ac = base; const acum = res.map(v => (ac += v));
    const T = a => sum(a, x => x);
    const rows = [['Entradas', tE, 'total'], ...ent.map(d => ['  ' + d, por[d], 'item']), ['Saídas', tS, 'total'], ...sai.map(d => ['  ' + d, por[d], 'item']), ['Resultado do mês', res, 'sub'], ['Saldo acumulado', acum, 'grand']];
    return { tipo: 'destino', cols: meses.map(mesCurto), proj: meses.map(proj), rows: rows.map(([n, vs, s]) => [n, s === 'grand' ? null : T(vs), vs, s]), base };
  }
  // por quadra: linhas = meses, colunas = quadras (entradas e saídas)
  const qs = (hfObra ? quadrasDaObra(hfObra) : E().quadras).filter(q => doPer.some(t => t.quadra === q));
  const z = () => Array(N).fill(0);
  const Q = Object.fromEntries(qs.map(q => [q, { e: z(), s: z() }])); const geral = z();
  for (const t of doPer) {
    const i = idx[mesDe(t)]; if (i === undefined) continue; const v = (t.valor || 0) + (pago(t) ? (t.juros || 0) : 0);
    const c = Q[t.quadra]; if (!c) { geral[i] += v; continue; }
    if (t.valor > 0) c.e[i] += t.valor; else c.s[i] += t.valor; if (pago(t) && t.juros) c.s[i] += t.juros;
  }
  const E_ = meses.map((_, i) => sum(qs, q => Q[q].e[i])), S_ = meses.map((_, i) => sum(qs, q => Q[q].s[i]));
  let ac = base; const acum = meses.map((_, i) => (ac += E_[i] + geral[i] + S_[i]));
  return { tipo: 'quadra', meses, proj: meses.map(proj), qs, Q, E_, S_, geral, acum, base };
}
function htmlHistorico(d) {
  const f = v => Math.abs(v) < 0.5 ? '<span class="muted">–</span>' : brl(v);
  if (d.tipo === 'destino') {
    return `<div class="tbl-wrap"><table class="tbl dem hist-tab"><thead><tr><th class="l">R$</th><th>Consolidado</th>${d.cols.map((c, i) => `<th class="${d.proj[i] ? 'proj' : ''}">${esc(c)}${d.proj[i] ? '<small>projetado</small>' : ''}</th>`).join('')}</tr></thead><tbody>
      ${d.rows.map(([n, tot, vs, s]) => `<tr class="dem-${s}"><td class="l">${esc(n.trim())}</td><td class="dem-tot ${tot === null ? '' : cls(tot)}">${tot === null ? '' : f(tot)}</td>${vs.map((v, i) => `<td class="${cls(v)} ${d.proj[i] ? 'proj' : ''}">${f(v)}</td>`).join('')}</tr>`).join('')}</tbody></table></div>`;
  }
  const porObra = !hfObra && variasObras();
  const grupos = porObra ? E().obras.map(o => [o, d.qs.filter(q => obraDe(q) === o)]).filter(x => x[1].length) : [['', d.qs]];
  const qOrd = grupos.flatMap(g => g[1]);
  const tot = a => sum(a, x => x);
  return `<div class="tbl-wrap"><table class="tbl dem hist-tab crono"><thead>
    ${porObra ? `<tr><th class="l"></th><th colspan="4" class="grp">Consolidado</th>${grupos.map(([o, qs]) => `<th colspan="${qs.length * 2}" class="grp">${esc(o)}</th>`).join('')}</tr>` : ''}
    <tr><th class="l"></th><th colspan="4" class="grp">${porObra ? '' : 'Consolidado'}</th>${qOrd.map(q => `<th colspan="2" class="grp q">${esc(q)}</th>`).join('')}</tr>
    <tr><th class="l">Mês</th><th>Entradas</th><th title="Lançamentos sem quadra: aportes, administrativo e gerais">Geral / aportes</th><th>Saídas</th><th>Saldo acumulado</th>${qOrd.map(() => '<th>Entradas</th><th>Saídas</th>').join('')}</tr></thead><tbody>
    ${d.meses.map((m, i) => `<tr class="${d.proj[i] ? 'proj' : ''}"><td class="l">${mesCurto(m)}${d.proj[i] ? ' <small class="muted">proj.</small>' : ''}</td><td class="pos">${f(d.E_[i])}</td><td class="${cls(d.geral[i])}">${f(d.geral[i])}</td><td class="neg">${f(d.S_[i])}</td><td class="dem-tot ${cls(d.acum[i])}"><b>${brl(d.acum[i])}</b></td>${qOrd.map(q => `<td class="pos">${f(d.Q[q].e[i])}</td><td class="neg">${f(d.Q[q].s[i])}</td>`).join('')}</tr>`).join('')}
    <tr class="dem-grand"><td class="l">Total</td><td>${f(tot(d.E_))}</td><td>${f(tot(d.geral))}</td><td>${f(tot(d.S_))}</td><td class="dem-tot">${brl(d.acum[d.acum.length - 1] ?? d.base)}</td>${qOrd.map(q => `<td>${f(tot(d.Q[q].e))}</td><td>${f(tot(d.Q[q].s))}</td>`).join('')}</tr>
    <tr class="dem-sub"><td class="l">Resultado da quadra</td><td colspan="4"></td>${qOrd.map(q => { const r = tot(d.Q[q].e) + tot(d.Q[q].s); return `<td colspan="2" class="${cls(r)}">${f(r)}</td>`; }).join('')}</tr>
    </tbody></table></div>`;
}
VIEWS.movimento = {
  titulo: 'Histórico financeiro',
  refresh(el) {
    if (hfEmp && !empPorId(hfEmp)) hfEmp = ''; if (hfObra && !E().obras.includes(hfObra)) hfObra = '';
    const d = tabelaHistorico();
    const todos = S.list.filter(pago).map(t => (t.pag || t.venc || '').slice(0, 7)).filter(m => /^\d{4}-\d{2}$/.test(m)).sort();
    const opMes = mesesEntre(todos[0] || hoje().slice(0, 7), hoje().slice(0, 7)).reverse().map(m => [m, mesCurto(m)]);
    el.innerHTML = `<div class="filtros card">
      ${variasEmpresas() ? `<label class="tgl">Empresa <select id="hE">${opts(empOpts(), hfEmp, 'Grupo consolidado')}</select></label>` : ''}
      ${variasObras() ? `<label class="tgl">Obra <select id="hO">${opts(E().obras, hfObra, 'Todas as obras')}</select></label>` : ''}
      <label class="tgl">De <select id="hI">${opts(opMes.slice().reverse(), hfIni)}</select></label><label class="tgl">até <select id="hF">${opts(opMes, hfFim)}</select></label>
      <label class="tgl"><input type="checkbox" id="hP" ${hfProj ? 'checked' : ''}> Incluir projeção (contas em aberto)</label>
      <button class="btn" id="hX">Exportar Excel</button></div>
      <div class="seg">${[['destino', 'Por destino (Fluxo bancário)'], ['quadra', 'Por quadra (Cronograma)']].map(([v, l]) => `<button class="${v === hfAba ? 'on' : ''}" data-act="hfAba" data-v="${v}">${l}</button>`).join('')}</div>
      <div class="card">${htmlHistorico(d)}
      <p class="muted small">Realizado = o que foi pago ou recebido, no mês do pagamento, como nas abas Fluxo Bancário e Cronograma da planilha.${hfProj ? ' Meses marcados como projetado trazem as contas em aberto pelo vencimento (as vencidas entram no mês atual).' : ''} O saldo acumulado parte do saldo antes do primeiro mês${hfObra ? ' da obra (sem os saldos iniciais das contas)' : ''}.${d.tipo === 'quadra' ? ' "Geral / aportes" reúne os lançamentos sem quadra.' : ''}</p></div>`;
    const on = (id, fn) => { const x = $('#' + id, el); if (x) x.onchange = e => { fn(e.target.type === 'checkbox' ? e.target.checked : e.target.value); this.refresh(el); }; };
    on('hE', v => hfEmp = v); on('hO', v => hfObra = v); on('hI', v => { hfIni = v; if (hfFim < v) hfFim = v; }); on('hF', v => { hfFim = v; if (hfIni > v) hfIni = v; }); on('hP', v => hfProj = v);
    $('#hX', el).onclick = () => exportarHistorico(d);
  }
};
ACOES.hfAba = d => { hfAba = d.v; refreshView(); };
async function exportarHistorico(d) {
  await carregarLib('xlsx');
  let aoa;
  if (d.tipo === 'destino') aoa = [['R$', 'Consolidado', ...d.cols.map((c, i) => c + (d.proj[i] ? ' (proj.)' : ''))], ...d.rows.map(([n, tot, vs]) => [n.trim(), tot === null ? '' : r2(tot), ...vs.map(r2)])];
  else aoa = [['Mês', 'Entradas', 'Geral / aportes', 'Saídas', 'Saldo acumulado', ...d.qs.flatMap(q => [q + ' entradas', q + ' saídas'])], ...d.meses.map((m, i) => [mesCurto(m) + (d.proj[i] ? ' (proj.)' : ''), r2(d.E_[i]), r2(d.geral[i]), r2(d.S_[i]), r2(d.acum[i]), ...d.qs.flatMap(q => [r2(d.Q[q].e[i]), r2(d.Q[q].s[i])])])];
  const ws = XLSX.utils.aoa_to_sheet(aoa); ws['!cols'] = [{ wch: 24 }, ...aoa[0].slice(1).map(() => ({ wch: 13 }))];
  const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, d.tipo === 'destino' ? 'Fluxo bancário' : 'Cronograma');
  baixarArquivo(new Blob([XLSX.write(wb, { type: 'array', bookType: 'xlsx' })], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), `SFG Histórico ${d.tipo === 'destino' ? 'por destino' : 'por quadra'} ${hfIni} a ${hfFim}.xlsx`);
}

/* ---------- Fluxo do lote: etapas, dependências e tempos ---------- */
const GRUPOS_FX = ['Terreno', 'Obra', 'Venda e repasse'];
const FLUXO_PADRAO = [
  ['t1', 'Escolha do terreno', 'Terreno', [], ''], ['t2', 'Contrato do terreno', 'Terreno', ['t1'], ''], ['t3', 'Declaração', 'Terreno', ['t2'], ''],
  ['t4', 'IPTU', 'Terreno', ['t2'], ''], ['t5', 'CND IPTU', 'Terreno', ['t4'], 'CND IPTU'], ['t6', 'ITBI (antes do último pagamento)', 'Terreno', ['t2'], 'ITBI'],
  ['t7', 'Taxa de escritura', 'Terreno', ['t2'], ''], ['t8', 'Autorização de escritura', 'Terreno', ['t2'], 'Autorização Escritura'], ['t9', 'Certidão (escritura)', 'Terreno', ['t2'], ''],
  ['t10', 'Escritura', 'Terreno', ['t5', 't6', 't7', 't8', 't9'], 'Escritura'], ['t11', 'Registro da escritura', 'Terreno', ['t10'], 'Registro Escritura'],
  ['o1', 'Escolha do construtor', 'Obra', [], ''], ['o2', 'Contrato do construtor', 'Obra', ['o1'], ''], ['o3', 'Projetos', 'Obra', ['o2'], 'Projetos'],
  ['o4', 'ART', 'Obra', ['o3'], 'ART'], ['o5', 'Taxa ART', 'Obra', ['o4'], ''], ['o6', 'Licença de construção', 'Obra', ['o3', 'o5'], 'Licença de Construção'],
  ['o7', 'Habite-se', 'Obra', ['o6'], 'Habite-se'], ['o8', 'Averbação do habite-se', 'Obra', ['o7'], 'Averbação Habite-se'], ['o9', 'ISS', 'Obra', ['o7'], 'ISS'], ['o10', 'SCPO', 'Obra', ['o9'], 'SCPO'],
  ['v0', 'Cliente', 'Venda e repasse', [], ''], ['v1', 'Certidão', 'Venda e repasse', ['t11', 'o8'], 'Certidão Venda'], ['v2', 'Taxa de vistoria', 'Venda e repasse', ['v1'], 'Taxa Vistoria'],
  ['v3', 'Avaliação', 'Venda e repasse', ['v2', 'v0'], ''], ['v4', 'Conformidade', 'Venda e repasse', ['v3'], ''], ['v5', 'Firmas', 'Venda e repasse', ['v4'], 'Firma'], ['v6', 'Buscas', 'Venda e repasse', ['v4'], 'Buscas'],
  ['v7', 'Assinatura do contrato', 'Venda e repasse', ['v5', 'v6'], ''], ['v8', 'Prefeitura', 'Venda e repasse', ['v7'], ''], ['v9', 'Registro', 'Venda e repasse', ['v8'], ''],
  ['v10', 'Digitalização', 'Venda e repasse', ['v9'], ''], ['v11', 'Corretagem', 'Venda e repasse', ['v10'], ''], ['v12', 'Termo de recebimento', 'Venda e repasse', ['v11'], 'Termo de Entrega']
].map(([id, nome, grupo, dep, doc]) => ({ id, nome, grupo, dep, doc, prazo: 0 }));
const etapasFx = () => { try { const e = S.fluxo && S.fluxo.etapas && JSON.parse(S.fluxo.etapas); if (Array.isArray(e) && e.length) return e; } catch (e) { /* padrão */ } return FLUXO_PADRAO; };
const regLote = (q, l) => { try { return JSON.parse((S.fluxoLotes || {})[q + '|' + l] || ((S.fluxo && S.fluxo.lotes) || {})[q + '|' + l] || '{}'); } catch (e) { return {}; } };
const idBloco = s => encodeURIComponent(s).replace(/\./g, '%2E');
const difDias = (a, b) => Math.round((new Date(b + 'T12:00') - new Date(a + 'T12:00')) / 864e5);
function niveisFx(es) {
  const by = Object.fromEntries(es.map(e => [e.id, e])), nv = {};
  const f = (id, pilha = new Set()) => { if (nv[id] !== undefined) return nv[id]; if (pilha.has(id)) return 0; pilha.add(id); const e = by[id]; const ds = (e?.dep || []).filter(d => by[d]); nv[id] = ds.length ? 1 + Math.max(...ds.map(d => f(d, pilha))) : 0; return nv[id]; };
  es.forEach(e => f(e.id)); return nv;
}
function ordemFx(es) { const nv = niveisFx(es); const g = x => { const i = GRUPOS_FX.indexOf(x.grupo); return i < 0 ? 99 : i; }; return es.slice().sort((a, b) => nv[a.id] - nv[b.id] || g(a) - g(b)); }
/* Situação de cada etapa para um lote. */
function estadoFx(q, l) {
  const es = etapasFx(), reg = regLote(q, l), out = {}, h = hoje(); const docs = docsCk();
  for (const e of ordemFx(es)) {
    const r = reg[e.id] || {}; let f = r.f || '', ok = !!r.f, auto = false;
    if (!ok && e.doc) { const i = docs.indexOf(e.doc); if (i >= 0) { const st = estadoDoc(q, l, i); if (st.feito) { ok = true; auto = true; f = st.arq ? st.arq.d : ''; } } }
    const deps = (e.dep || []).map(d => out[d]).filter(Boolean); const depsOk = deps.every(d => d.ok);
    const fimDeps = deps.map(d => d.f).filter(Boolean).sort().pop() || '';
    const ini = r.i || (depsOk && deps.length ? fimDeps : '');
    const dias = ok && f && ini ? Math.max(0, difDias(ini, f)) : null;
    const corre = !ok && r.i ? difDias(r.i, h) : null;
    const st = ok ? 'ok' : r.i ? 'andamento' : depsOk ? 'liberada' : 'bloqueada';
    const atraso = !!(e.prazo && ((dias !== null && dias > e.prazo) || (corre !== null && corre > e.prazo)));
    out[e.id] = { e, ok, f, i: r.i || '', ini, dias, corre, st, atraso, auto, o: r.o || '' };
  }
  return out;
}
const ST_FX = { ok: 'Concluída', andamento: 'Em andamento', liberada: 'Pode começar', bloqueada: 'Aguardando etapas anteriores' };
let fxAba = 'mapa', fxQ = '', fxL = '', fxObra = '';
function lotesFx() { const qs = (fxObra ? quadrasDaObra(fxObra) : E().quadras).filter(q => lotesDe(q).length); return (fxQ ? [fxQ] : qs).flatMap(q => lotesDe(q).map(l => [q, l])); }
const mediana = a => { if (!a.length) return null; const s = a.slice().sort((x, y) => x - y), m = s.length >> 1; return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2; };
function statsFx(lotes) {
  const est = lotes.map(([q, l]) => [q, l, estadoFx(q, l)]); const o = {};
  for (const e of etapasFx()) { const xs = est.map(x => x[2][e.id]); const ds = xs.map(x => x.dias).filter(x => x !== null); o[e.id] = { n: lotes.length, ok: xs.filter(x => x.ok).length, and: xs.filter(x => x.st === 'andamento').length, lib: xs.filter(x => x.st === 'liberada').length, atr: xs.filter(x => x.atraso).length, ds, med: ds.length ? sum(ds, x => x) / ds.length : null, mediana: mediana(ds), min: ds.length ? Math.min(...ds) : null, max: ds.length ? Math.max(...ds) : null }; }
  return { est, o };
}
/* Desenho do fluxograma: colunas pela ordem das dependências, faixas por grupo. */
function svgFx(info, acao) {
  const es = etapasFx(), nv = niveisFx(es); const W = 164, H = 50, GX = 34, GY = 12, LX = 118, PAD = 14;
  const grupos = [...GRUPOS_FX.filter(g => es.some(e => e.grupo === g)), ...[...new Set(es.map(e => e.grupo))].filter(g => !GRUPOS_FX.includes(g))];
  const pos = {}; let y = PAD; const faixas = [];
  for (const g of grupos) {
    const ns = es.filter(e => e.grupo === g); const porNv = {}; ns.forEach(e => (porNv[nv[e.id]] = porNv[nv[e.id]] || []).push(e));
    const linhas = Math.max(1, ...Object.values(porNv).map(a => a.length));
    Object.entries(porNv).forEach(([n, a]) => a.forEach((e, i) => { pos[e.id] = { x: LX + (+n) * (W + GX), y: y + i * (H + GY) }; }));
    const alt = linhas * (H + GY) - GY; faixas.push({ g, y: y - 8, h: alt + 16 }); y += alt + 30;
  }
  const larg = LX + (Math.max(0, ...Object.values(nv)) + 1) * (W + GX) + PAD, alt = y - 14;
  const corta = (s, n) => s.length > n ? s.slice(0, n - 1) + '…' : s;
  const setas = es.flatMap(e => (e.dep || []).filter(d => pos[d]).map(d => { const a = pos[d], b = pos[e.id]; const x1 = a.x + W, y1 = a.y + H / 2, x2 = b.x - 4, y2 = b.y + H / 2, mx = (x1 + x2) / 2; const ok = info(es.find(z => z.id === d))?.cls?.includes('ok'); return `<path class="fx-seta ${ok ? 'ok' : ''}" d="M${x1},${y1} C${mx},${y1} ${mx},${y2} ${x2},${y2}" marker-end="url(#fxSeta)"/>`; })).join('');
  const nos = es.map(e => { const p = pos[e.id], i = info(e) || {}; return `<g class="fx-no ${i.cls || ''}" data-act="${acao}" data-id="${esc(e.id)}" tabindex="0" role="button" aria-label="${esc(e.nome + (i.sub ? ', ' + i.sub : ''))}"><title>${esc(e.nome + (i.sub ? ' — ' + i.sub : ''))}</title><rect x="${p.x}" y="${p.y}" width="${W}" height="${H}" rx="8"/><text x="${p.x + 10}" y="${p.y + 20}" class="fx-nome">${esc(corta(e.nome, e.doc ? 21 : 25))}</text><text x="${p.x + 10}" y="${p.y + 38}" class="fx-sub">${esc(corta(i.sub || '', 30))}</text>${e.doc ? `<text x="${p.x + W - 10}" y="${p.y + 20}" class="fx-doc" text-anchor="end">doc</text>` : ''}</g>`; }).join('');
  return `<svg class="fx-svg" viewBox="0 0 ${larg} ${alt}" width="${larg}" height="${alt}" role="img" aria-label="Fluxograma das etapas"><defs><marker id="fxSeta" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="7" markerHeight="7" orient="auto-start-reverse"><path d="M0,0 L10,5 L0,10 z" class="fx-ponta"/></marker></defs>
    ${faixas.map((f, i) => `<rect class="fx-faixa ${i % 2 ? 'b' : ''}" x="0" y="${f.y}" width="${larg}" height="${f.h}"/><text class="fx-grupo" x="12" y="${f.y + 22}">${esc(f.g)}</text>`).join('')}${setas}${nos}</svg>`;
}
const fDias = v => v === null || v === undefined ? '—' : Math.round(v) + (Math.round(v) === 1 ? ' dia' : ' dias');
VIEWS.etapas = {
  titulo: a => a.length === 2 ? `Fluxo · quadra ${a[0]}, lote ${a[1]}` : 'Fluxo do lote',
  refresh(el, a) {
    if (a.length === 2) { fxQ = a[0]; fxL = a[1]; if (fxAba === 'tempos') fxAba = 'mapa'; }
    if (fxObra && !E().obras.includes(fxObra)) fxObra = '';
    const qs = (fxObra ? quadrasDaObra(fxObra) : E().quadras).filter(q => lotesDe(q).length); if (fxQ && !qs.includes(fxQ)) { fxQ = ''; fxL = ''; }
    if (fxL && !lotesDe(fxQ).includes(fxL)) fxL = '';
    const lotes = lotesFx(); const umLote = fxQ && fxL;
    const est = umLote ? estadoFx(fxQ, fxL) : null; const S_ = umLote ? null : statsFx(lotes);
    const info = e => { if (est) { const x = est[e.id]; return { cls: x.st + (x.atraso ? ' atraso' : ''), sub: x.ok ? (x.f ? 'concluída ' + dBR(x.f).slice(0, 5) + (x.dias !== null ? ' · ' + fDias(x.dias) : '') : 'concluída') : x.st === 'andamento' ? 'há ' + fDias(x.corre) : ST_FX[x.st].toLowerCase() }; }
      const s = S_.o[e.id]; const pct = s.n ? s.ok / s.n : 0; return { cls: pct >= 1 ? 'ok' : s.ok || s.and ? 'andamento' : s.lib ? 'liberada' : 'bloqueada', sub: `${s.ok}/${s.n} lotes${s.med !== null ? ' · média ' + fDias(s.med) : ''}` }; };
    const filtros = `<div class="filtros card">${variasObras() ? `<label class="tgl">Obra <select id="xO">${opts(E().obras, fxObra, 'Todas')}</select></label>` : ''}
      <label class="tgl">Quadra <select id="xQ">${opts(qs, fxQ, 'Todas')}</select></label>
      ${fxQ ? `<label class="tgl">Lote <select id="xL">${opts(lotesDe(fxQ), fxL, 'Todos os lotes')}</select></label>` : ''}
      <button class="btn rw" data-act="fxEtapas">Editar etapas</button></div>
      <div class="seg">${[['mapa', 'Fluxograma'], ['lote', 'Etapas do lote'], ['tempos', 'Tempos e comparação']].map(([v, l]) => `<button class="${v === fxAba ? 'on' : ''}" data-act="fxAba" data-v="${v}">${l}</button>`).join('')}</div>`;
    let corpo = '';
    if (fxAba === 'mapa') {
      corpo = `<div class="card"><h2>${umLote ? `Quadra ${esc(fxQ)}, lote ${esc(fxL)}` : `${lotes.length} lote(s)${fxQ ? ' da quadra ' + esc(fxQ) : ''}`}</h2>
        <p class="legenda-fx small"><span class="ok">concluída</span><span class="andamento">em andamento</span><span class="liberada">pode começar</span><span class="bloqueada">aguardando anteriores</span><span class="atraso">acima do prazo</span></p>
        <div class="fx-wrap">${svgFx(info, umLote ? 'fxEd' : 'fxNo')}</div>
        <p class="muted small">${umLote ? 'Toque numa etapa para registrar início e conclusão. Etapas ligadas a um documento do checklist (marca "doc") são concluídas sozinhas quando o documento entra no checklist ou no Drive.' : 'Escolha a quadra e o lote para registrar datas. Aqui cada etapa mostra quantos lotes já concluíram e o tempo médio.'}</p></div>`;
    } else if (fxAba === 'lote') {
      if (!umLote) corpo = `<div class="card"><p class="muted">Escolha a quadra e o lote acima.</p>${fxQ ? `<div class="chips">${lotesDe(fxQ).map(l => `<button class="chip" data-act="fxIr" data-q="${esc(fxQ)}" data-l="${esc(l)}">Lote ${esc(l)}</button>`).join('')}</div>` : ''}</div>`;
      else {
        const xs = ordemFx(etapasFx()).map(e => est[e.id]); const nOk = xs.filter(x => x.ok).length;
        const ciclo = (() => { const ini = xs.map(x => x.i || (x.ok && x.f && !(x.e.dep || []).length ? x.f : '')).filter(Boolean).sort()[0]; const fim = xs.every(x => x.ok) ? xs.map(x => x.f).filter(Boolean).sort().pop() : ''; return ini ? (fim ? fDias(difDias(ini, fim)) + ' (concluído)' : fDias(difDias(ini, hoje())) + ' até agora') : '—'; })();
        corpo = `<div class="grid kpis k4">${kpi('Etapas concluídas', `${nOk}/${xs.length}`, fPct.format(nOk / xs.length), '')}${kpi('Em andamento', xs.filter(x => x.st === 'andamento').length, '', '')}${kpi('Podem começar', xs.filter(x => x.st === 'liberada').length, '', '')}${kpi('Ciclo do lote', ciclo, 'da primeira etapa até agora', '')}</div>
          <div class="card"><div class="tbl-wrap"><table class="tbl fx-tab"><thead><tr><th class="l">Etapa</th><th class="l">Situação</th><th class="l">Início</th><th class="l">Conclusão</th><th>Duração</th><th>Prazo</th><th></th></tr></thead><tbody>
          ${GRUPOS_FX.concat([...new Set(xs.map(x => x.e.grupo))].filter(g => !GRUPOS_FX.includes(g))).map(g => { const gs = xs.filter(x => x.e.grupo === g); return gs.length ? `<tr class="dem-sub"><td class="l" colspan="7">${esc(g)}</td></tr>` + gs.map(x => `<tr class="${x.atraso ? 'atraso' : ''}"><td class="l">${esc(x.e.nome)}${x.e.doc ? ` <span class="muted small">· ${esc(x.e.doc)}</span>` : ''}</td><td class="l"><span class="chip fx-${x.st}">${ST_FX[x.st]}${x.auto ? ' (checklist)' : ''}</span></td><td class="l">${x.i ? dBR(x.i) : x.ini ? `<span class="muted">${dBR(x.ini)}</span>` : '—'}</td><td class="l">${x.f ? dBR(x.f) : x.ok ? 'sem data' : '—'}</td><td class="${x.atraso ? 'neg' : ''}">${x.ok ? fDias(x.dias) : x.corre !== null ? fDias(x.corre) + ' até agora' : '—'}</td><td>${x.e.prazo ? x.e.prazo + ' d' : '—'}</td><td><button class="btn sm rw" data-act="fxEd" data-id="${esc(x.e.id)}">Registrar</button></td></tr>`).join('') : ''; }).join('')}
          </tbody></table></div><p class="muted small">Sem data de início, a duração conta a partir da conclusão da última etapa anterior (em cinza).</p></div>`;
      }
    } else {
      const S2 = statsFx(lotes); const qsC = (fxQ ? [fxQ] : qs).slice(0, 10);
      const porQ = Object.fromEntries(qsC.map(q => [q, statsFx(lotesDe(q).map(l => [q, l])).o]));
      corpo = `<div class="card"><h2>Tempo por etapa <span class="muted small">${lotes.length} lote(s)</span></h2><div class="tbl-wrap"><table class="tbl fx-tab"><thead><tr><th class="l">Etapa</th><th>Concluídos</th><th>Média</th><th>Mediana</th><th>Mais rápido</th><th>Mais lento</th><th>Prazo</th><th>Acima do prazo</th>${qsC.length > 1 ? qsC.map(q => `<th>Média ${esc(q)}</th>`).join('') : ''}</tr></thead><tbody>
        ${ordemFx(etapasFx()).map(e => { const s = S2.o[e.id]; return `<tr><td class="l">${esc(e.nome)} <span class="muted small">${esc(e.grupo)}</span></td><td>${s.ok}/${s.n}</td><td>${fDias(s.med)}</td><td>${fDias(s.mediana)}</td><td>${fDias(s.min)}</td><td>${fDias(s.max)}</td><td>${e.prazo ? e.prazo + ' d' : '—'}</td><td class="${s.atr ? 'neg' : ''}">${s.atr || '—'}</td>${qsC.length > 1 ? qsC.map(q => `<td>${fDias(porQ[q][e.id].med)}</td>`).join('') : ''}</tr>`; }).join('')}</tbody></table></div>
        <p class="muted small">A duração de cada etapa vai da data de início (ou da conclusão da etapa anterior) até a conclusão, em dias corridos. Só entram etapas concluídas com data. Use o filtro de obra e quadra para comparar empreendimentos.</p></div>`;
    }
    el.innerHTML = filtros + corpo;
    const on = (id, fn) => { const x = $('#' + id, el); if (x) x.onchange = e => { fn(e.target.value); if (location.hash.split('/').length > 2) location.hash = '#/etapas'; else this.refresh(el, []); }; };
    on('xO', v => { fxObra = v; fxQ = ''; fxL = ''; }); on('xQ', v => { fxQ = v; fxL = ''; }); on('xL', v => { fxL = v; if (v && fxAba === 'tempos') fxAba = 'mapa'; });
  }
};
ACOES.fxAba = d => { fxAba = d.v; refreshView(); };
ACOES.fxIr = d => { fxQ = d.q; fxL = d.l; refreshView(); };
ACOES.fxNo = d => {
  const e = etapasFx().find(x => x.id === d.id); if (!e) return; const ls = lotesFx();
  const xs = ls.map(([q, l]) => [q, l, estadoFx(q, l)[e.id]]);
  folha(e.nome, `<p class="muted small">${esc(e.grupo)}${(e.dep || []).length ? ' · depois de ' + esc(e.dep.map(x => etapasFx().find(z => z.id === x)?.nome).filter(Boolean).join(', ')) : ''}${e.prazo ? ' · prazo ' + e.prazo + ' dias' : ''}</p>
    <div class="tbl-wrap"><table class="tbl"><thead><tr><th class="l">Lote</th><th class="l">Situação</th><th class="l">Conclusão</th><th>Duração</th></tr></thead><tbody>${xs.map(([q, l, x]) => `<tr class="click" data-act="fxIrLote" data-q="${esc(q)}" data-l="${esc(l)}"><td class="l">${esc(q)} · ${esc(l)}</td><td class="l"><span class="chip fx-${x.st}">${ST_FX[x.st]}</span></td><td class="l">${x.f ? dBR(x.f) : '—'}</td><td class="${x.atraso ? 'neg' : ''}">${x.ok ? fDias(x.dias) : x.corre !== null ? fDias(x.corre) + ' até agora' : '—'}</td></tr>`).join('')}</tbody></table></div>`, { largo: true });
};
ACOES.fxIrLote = d => { document.querySelector('.sheet [data-close]')?.click(); location.hash = '#/etapas/' + encodeURIComponent(d.q) + '/' + encodeURIComponent(d.l); };
function salvarRegLote(q, l, reg) { const k = q + '|' + l; S.fluxoLotes = { ...(S.fluxoLotes || {}), [k]: JSON.stringify(reg) }; gravar(doc(db, 'fluxoLotes', 'Q_' + idBloco(q)), { q, lotes: { [l]: JSON.stringify(reg) } }, { merge: true }).catch(onErroEscrita); }
ACOES.fxEd = d => {
  const q = fxQ, l = fxL; if (!q || !l) return; const e = etapasFx().find(x => x.id === d.id); if (!e) return;
  const x = estadoFx(q, l)[e.id]; const r = regLote(q, l)[e.id] || {};
  const f = folha(e.nome, `<form class="form" id="fx"><p class="muted small">Quadra ${esc(q)}, lote ${esc(l)} · ${esc(e.grupo)} · ${ST_FX[x.st]}${x.auto ? ' pelo checklist' : ''}${(e.dep || []).length ? '<br>Depois de: ' + esc(e.dep.map(z => etapasFx().find(w => w.id === z)?.nome).filter(Boolean).join(', ')) : ''}</p>
    <div class="g2"><label>Início (solicitação)<input type="date" name="i" value="${esc(r.i || '')}"></label><label>Conclusão<input type="date" name="f" value="${esc(r.f || '')}"></label></div>
    <div class="acts left"><button type="button" class="btn sm" id="hI">Começou hoje</button><button type="button" class="btn sm" id="hF">Concluiu hoje</button></div>
    <label>Observação<input name="o" value="${esc(r.o || '')}" placeholder="Ex.: protocolo 1234, aguardando cartório"></label>
    <label class="chk"><input type="checkbox" name="todos"> Aplicar a todos os lotes da quadra ${esc(q)}</label>
    <label class="chk"><input type="checkbox" name="soVazios" checked> Só nos lotes que ainda não têm data nesta etapa</label>
    <div class="acts"><button type="button" class="btn danger ghost" id="lim">Limpar datas</button><button type="button" class="btn" data-close>Cancelar</button><button class="btn primary">Salvar</button></div></form>`);
  const F = f.el.querySelector('#fx');
  F.querySelector('#hI').onclick = () => { F.i.value = hoje(); }; F.querySelector('#hF').onclick = () => { F.f.value = hoje(); if (!F.i.value && !x.ini) F.i.value = hoje(); };
  const aplica = val => { const alvos = F.todos.checked ? lotesDe(q) : [l]; let n = 0; for (const lt of alvos) { const reg = regLote(q, lt); if (F.todos.checked && F.soVazios.checked && lt !== l && (reg[e.id]?.i || reg[e.id]?.f)) continue; if (val) reg[e.id] = val; else delete reg[e.id]; salvarRegLote(q, lt, reg); n++; } return n; };
  F.querySelector('#lim').onclick = () => { const n = aplica(null); f.fechar(); refreshView(); toast(n > 1 ? `Datas limpas em ${n} lotes` : 'Datas limpas'); };
  F.onsubmit = ev => { ev.preventDefault(); if (F.i.value && F.f.value && F.f.value < F.i.value) { toast('A conclusão é anterior ao início.', 'err'); return; } const v = {}; if (F.i.value) v.i = F.i.value; if (F.f.value) v.f = F.f.value; if (F.o.value.trim()) v.o = F.o.value.trim(); const n = aplica(Object.keys(v).length ? v : null); f.fechar(); refreshView(); toast(n > 1 ? `Etapa salva em ${n} lotes` : 'Etapa salva'); };
};
function salvarEtapas(es) { S.fluxo = { ...(S.fluxo || {}), etapas: JSON.stringify(es) }; gravar(doc(db, 'config', 'fluxo'), { etapas: S.fluxo.etapas }, { merge: true }).catch(onErroEscrita); refreshView(); }
function temCiclo(es) { const by = Object.fromEntries(es.map(e => [e.id, e])); const vis = {}; const dfs = id => { if (vis[id] === 1) return true; if (vis[id] === 2) return false; vis[id] = 1; for (const d of (by[id]?.dep || [])) if (by[d] && dfs(d)) return true; vis[id] = 2; return false; }; return es.some(e => dfs(e.id)); }
ACOES.fxEtapas = () => {
  const es = etapasFx();
  const f = folha('Etapas do fluxo', `<p class="muted small">Defina as etapas, de quais outras cada uma depende, o prazo esperado e, se quiser, o documento do checklist que conclui a etapa sozinho.</p>
    ${GRUPOS_FX.concat([...new Set(es.map(e => e.grupo))].filter(g => !GRUPOS_FX.includes(g))).map(g => { const gs = ordemFx(es).filter(e => e.grupo === g); return gs.length ? `<div class="lbl">${esc(g)}</div><div class="chips" style="margin-bottom:10px">${gs.map(e => `<button class="chip" data-act="fxEtapa" data-id="${esc(e.id)}">${esc(e.nome)}${e.prazo ? ' · ' + e.prazo + ' d' : ''}</button>`).join('')}</div>` : ''; }).join('')}
    <div class="acts"><button class="btn ghost" data-act="fxPadrao">Voltar ao modelo padrão</button><button class="btn primary" data-act="fxEtapa">Nova etapa</button></div>`, { largo: true });
  f.el.addEventListener('click', ev => { if (ev.target.closest('[data-act="fxEtapa"],[data-act="fxPadrao"]')) f.fechar(); });
};
ACOES.fxPadrao = async () => { if (await confirmar('Voltar as etapas ao modelo padrão (baseado na aba Fluxograma da planilha)? As datas já registradas nos lotes continuam guardadas.', 'Voltar ao padrão', false)) salvarEtapas(FLUXO_PADRAO); };
ACOES.fxEtapa = d => {
  const es = etapasFx(); const e = d.id ? es.find(x => x.id === d.id) : null;
  const grupos = [...new Set([...GRUPOS_FX, ...es.map(x => x.grupo)])];
  const f = folha(e ? e.nome : 'Nova etapa', `<form class="form" id="fx"><label>Nome<input name="nome" required value="${esc(e?.nome || '')}"></label>
    <div class="g2"><label>Grupo<input name="grupo" list="dlGr" required value="${esc(e?.grupo || GRUPOS_FX[0])}"><datalist id="dlGr">${grupos.map(g => `<option value="${esc(g)}">`).join('')}</datalist></label><label>Prazo esperado (dias)<input name="prazo" type="number" min="0" inputmode="numeric" value="${e?.prazo || ''}" placeholder="opcional"></label></div>
    <label>Concluída automaticamente pelo documento do checklist<select name="doc">${opts(docsCk(), e?.doc || '', 'Nenhum')}</select></label>
    <div class="lbl">Depende de</div><div class="chips fx-deps">${ordemFx(es).filter(x => !e || x.id !== e.id).map(x => `<label class="chip"><input type="checkbox" name="dep" value="${esc(x.id)}" ${e && (e.dep || []).includes(x.id) ? 'checked' : ''}> ${esc(x.nome)}</label>`).join('')}</div>
    <div class="acts">${e ? '<button type="button" class="btn danger ghost" id="rm">Excluir etapa</button>' : ''}<button type="button" class="btn" data-close>Cancelar</button><button class="btn primary">Salvar</button></div></form>`, { largo: true });
  const F = f.el.querySelector('#fx');
  F.onsubmit = ev => {
    ev.preventDefault(); const n = { id: e ? e.id : 'e' + uid().slice(-6), nome: F.nome.value.trim(), grupo: F.grupo.value.trim() || 'Outros', prazo: +F.prazo.value || 0, doc: F.doc.value, dep: $$('[name=dep]:checked', F).map(c => c.value) };
    const nova = e ? es.map(x => x.id === e.id ? n : x) : [...es, n];
    if (temCiclo(nova)) { toast('Essa dependência cria um ciclo (uma etapa dependeria dela mesma).', 'err'); return; }
    salvarEtapas(nova); f.fechar(); toast('Etapa salva');
  };
  const rm = F.querySelector('#rm'); if (rm) rm.onclick = async () => { if (await confirmar(`Excluir a etapa <b>${esc(e.nome)}</b>? As etapas que dependiam dela passam a depender das anteriores a ela.`, 'Excluir')) { const nova = es.filter(x => x.id !== e.id).map(x => (x.dep || []).includes(e.id) ? { ...x, dep: [...new Set([...x.dep.filter(z => z !== e.id), ...(e.dep || [])])] } : x); salvarEtapas(nova); f.fechar(); } };
};

/* ---------- Contas recorrentes ---------- */
const FREQ = [['1', 'Mensal'], ['2', 'Bimestral'], ['3', 'Trimestral'], ['6', 'Semestral'], ['12', 'Anual']];
const HORIZ_REC = 6;
const recorrencias = () => Object.values((S.recorrencias && S.recorrencias.itens) || {}).map(v => { try { return JSON.parse(v); } catch (e) { return null; } }).filter(Boolean);
function salvarRec(r) {
  S.recorrencias = { itens: { ...((S.recorrencias && S.recorrencias.itens) || {}), [r.id]: JSON.stringify(r) } };
  gravar(doc(db, 'config', 'recorrencias'), { itens: { [r.id]: JSON.stringify(r) } }, { merge: true }).catch(onErroEscrita);
}
const tagMes = d => ` (${MES3[+d.slice(5, 7) - 1]}/${d.slice(2, 4)})`;
const semTag = s => s.replace(/\s\((jan|fev|mar|abr|mai|jun|jul|ago|set|out|nov|dez)\/\d{2}\)$/, '');
let _gerandoRec = false;
function gerarRecorrencias() {
  if (_gerandoRec || somenteLeitura() || !S.user || !S.loaded.titulos || S.recorrencias === undefined) return;
  const lim = addMonths(hoje(), HORIZ_REC); const novos = [], upd = [];
  for (const r of recorrencias()) {
    if (!r.ativo || !r.inicio) continue;
    let ult = r.ultimo || '', mudou = false;
    for (let k = 0; k < 600; k++) {
      const base = addMonths(r.inicio, k * (r.passo || 1));
      if (base > lim || (r.fim && base > r.fim)) break;
      if (ult && base <= ult) continue;
      const id = 'rc' + r.id + base.replace(/-/g, '');
      if (!S.byId.has(id)) {
        const venc = r.util ? proximoDiaUtil(base) : base;
        const t = { id, item: r.item + tagMes(base), valor: r.valor, juros: 0, nf: 0, nfNum: '', local: r.local || E().obras[0], quadra: r.quadra || '-', lote: '-', destino: r.destino, venc, status: 'aberto', forn: r.forn || '', obs: r.obs || '', grupo: 'rec' + r.id, rec: r.id, criado: Date.now() };
        if (venc !== base) t.vencOrig = base; if (r.emp) t.emp = r.emp; novos.push(t);
      }
      ult = base; mudou = true;
    }
    if (mudou) upd.push({ ...r, ultimo: ult });
  }
  if (!novos.length && !upd.length) return;
  _gerandoRec = true;
  if (novos.length) salvarTitulos(novos, { acao: 'recorrente' });
  upd.forEach(salvarRec);
  setTimeout(() => { _gerandoRec = false; }, 3000);
}
function proxRec(r) { const ab = S.list.filter(t => t.rec === r.id && !pago(t)).sort((a, b) => a.venc.localeCompare(b.venc)); return ab[0]; }
VIEWS.recorrentes = {
  titulo: 'Contas recorrentes',
  refresh(el) {
    const rs = recorrencias().sort((a, b) => (b.ativo - a.ativo) || a.item.localeCompare(b.item));
    const mensal = r => r.valor / (r.passo || 1);
    const at = rs.filter(r => r.ativo);
    el.innerHTML = `<div class="grid kpis k4">${kpi('Ativas', at.length, rs.length - at.length ? (rs.length - at.length) + ' encerrada(s) ou pausada(s)' : '', '')}${kpi('Saídas fixas por mês', brl(sum(at.filter(r => r.valor < 0), mensal)), 'média mensal', 'neg')}${kpi('Entradas fixas por mês', brl(sum(at.filter(r => r.valor > 0), mensal)), 'média mensal', 'pos')}${kpi('Lançadas em aberto', S.list.filter(t => t.rec && !pago(t)).length, 'próximos ' + HORIZ_REC + ' meses', '')}</div>
      <div class="card"><h2>Contas recorrentes <button class="btn sm primary rw" data-act="recNova">Nova conta recorrente</button></h2>
      <div class="tbl-wrap"><table class="tbl"><thead><tr><th class="l">Descrição</th><th>Valor</th><th class="l">Repete</th><th class="l">Próximo vencimento</th><th class="l">Termina</th>${variasEmpresas() ? '<th class="l">CNPJ</th>' : ''}<th class="l">Situação</th></tr></thead><tbody>
      ${rs.map(r => { const p = proxRec(r); return `<tr class="click" data-act="recEd" data-id="${esc(r.id)}"><td class="l"><b>${esc(r.item)}</b><br><span class="muted small">${esc([r.destino, r.forn, r.quadra && r.quadra !== '-' ? 'Quadra ' + r.quadra : ''].filter(Boolean).join(' · '))}</span></td><td class="${cls(r.valor)}">${brl2(r.valor)}</td><td class="l">${FREQ.find(x => +x[0] === (r.passo || 1))?.[1] || ''}, dia ${+r.inicio.slice(8, 10)}</td><td class="l">${p ? dBR(p.venc) : '—'}</td><td class="l">${r.fim ? dBR(r.fim) : 'sem data final'}</td>${variasEmpresas() ? `<td class="l">${esc(empPorId(r.emp)?.nome || '—')}</td>` : ''}<td class="l"><span class="chip ${r.ativo ? 'ok' : ''}">${r.ativo ? 'Ativa' : 'Pausada'}</span></td></tr>`; }).join('') || `<tr><td class="l muted" colspan="7">Nenhuma conta recorrente. Crie aqui ou marque "Conta recorrente" ao lançar.</td></tr>`}</tbody></table></div>
      <p class="muted small">O app mantém os próximos ${HORIZ_REC} meses de cada conta lançados em Contas a pagar ou a receber e cria os meses seguintes sozinho, sempre que alguém com permissão de edição abre o app. Excluir um mês não faz ele voltar.</p></div>`;
  }
};
ACOES.recNova = () => formNovo({ recorrente: true });
ACOES.recEd = d => {
  const r = recorrencias().find(x => x.id === d.id); if (!r) return;
  const ab = S.list.filter(t => t.rec === r.id && !pago(t));
  const f = folha('Conta recorrente', `<form class="form" id="fx">
    <label>Descrição<input name="item" required value="${esc(r.item)}"></label>
    <div class="g2"><label>Valor (R$)<input name="valor" inputmode="decimal" required value="${fNum2.format(Math.abs(r.valor))}"></label><label>Destino<select name="destino">${optsDestino(r.destino)}</select></label></div>
    <div class="g2"><label>Fornecedor / cliente<input name="forn" value="${esc(r.forn || '')}"></label><label>Termina em <span class="muted">(opcional)</span><input type="date" name="fim" value="${esc(r.fim || '')}"></label></div>
    ${variasEmpresas() ? `<label>CNPJ da nota / pagamento<select name="emp">${opts(empOpts(), r.emp || '', 'Escolha o CNPJ…')}</select></label>` : ''}
    <label class="chk"><input type="checkbox" name="ativo" ${r.ativo ? 'checked' : ''}> Ativa (desmarque para pausar)</label>
    ${ab.length ? `<label class="chk"><input type="checkbox" name="aplicar" checked> Aplicar também aos ${ab.length} mês(es) já lançados em aberto</label>` : ''}
    <p class="muted small">Repete ${(FREQ.find(x => +x[0] === (r.passo || 1))?.[1] || '').toLowerCase()} desde ${dBR(r.inicio)}. Criada por ${quem(r.por)}.</p>
    <div class="acts"><button type="button" class="btn danger ghost" id="rDel">Excluir</button><button type="button" class="btn" data-close>Cancelar</button><button class="btn primary">Salvar</button></div></form>`);
  const F = f.el.querySelector('#fx');
  F.onsubmit = e => {
    e.preventDefault(); const v = parseBR(F.valor.value); if (!(v > 0)) { toast('Valor inválido', 'err'); return; }
    const n = { ...r, item: F.item.value.trim(), valor: Math.sign(r.valor || -1) * r2(v), destino: F.destino.value, forn: F.forn.value.trim(), fim: F.fim.value || '', ativo: F.ativo.checked };
    if (F.emp) n.emp = F.emp.value;
    salvarRec(n);
    if (F.aplicar && F.aplicar.checked) {
      const mud = ab.filter(t => !n.fim || (t.vencOrig || t.venc) <= n.fim).map(t => { const x = { ...t, item: n.item + (t.item.match(/\s\(\w{3}\/\d{2}\)$/) || [''])[0], valor: n.valor, destino: n.destino, forn: n.forn }; if (n.emp) x.emp = n.emp; return x; });
      if (mud.length) salvarTitulos(mud);
      const fora = n.fim ? ab.filter(t => (t.vencOrig || t.venc) > n.fim) : []; if (fora.length) excluirTitulos(fora.map(t => t.id));
    }
    f.fechar(); toast('Conta recorrente salva');
  };
  F.querySelector('#rDel').onclick = async () => {
    if (!(await confirmar(`Excluir a conta recorrente <b>${esc(r.item)}</b>?${ab.length ? ` Os ${ab.length} mês(es) já lançados em aberto também serão excluídos. Os já pagos ficam.` : ''}`, 'Excluir'))) return;
    gravar(doc(db, 'config', 'recorrencias'), { itens: { [r.id]: deleteField() } }, { merge: true }).catch(onErroEscrita);
    if (S.recorrencias?.itens) delete S.recorrencias.itens[r.id];
    if (ab.length) excluirTitulos(ab.map(t => t.id));
    f.fechar(); toast('Conta recorrente excluída');
  };
};

/* ---------- Demonstrativos: DRE, DFC e Balanço (gerenciais) ---------- */
let demEmp = '', demAno = '', demMes = '', demVis = 'mensal', demReg = 'caixa', demAba = 'dre';
const VISOES = [['mensal', 'Mensal'], ['trimestral', 'Trimestral'], ['semestral', 'Semestral'], ['anual', 'Anual (comparar anos)'], ['mes', 'Um mês']];
const FIN_REC = ['Rendimento'];
const valDem = t => (t.valor || 0);
function dreDe(L, ini, fim, regime) {
  const noPer = t => regime === 'caixa' ? pago(t) && (t.pag || t.venc) >= ini && (t.pag || t.venc) <= fim : t.venc >= ini && t.venc <= fim;
  const ts = L.filter(t => noPer(t) && secaoDe(t.destino) !== 'aporte');
  const by = f => { const o = {}; for (const t of ts.filter(f)) o[t.destino] = (o[t.destino] || 0) + valDem(t); return o; };
  const rec = by(t => secaoDe(t.destino) === 'rec' && !FIN_REC.includes(t.destino));
  const imp = sum(ts.filter(t => t.destino === 'Impostos'), valDem);
  const custo = by(t => secaoDe(t.destino) === 'op');
  const desp = by(t => secaoDe(t.destino) === 'nonop' && t.destino !== 'Impostos');
  const fin = sum(ts.filter(t => FIN_REC.includes(t.destino)), valDem);
  const juros = sum(L.filter(t => pago(t) && (regime === 'caixa' ? (t.pag || t.venc) >= ini && (t.pag || t.venc) <= fim : t.venc >= ini && t.venc <= fim)), t => t.juros || 0);
  const S_ = o => sum(Object.values(o), v => v);
  const rb = S_(rec), rl = rb + imp, cu = S_(custo), lb = rl + cu, de = S_(desp), ro = lb + de, res = ro + fin + juros;
  return { rec, imp, custo, desp, fin, juros, rb, rl, cu, lb, de, ro, res };
}
function dfcDe(L, contas, ini, fim) {
  const base = sum(contas, c => +c.inicial || 0);
  const real = L.filter(pago); const dt = t => t.pag || t.venc;
  const antes = real.filter(t => dt(t) < ini), per = real.filter(t => dt(t) >= ini && dt(t) <= fim);
  const v = t => (t.valor || 0) + (t.juros || 0);
  const g = f => sum(per.filter(f), v);
  const sec = t => secaoDe(t.destino);
  const o = {
    ini: base + sum(antes, v),
    receb: g(t => sec(t) === 'rec' && !FIN_REC.includes(t.destino)),
    obra: g(t => sec(t) === 'op'),
    desp: g(t => sec(t) === 'nonop' && t.destino !== 'Impostos'),
    imp: g(t => t.destino === 'Impostos'),
    fin: g(t => FIN_REC.includes(t.destino)),
    apIn: sum(per.filter(t => sec(t) === 'aporte' && v(t) > 0), v), apOut: sum(per.filter(t => sec(t) === 'aporte' && v(t) < 0), v)
  };
  o.oper = o.receb + o.obra + o.desp + o.imp + o.fin; o.finc = o.apIn + o.apOut; o.vari = o.oper + o.finc; o.fim = o.ini + o.vari;
  return o;
}
function balancoDe(L, emp, D) {
  const cs = contasEmp(emp), nomes = new Set(cs.map(c => c.nome)), ini = sum(cs, c => +c.inicial || 0);
  const v = t => (t.valor || 0) + (t.juros || 0); const dt = t => t.pag || t.venc;
  const realL = L.filter(t => pago(t) && dt(t) <= D);
  const caixaTeo = ini + sum(realL, v);
  const dispon = emp ? ini + sum(S.list.filter(t => pago(t) && dt(t) <= D && nomes.has(t.conta || '')), v) : caixaTeo;
  const existia = t => t.venc <= D || (t.criado && isoLocal(new Date(t.criado)) <= D);
  const ab = L.filter(t => (!pago(t) || dt(t) > D) && existia(t));
  const lim = addDays(D, 365); const ap0 = t => secaoDe(t.destino) === 'aporte'; const emprest = t => ap0(t) && (['emprestimo', 'devolucao', 'terceiro'].includes(t.aporteTipo) || (t.aporteTipo === 'conversivel' && !(t.conversao && t.conversao <= D))); const ap = t => ap0(t) && !emprest(t);
  const o = {
    dispon, inter: caixaTeo - dispon,
    recCP: sum(ab.filter(t => t.valor > 0 && t.venc <= lim), t => t.valor), recLP: sum(ab.filter(t => t.valor > 0 && t.venc > lim), t => t.valor),
    pagCP: -sum(ab.filter(t => t.valor < 0 && t.venc <= lim), t => t.valor), pagLP: -sum(ab.filter(t => t.valor < 0 && t.venc > lim), t => t.valor),
    capital: ini, aportes: sum(realL.filter(ap), t => t.valor) + sum(ab.filter(ap), t => t.valor), emprest: sum(realL.filter(emprest), t => t.valor) + sum(ab.filter(emprest), t => t.valor),
    resReal: sum(realL.filter(t => !ap0(t)), v) + sum(realL.filter(ap0), t => t.juros || 0), resAReal: sum(ab.filter(t => !ap0(t)), t => t.valor)
  };
  o.ativo = o.dispon + Math.max(0, o.inter) + o.recCP + o.recLP; o.passivo = o.pagCP + o.pagLP + Math.max(0, -o.inter) + o.emprest; o.pl = o.capital + o.aportes + o.resReal + o.resAReal;
  return o;
}
function periodosDem() {
  const bloco = (n, nome) => [...Array(12 / n)].map((_, i) => { const ini = `${demAno}-${String(i * n + 1).padStart(2, '0')}-01`; return { ini, fim: addDays(addMonths(ini, n), -1), nome: nome(i) }; });
  const ano = { ini: demAno + '-01-01', fim: demAno + '-12-31', nome: 'Ano', tot: true };
  if (demVis === 'trimestral') return [...bloco(3, i => (i + 1) + 'º tri'), ano];
  if (demVis === 'semestral') return [...bloco(6, i => (i + 1) + 'º sem'), ano];
  if (demVis === 'anual') return [...Array(4)].map((_, i) => String(+demAno - 3 + i)).map(a => ({ ini: a + '-01-01', fim: a + '-12-31', nome: a }));
  if (demVis === 'mes') { if (!demMes) demMes = hoje().slice(5, 7); const ini = `${demAno}-${demMes}-01`; return [{ ini, fim: addDays(addMonths(ini, 1), -1), nome: MES3[+demMes - 1] + '/' + demAno.slice(2) }]; }
  const ms = [...Array(12)].map((_, i) => { const ini = `${demAno}-${String(i + 1).padStart(2, '0')}-01`; return { ini, fim: addDays(addMonths(ini, 1), -1), nome: MES3[i] }; });
  return [...ms, { ini: demAno + '-01-01', fim: demAno + '-12-31', nome: 'Ano', tot: true }];
}
/* Monta as linhas de cada demonstrativo: [rótulo, [valores], estilo] */
function linhasDem() {
  const L = listaEmp(demEmp), ps = periodosDem();
  if (demAba === 'dre') {
    const ds = ps.map(p => dreDe(L, p.ini, p.fim, demReg)); const A = ds[ds.length - 1];
    const det = (k, sinal) => Object.keys(A[k]).sort((a, b) => Math.abs(A[k][b]) - Math.abs(A[k][a])).map(d => ['  ' + d, ds.map(x => x[k][d] || 0), 'item']);
    return { cols: ps.map(p => p.nome), rows: [
      ['Receita bruta', ds.map(x => x.rb), 'sub'], ...det('rec'),
      ['(−) Impostos sobre a receita', ds.map(x => x.imp), 'item'],
      ['Receita líquida', ds.map(x => x.rl), 'total'],
      ['(−) Custo das obras', ds.map(x => x.cu), 'sub'], ...det('custo'),
      ['Lucro bruto', ds.map(x => x.lb), 'total'],
      ['Margem bruta', ds.map(x => x.rl ? x.lb / x.rl : null), 'pct'],
      ['(−) Despesas', ds.map(x => x.de), 'sub'], ...det('desp'),
      ['Resultado operacional', ds.map(x => x.ro), 'total'],
      ['(+) Receitas financeiras', ds.map(x => x.fin), 'item'], ['(−) Juros e multas', ds.map(x => x.juros), 'item'],
      ['Resultado do período', ds.map(x => x.res), 'grand'],
      ['Margem líquida', ds.map(x => x.rl ? x.res / x.rl : null), 'pct']
    ] };
  }
  if (demAba === 'dfc') {
    const cs = contasEmp(demEmp); const ds = ps.map(p => dfcDe(L, cs, p.ini, p.fim));
    return { cols: ps.map(p => p.nome), rows: [
      ['Saldo inicial de caixa', ds.map(x => x.ini), 'sub'],
      ['Atividades operacionais', ds.map(x => x.oper), 'total'],
      ['  Recebimentos de vendas e receitas', ds.map(x => x.receb), 'item'], ['  Pagamentos de obra', ds.map(x => x.obra), 'item'],
      ['  Despesas pagas', ds.map(x => x.desp), 'item'], ['  Impostos pagos', ds.map(x => x.imp), 'item'], ['  Rendimentos recebidos', ds.map(x => x.fin), 'item'],
      ['Financiamento (aportes e retiradas)', ds.map(x => x.finc), 'total'],
      ['  Aportes recebidos', ds.map(x => x.apIn), 'item'], ['  Aportes enviados / retiradas', ds.map(x => x.apOut), 'item'],
      ['Variação do caixa', ds.map(x => x.vari), 'total'],
      ['Saldo final de caixa', ds.map(x => x.fim), 'grand']
    ] };
  }
  const pp = ps.filter(p => !p.tot);
  const datas = pp.length === 1 ? [addDays(pp[0].ini, -1), pp[0].fim] : pp.map(p => p.fim);
  const bs = datas.map(d => balancoDe(L, demEmp, d));
  const linha = (n, f, s) => [n, bs.map(f), s];
  return { cols: datas.map(d => pp.length === 1 ? dBR(d) : demVis === 'anual' ? d.slice(0, 4) : dBR(d).slice(0, 5)), rows: [
    linha('ATIVO', b => b.ativo, 'grand'),
    linha('  Caixa e bancos', b => b.dispon, 'item'),
    linha('  Contas a receber (até 12 meses)', b => b.recCP, 'item'),
    linha('  Contas a receber (após 12 meses)', b => b.recLP, 'item'),
    ...(demEmp ? [linha('  Créditos com empresas do grupo', b => Math.max(0, b.inter), 'item')] : []),
    linha('PASSIVO', b => b.passivo, 'grand'),
    linha('  Contas a pagar (até 12 meses)', b => b.pagCP, 'item'),
    linha('  Contas a pagar (após 12 meses)', b => b.pagLP, 'item'),
    ...(demEmp ? [linha('  Débitos com empresas do grupo', b => Math.max(0, -b.inter), 'item')] : []),
    linha('  Empréstimos e dívidas conversíveis', b => b.emprest, 'item'),
    linha('PATRIMÔNIO LÍQUIDO', b => b.pl, 'grand'),
    linha('  Capital inicial (saldo inicial das contas)', b => b.capital, 'item'),
    linha('  Aportes de capital e retiradas', b => b.aportes, 'item'),
    linha('  Resultado acumulado realizado', b => b.resReal, 'item'),
    linha('  Resultado a realizar (a receber − a pagar)', b => b.resAReal, 'item'),
    linha('Passivo + patrimônio líquido', b => b.passivo + b.pl, 'total')
  ] };
}
const fmtDem = (v, s) => v === null || v === undefined ? '—' : s === 'pct' ? fPct.format(v) : Math.abs(v) < 0.5 ? '–' : brl(v);
VIEWS.demonstrativos = {
  titulo: 'DRE, DFC e BP',
  refresh(el) {
    const h = hoje(); if (!demAno) demAno = h.slice(0, 4); if (demEmp && !empPorId(demEmp)) demEmp = '';
    const anos = [...new Set([h.slice(0, 4), ...S.list.map(t => (dataRef(t) || '').slice(0, 4)).filter(a => /^\d{4}$/.test(a))])].sort().reverse();
    const { cols, rows } = linhasDem();
    el.innerHTML = `<div class="filtros card">
      ${variasEmpresas() ? `<label class="tgl">Empresa <select id="dE">${opts(empOpts(), demEmp, 'Grupo consolidado')}</select></label>` : ''}
      <label class="tgl">Ano <select id="dA">${opts(anos, demAno)}</select></label>
      <label class="tgl">Visão <select id="dV">${opts(VISOES, demVis)}</select></label>
      ${demVis === 'mes' ? `<label class="tgl">Mês <select id="dM">${opts(MESES.map((m, i) => [String(i + 1).padStart(2, '0'), m]), demMes || hoje().slice(5, 7))}</select></label>` : ''}
      ${demAba === 'dre' ? `<label class="tgl">Regime <select id="dR">${opts([['caixa', 'Caixa (pelo pagamento)'], ['comp', 'Competência (pelo vencimento)']], demReg)}</select></label>` : ''}
      <button class="btn primary" id="dPdf">Baixar PDF</button></div>
      <div class="seg">${[['dre', 'DRE'], ['dfc', 'Fluxo de caixa (DFC)'], ['bal', 'BP']].map(([v, l]) => `<button class="${v === demAba ? 'on' : ''}" data-act="demAba" data-v="${v}">${l}</button>`).join('')}</div>
      <div class="card"><h2>${demAba === 'dre' ? 'Demonstração do resultado' : demAba === 'dfc' ? 'Demonstração do fluxo de caixa (método direto)' : 'Balanço patrimonial (BP)'} <span class="muted small">${esc(nomeEmpPDF(demEmp))}</span></h2>
      <div class="tbl-wrap"><table class="tbl dem"><thead><tr><th class="l"></th>${cols.map(c => `<th>${esc(c)}</th>`).join('')}</tr></thead><tbody>
      ${rows.map(([n, vs, s]) => `<tr class="dem-${s}"><td class="l">${esc(n.trim())}</td>${vs.map((v, i) => `<td class="${s === 'pct' || Math.abs(v || 0) < 0.5 ? 'muted' : cls(v || 0)}${i === vs.length - 1 && vs.length > 2 ? ' dem-tot' : ''}">${fmtDem(v, s)}</td>`).join('')}</tr>`).join('')}</tbody></table></div>
      <p class="muted small">${notaDem()}</p></div>`;
    const on = (id, fn) => { const x = $('#' + id, el); if (x) x.onchange = e => { fn(e.target.value); this.refresh(el); }; };
    on('dE', v => demEmp = v); on('dA', v => demAno = v); on('dM', v => demMes = v); on('dV', v => demVis = v); on('dR', v => demReg = v);
    $('#dPdf', el).onclick = () => pdfDemonstrativos().catch(e => { console.error(e); toast('Erro ao gerar PDF: ' + e.message, 'err'); });
  }
};
ACOES.demAba = d => { demAba = d.v; refreshView(); };
function notaDem() {
  const base = 'Demonstrativos gerenciais montados a partir dos lançamentos do app; não substituem os da contabilidade.';
  if (demAba === 'dre') return base + (demReg === 'caixa' ? ' Regime de caixa: entra o que foi pago ou recebido no período.' : ' Regime de competência: entra o que vence no período, pago ou não.') + ' Custos de obra são lançados como custo quando ocorrem (não há estoque de imóveis). Aportes ficam fora.';
  if (demAba === 'dfc') return base + (demEmp ? ' O caixa da empresa considera todos os lançamentos do CNPJ, mesmo os pagos por conta de outra empresa do grupo.' : '');
  return base + ' Contas a pagar e a receber incluem as parcelas em aberto na data que já tinham sido lançadas ou vencidas até ela.' + (demEmp ? ' Créditos e débitos com empresas do grupo aparecem quando um CNPJ pagou ou recebeu por conta bancária de outro.' : '');
}
async function pdfDemonstrativos() {
  toast('Gerando PDF…'); await carregarLib('pdf');
  const abaAtual = demAba; const ps = periodosDem(); const largo = ps.length > 2;
  const pdf = new jsPDF({ unit: 'pt', format: 'a4', orientation: largo ? 'landscape' : 'portrait' }); const W = pdf.internal.pageSize.getWidth(), H = pdf.internal.pageSize.getHeight();
  const NAVY = [11, 32, 70]; const P = s => String(s).replace(/−/g, '-').replace(/[  ]/g, ' ');
  try { const lg = await imgParaPNG(LOGO_COMPLETO, 480); pdf.addImage(lg.data, 'PNG', 36, 16, 96, 96 * lg.r); } catch (e) { /* sem logo */ }
  pdf.setDrawColor(...NAVY); pdf.setLineWidth(1.2); pdf.line(36, 92, W - 36, 92);
  pdf.setTextColor(...NAVY); pdf.setFont('helvetica', 'bold'); pdf.setFontSize(16); pdf.text('Demonstrativos financeiros', W - 36, 38, { align: 'right' });
  pdf.setFont('helvetica', 'normal'); pdf.setFontSize(10.5); pdf.text(P(nomeEmpPDF(demEmp)), W - 36, 56, { align: 'right' });
  pdf.setFontSize(8.5); pdf.setTextColor(91, 107, 130); pdf.text(P(`${demVis === 'mes' ? MESES[+demMes - 1] + ' de ' + demAno : demVis === 'anual' ? `${+demAno - 3} a ${demAno}` : `${demAno}, visão ${VISOES.find(x => x[0] === demVis)[1].toLowerCase()}`} · DRE em regime de ${demReg === 'caixa' ? 'caixa' : 'competência'} · gerado em ${new Date().toLocaleString('pt-BR', { dateStyle: 'short', timeStyle: 'short' })}`), W - 36, 72, { align: 'right' });
  let y = 112;
  for (const [aba, tit] of [['dre', 'Demonstração do resultado (DRE)'], ['dfc', 'Demonstração do fluxo de caixa (DFC)'], ['bal', 'Balanço patrimonial (BP)']]) {
    demAba = aba; const { cols, rows } = linhasDem();
    if (y > H - 160) { pdf.addPage(); y = 50; }
    pdf.setFont('helvetica', 'bold'); pdf.setFontSize(12); pdf.setTextColor(...NAVY); pdf.text(tit, 36, y); y += 8;
    autoTable(pdf, { startY: y, head: [['', ...cols]], body: rows.map(([n, vs, s]) => [P(n), ...vs.map(v => P(fmtDem(v, s)))]), margin: { left: 36, right: 36 },
      styles: { fontSize: largo ? 7 : 8.5, cellPadding: largo ? 2.5 : 4, textColor: [15, 27, 45] }, headStyles: { fillColor: NAVY, textColor: 255, fontStyle: 'bold' },
      columnStyles: { 0: { cellWidth: largo ? 150 : 250 } },
      didParseCell: dd => { if (dd.column.index > 0) dd.cell.styles.halign = 'right'; if (dd.section === 'body') { const s = rows[dd.row.index][2]; if (s === 'total' || s === 'grand') { dd.cell.styles.fontStyle = 'bold'; dd.cell.styles.fillColor = s === 'grand' ? [226, 231, 240] : [243, 245, 249]; } if (s === 'pct') dd.cell.styles.textColor = [91, 107, 130]; } } });
    y = pdf.lastAutoTable.finalY + 10; pdf.setFont('helvetica', 'normal'); pdf.setFontSize(7.5); pdf.setTextColor(91, 107, 130);
    const nt = pdf.splitTextToSize(P(notaDem()), W - 72); pdf.text(nt, 36, y + 6); y += nt.length * 9 + 22;
  }
  demAba = abaAtual;
  const np = pdf.getNumberOfPages(); for (let i = 1; i <= np; i++) { pdf.setPage(i); pdf.setFontSize(8); pdf.setTextColor(140); pdf.text(`${i} / ${np}`, W - 36, H - 20, { align: 'right' }); }
  await baixarArquivo(pdf.output('blob'), `SFG Demonstrativos ${demAno}${demVis === 'mes' ? '-' + demMes : ' ' + demVis}${demEmp ? ' ' + (empPorId(demEmp)?.nome || '') : ''}.pdf`);
}

/* ---------- Histórico ---------- */
const ACAO_NOME = { criou: 'criou', alterou: 'alterou', baixa: 'deu baixa', estorno: 'estornou a baixa', restaurou: 'restaurou da lixeira', excluiu: 'excluiu', recorrente: 'gerou (conta recorrente)' };
const quem = e => esc((e || '').split('@')[0] || 'alguém');
const quando = ms => new Date(ms).toLocaleString('pt-BR', { day: '2-digit', month: '2-digit', year: '2-digit', hour: '2-digit', minute: '2-digit' });
function valHist(k, v) {
  if (v === '' || v === null || v === undefined) return '—';
  if (k === 'valor' || k === 'juros') return brl2(+v || 0);
  if (k === 'venc' || k === 'pag') return dBR(v);
  if (k === 'emp') return empPorId(v)?.nome || 'outro CNPJ';
  if (k === 'status') return v === 'pago' ? 'pago' : 'em aberto';
  return String(v);
}
function histTxt(h) {
  const m = h.m ? Object.entries(h.m).map(([k, [a, b]]) => k === 'anexos' ? `anexos ${a} → ${b}` : `${(CAMPOS_HIST.find(x => x[0] === k) || [k, k])[1].toLowerCase()}: ${esc(valHist(k, a))} → ${esc(valHist(k, b))}`).join('; ') : '';
  return `<span class="muted small">${quando(h.em)}</span> <b>${quem(h.por)}</b> ${ACAO_NOME[h.a] || h.a}${m ? ' · ' + m : ''}`;
}
VIEWS.historico = {
  titulo: 'Histórico e lixeira',
  refresh(el) {
    const ev = []; const vistos = new Set();
    if (!S.histCompleto && !this._carregando) { this._carregando = true; carregarHistCompleto().then(() => refreshView()).catch(e => console.warn(e)); }
    for (const x of (S.histCompleto || [])) { const t = S.byId.get(x.t) || { id: x.t, item: x.item || '(excluído)', valor: x.valor || 0 }; vistos.add(x.t + '_' + x.em); ev.push({ ...x, t, morto: !S.byId.has(x.t) }); }
    for (const t of S.list) for (const h of (t.hist || [])) if (!vistos.has(t.id + '_' + h.em)) ev.push({ ...h, t });
    const lx = Object.entries((S.lixeira && S.lixeira.items) || {}).map(([id, v]) => { try { const o = JSON.parse(v); return { id, ...o }; } catch (e) { return null; } }).filter(Boolean).sort((a, b) => b.em - a.em);
    for (const x of lx) if (!vistos.has(x.id + '_' + x.em)) ev.push({ em: x.em, por: x.por, a: 'excluiu', t: x.t, morto: true });
    ev.sort((a, b) => b.em - a.em);
    const users = [...new Set(ev.map(e => e.por).filter(Boolean))];
    if (histU && !users.includes(histU)) histU = '';
    const vis = ev.filter(e => !histU || e.por === histU).slice(0, 300);
    el.innerHTML = `<div class="filtros card"><label class="tgl">Usuário <select id="hU">${opts(users.map(u => [u, u]), histU, 'Todos')}</select></label><span class="muted small">Toda mudança fica registrada: quem fez, quando e o que mudou. O que é excluído fica 90 dias na lixeira.</span></div>
      <div class="grid lg"><div class="card"><h2>Atividade recente</h2><div class="tlist">${vis.map(e => `<div class="tr hist-row ${e.morto ? '' : 'click'}" ${e.morto ? '' : `data-act="abrir" data-id="${esc(e.t.id)}"`}><div class="td-main"><div class="td-item">${esc(e.t.item)}</div><div class="td-meta">${histTxt(e)}</div></div><div class="td-val"><b class="${cls(e.t.valor)}">${brl2(e.t.valor)}</b></div></div>`).join('') || '<p class="muted pad">Nenhuma alteração registrada ainda. O histórico começa a ser gravado a partir de agora.</p>'}</div></div>
      <div class="card"><h2>Lixeira <span class="muted small">${lx.length} item(ns)</span>${lx.length ? '<button class="btn sm danger ghost rw" data-act="lxVazia">Esvaziar</button>' : ''}</h2><div class="tlist">${lx.slice(0, 200).map(x => `<div class="tr"><div class="td-main"><div class="td-item">${esc(x.t.item)}</div><div class="td-meta"><span>${dBR(x.t.venc)}</span><span>excluído por ${quem(x.por)} em ${quando(x.em)}</span></div></div><div class="td-val"><b class="${cls(x.t.valor)}">${brl2(x.t.valor)}</b><button class="btn sm rw" data-act="lxRest" data-id="${esc(x.id)}">Restaurar</button></div></div>`).join('') || '<p class="muted pad">A lixeira está vazia.</p>'}</div></div></div>`;
    $('#hU', el).onchange = e => { histU = e.target.value; this.refresh(el); };
  }
};
let histU = '';
ACOES.lxRest = d => { const v = S.lixeira?.items?.[d.id]; if (!v) return; const o = JSON.parse(v); if (S.byId.has(d.id)) { toast('Esse lançamento já existe.', 'err'); return; } salvarTitulos([o.t], { acao: 'restaurou' }); gravar(doc(db, 'lixeira', (S.lixeiraOnde || {})[d.id] || hoje().slice(0, 7)), { items: { [d.id]: deleteField() } }, { merge: true }).catch(onErroEscrita); toast('Lançamento restaurado'); };
ACOES.lxVazia = async () => { const ks = Object.keys(S.lixeira?.items || {}); if (ks.length && await confirmar(`Apagar definitivamente os <b>${ks.length}</b> itens da lixeira?`, 'Apagar')) { for (const b of new Set(Object.values(S.lixeiraOnde || {}))) gravar(doc(db, 'lixeira', b), { items: {} }).catch(onErroEscrita); toast('Lixeira esvaziada'); } };

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
  refresh(el) { el.innerHTML = `<button class="menu-item busca-item" data-act="busca">${ic('busca')}<span>Buscar ou perguntar</span></button><div class="card"><h2>Aparência</h2>${seletorTema()}</div>` + NAV.filter(([k]) => !['inicio', 'pagar', 'receber', 'obras'].includes(k)).map(([k, n, ar]) => k === '#' ? `<div class="eyebrow" style="margin:14px 4px 6px">${n}</div>` : `<a class="menu-item" href="#/${k}">${ic(ar ? 'area_' + ar : k)}<span>${n}</span>${ar ? `<small class="muted">${AREAS[ar][1].map(x => x[1]).join(', ')}</small>` : ''}</a>`).join(''); }
};

/* ---------- Dados ---------- */
VIEWS.dados = {
  titulo: 'Dados e backup',
  refresh(el) {
    el.innerHTML = `<div class="grid two">
      <div class="card"><h2>Exportar</h2><p class="muted">Gera uma planilha com Extrato, Contas a Pagar, Contas a Receber, Previsto e Checklist. Faça isso de tempos em tempos como backup.</p><div class="acts left"><button class="btn primary" id="bXls">Exportar Excel</button><button class="btn" id="bJson">Backup completo (.json)</button></div></div>
      <div class="card"><h2>Importar</h2><p class="muted">Importa a planilha <b>SFG Projeto</b> (abas Extrato, Contas a Pagar, Contas a Receber, Previsto e Checklist) ou um backup .json. <b>Substitui todos os dados atuais.</b> Precisa de internet.</p>
        <div class="acts left"><label class="btn file">Importar planilha .xlsx<input type="file" id="iXls" accept=".xlsx,.xlsm" hidden></label><label class="btn file">Restaurar backup .json<input type="file" id="iJson" accept=".json,application/json" hidden></label></div><p id="impStatus" class="muted"></p></div>
      ${cardBackup()}
      ${cardDrive()}
      ${cardAgenda()}
      <div class="card"><h2>Conta</h2><div class="kv"><span>Usuário</span><b>${esc(S.user?.email || '')}</b></div><div class="kv"><span>Lançamentos</span><b>${S.list.length.toLocaleString('pt-BR')}</b></div><div class="kv"><span>Versão</span><b>${VERSAO}</b></div><div class="acts left"><button class="btn danger ghost" id="bSair">Sair deste aparelho</button></div></div>
    </div>`;
    $('#bXls', el).onclick = () => exportarExcel().catch(e => toast(e.message, 'err')); $('#bJson', el).onclick = () => exportarJSON();
    $('#iXls', el).onchange = e => importarArquivo(e.target.files[0], 'xlsx');
    $('#iJson', el).onchange = e => importarArquivo(e.target.files[0], 'json');
    $('#bSair', el).onclick = async () => { if (await confirmar('Sair? Alterações que ainda não sincronizaram podem se perder se você sair sem internet.', 'Sair')) signOut(auth); };
  }
};

async function baixarArquivo(blob, nome) {
  try { const file = new File([blob], nome, { type: blob.type }); if (navigator.canShare && navigator.canShare({ files: [file] }) && /iPhone|iPad|Android/i.test(navigator.userAgent)) { await navigator.share({ files: [file], title: nome }); return; } } catch (e) { if (e.name === 'AbortError') return; }
  const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = nome; document.body.appendChild(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 2000);
}
async function exportarExcel(baixar = true) {
  await carregarLib('xlsx');
  const cab = ['Item', 'R$', 'Juros e Multa', 'NF', 'Local', 'Quadra', 'Lote', 'Destino', 'Data', 'Vencimento', 'Conta', 'Fornecedor', 'Observação', 'Quantidade', 'Unidade', 'Nº NF', 'Anexo (Drive)', 'Empresa', 'CNPJ', 'Conferência CNPJ', 'Pertence a', 'Tipo de aporte', 'Etapa da obra', 'Contrato'];
  const lin = t => [t.item, t.valor, t.juros || 0, !!t.nf, t.local || E().local, t.quadra, isNaN(+t.lote) || t.lote === '' ? t.lote : +t.lote, t.destino, pago(t) ? (t.pag || t.venc) : t.venc, t.venc, t.conta || '', t.forn || '', t.obs || '', t.qtd || '', t.unid || '', t.nfNum || '', anexosDe(t).map(a => a.link).filter(Boolean).join(' '), empPorId(empT(t))?.nome || '', fCnpj(empPorId(empT(t))?.cnpj || ''), !variasEmpresas() ? '' : empDiverge(t) ? 'diferente da obra' : empDef(t) ? 'ok' : 'não definido', ehEmpresa(t) ? 'Empresa' : 'Obra', t.aporteTipo ? tipoAporteNome(t.aporteTipo) : '', t.etapaObra || '', t.contrato ? ((S.contratos || []).find(c => c.id === t.contrato) || {}).forn || t.contrato : ''];
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
  try { abasCadastro(wb); } catch (e) { console.warn(e); }
  const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  const blob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  if (baixar) baixarArquivo(blob, `SFG Financeiro ${hoje()}.xlsx`);
  return blob;
}
/* Módulos registram aqui o que entra no backup completo e como restaurar. */
const BACKUP_EXTRA = [];
const CONFIG_EXTRA = [], BLOCOS_EXTRA = [];
/* Um documento em config/ que o app escuta e guarda no backup. */
function usarConfig(nome) { if (!CONFIG_EXTRA.includes(nome)) CONFIG_EXTRA.push(nome); BACKUP_EXTRA.push({ nome: 'config:' + nome, get: () => S[nome] || null, set: (v, ops) => { if (v) ops.push(b => b.set(doc(db, 'config', nome), v)); } }); }
/* Uma coleção em blocos que o app escuta e guarda no backup. */
function usarBlocos(col, aoMudar) { BLOCOS_EXTRA.push([col, aoMudar]); BACKUP_EXTRA.push({ nome: 'col:' + col, get: () => S['raw_' + col] || null, set: (v, ops) => { if (v) for (const id in v) ops.push(b => b.set(doc(db, col, id), v[id])); } }); }
['recorrencias', 'fluxo', 'docsdrive', 'agenda'].forEach(n => BACKUP_EXTRA.push({ nome: 'config:' + n, get: () => S[n] || null, set: (v, ops) => { if (v) ops.push(b => b.set(doc(db, 'config', n), v)); } }));
['fluxoLotes', 'docsArq'].forEach(c => BACKUP_EXTRA.push({ nome: 'col:' + c, get: () => S['raw_' + c] || null, set: (v, ops) => { if (v) for (const id in v) ops.push(b => b.set(doc(db, c, id), v[id])); } }));
function exportarJSON(baixar = true) {
  const data = { formato: 'sfg-financeiro', versao: 1, exportado: new Date().toISOString(), estrutura: S.estrutura, orcamento: S.orcamento, checklist: S.checklist, vendas: S.vendas, obras: S.obras, fornecedores: S.fornecedores, simulacoes: S.simulacoes, titulos: S.list.map(limpar) };
  data.extras = {}; for (const x of BACKUP_EXTRA) { try { data.extras[x.nome] = x.get(); } catch (e) { console.warn(e); } }
  const blob = new Blob([JSON.stringify(data)], { type: 'application/json' });
  if (baixar) baixarArquivo(blob, `SFG Financeiro backup ${hoje()}.json`);
  return blob;
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
    if (s(r[15])) t.nfNum = s(r[15]); if (s(r[16])) t.anexoLink = s(r[16]);
    if (status === 'pago') { t.pag = dt(r[8]) || venc; t.conta = s(r[10]) || 'BTG'; }
    if (s(r[20]) === 'Empresa') { t.ambito = 'empresa'; t.local = ''; } { const ta = TIPOS_APORTE.find(x => x[1] === s(r[21])); if (ta) t.aporteTipo = ta[0]; } if (s(r[22])) t.etapaObra = s(r[22]);
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
    if (tipo === 'xlsx') { await carregarLib('xlsx'); } if (tipo === 'xlsx') data = lerPlanilha(new Uint8Array(await file.arrayBuffer()));
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
  for (const x of BACKUP_EXTRA) if (data.extras && data.extras[x.nome] !== undefined) { try { x.set(data.extras[x.nome], ops); } catch (e) { console.warn(e); } }
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
  if (u) { S.user = u; iniciarDados(); montarShell(); setTimeout(() => sincronizarDrive(false), 6000); setTimeout(() => carregarLib('chart').catch(() => {}), 2500); }
  else { S.user = null; pararDados(); document.removeEventListener('click', onClick); atual = { nome: null, args: [], view: null }; telaLogin(); }
});
window.__sfg = { S, hoje, E };
