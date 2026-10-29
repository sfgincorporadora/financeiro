// Firebase simulado em memória (persistido no localStorage) para testes de interface.
const LS = 'fakefs';
const DEL = { __del: true };
let store = {}; try { store = JSON.parse(localStorage.getItem(LS) || '{}'); } catch (e) {}
const save = () => localStorage.setItem(LS, JSON.stringify(store));
const listeners = new Set();
let user = null; try { user = JSON.parse(localStorage.getItem('fakeuser') || 'null'); } catch (e) {}
const authL = new Set();
export const ALLOWED = ['luismatagalvao@gmail.com'];
const acc = () => store['config/acesso'] || {};
const isEditor = () => user && (ALLOWED.includes(user.email) || (acc().editores || []).includes(user.email));
const allowed = () => user && (isEditor() || (acc().leitores || []).includes(user.email));
const canWrite = path => user && (path === 'config/acesso' ? ALLOWED.includes(user.email) : isEditor());
// app
export function initializeApp(c) { return { c }; }
// auth
export const indexedDBLocalPersistence = {}, browserLocalPersistence = {};
export function initializeAuth() { return { get currentUser() { return user; } }; }
export function connectAuthEmulator() {}
export function onAuthStateChanged(a, fn) { authL.add(fn); setTimeout(() => fn(user), 10); return () => authL.delete(fn); }
export async function signInWithEmailAndPassword(a, email, senha) {
  await new Promise(r => setTimeout(r, 50));
  if (senha !== 'teste123') { const e = new Error('bad'); e.code = 'auth/invalid-credential'; throw e; }
  user = { email, uid: 'u-' + email }; localStorage.setItem('fakeuser', JSON.stringify(user)); authL.forEach(f => f(user)); return { user };
}
export async function signOut() { user = null; localStorage.removeItem('fakeuser'); authL.forEach(f => f(null)); }
export async function sendPasswordResetEmail() {}
// firestore
export function initializeFirestore() { return {}; }
export const persistentLocalCache = () => ({}), persistentMultipleTabManager = () => ({});
export function connectFirestoreEmulator() {}
export const collection = (db, name) => ({ kind: 'col', path: name });
export const doc = (db, c, id) => ({ kind: 'doc', path: c + '/' + id, col: c, id });
export const deleteField = () => DEL;
function merge(a, b) { const o = { ...(a || {}) }; for (const k in b) { const v = b[k]; if (v && v.__del) delete o[k]; else if (v && typeof v === 'object' && !Array.isArray(v) && o[k] && typeof o[k] === 'object') o[k] = merge(o[k], v); else o[k] = v; } return o; }
function strip(o) { if (!o || typeof o !== 'object' || Array.isArray(o)) return o; const r = {}; for (const k in o) { if (o[k] && o[k].__del) continue; r[k] = strip(o[k]); } return r; }
const deny = () => { const e = new Error('Missing or insufficient permissions.'); e.code = 'permission-denied'; return e; };
let prevCols = {};
function notify() { save(); listeners.forEach(l => l.fire()); }
export async function setDoc(ref, data, opt) { if (!canWrite(ref.path)) throw deny(); store[ref.path] = opt && opt.merge ? merge(store[ref.path], data) : strip(data); window.__writes = (window.__writes || 0) + 1; notify(); }
export async function getDocs(ref) { if (!allowed()) throw deny(); const docs = Object.keys(store).filter(p => p.startsWith(ref.path + '/')).map(p => ({ id: p.slice(ref.path.length + 1), data: () => store[p] })); return { docs, forEach: f => docs.forEach(f), size: docs.length }; }
export async function deleteDoc(ref) { if (!canWrite(ref.path)) throw deny(); delete store[ref.path]; notify(); }
export async function getDoc(ref) { if (!allowed()) throw deny(); const d = store[ref.path]; return { exists: () => !!d, data: () => d }; }
export function writeBatch() { const ops = []; return { set: (r, d) => ops.push(() => { store[r.path] = strip(d); }), delete: r => ops.push(() => { delete store[r.path]; }), commit: async () => { if (!isEditor()) throw deny(); ops.forEach(f => f()); notify(); } }; }
export function onSnapshot(ref, a, b, c) {
  const opts = typeof a === 'object' && !a.call ? a : null; const next = opts ? b : a; const err = opts ? c : b;
  const L = { seen: {}, fire() {
    if (!allowed()) { err && err(deny()); return; }
    if (ref.kind === 'doc') { const d = store[ref.path]; next({ exists: () => !!d, data: () => d, metadata: { hasPendingWrites: false } }); return; }
    const ch = []; const now = {};
    for (const p in store) if (p.startsWith(ref.path + '/')) { const id = p.slice(ref.path.length + 1); const s = JSON.stringify(store[p]); now[id] = s; if (L.seen[id] !== s) ch.push({ type: L.seen[id] ? 'modified' : 'added', doc: { id, data: () => store[p] } }); }
    for (const id in L.seen) if (!(id in now)) ch.push({ type: 'removed', doc: { id, data: () => ({}) } });
    L.seen = now;
    const docs = Object.keys(now).map(id => ({ id, data: () => store[ref.path + '/' + id] }));
    next({ docChanges: () => ch, forEach: f => docs.forEach(f), metadata: { hasPendingWrites: false, fromCache: false } });
  } };
  listeners.add(L); setTimeout(() => L.fire(), 5);
  return () => listeners.delete(L);
}
