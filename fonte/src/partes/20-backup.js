/* =========================================================
   Backup automático no Google Drive (toda semana, guarda os 12 últimos)
   ========================================================= */
usarConfig('backup');
const BKP = { rodando: false, erro: '' };
const BKP_DIAS = 7, BKP_MANTER = 12;
async function pastaBackups() {
  const raiz = await raizSFG(); const c = S.backup || {};
  if (c.pasta && await pastaExiste(c.pasta)) return c.pasta;
  const p = await novaPasta('Backups', raiz); S.backup = { ...c, pasta: p.id, link: p.webViewLink };
  gravar(doc(db, 'config', 'backup'), S.backup).catch(onErroEscrita); return p.id;
}
async function apagarNoDrive(id) {
  const tk = await driveToken(false); if (!tk || !id) return;
  await fetch('https://www.googleapis.com/drive/v3/files/' + encodeURIComponent(id), { method: 'DELETE', headers: { Authorization: 'Bearer ' + tk } }).catch(() => {});
}
async function fazerBackup(interativo) {
  if (BKP.rodando || somenteLeitura() || !driveConfigurado() || !navigator.onLine || !S.loaded.titulos) return;
  let tk; try { tk = await driveToken(interativo); } catch (e) { if (interativo) toast(e.message, 'err'); return; }
  if (!tk) return;
  BKP.rodando = true; BKP.erro = ''; refreshView();
  try {
    const pasta = await pastaBackups();
    const nome = `SFG backup ${hoje()} ${new Date().toTimeString().slice(0, 5).replace(':', 'h')}`;
    const j = await enviarArquivoDrive(exportarJSON(false), nome + '.json', pasta);
    const x = await enviarArquivoDrive(await exportarExcel(false), nome + '.xlsx', pasta);
    const c = { ...(S.backup || {}) }; const lista = [...(c.arquivos || []), { d: Date.now(), j: j.id, x: x.id, n: S.list.length }];
    while (lista.length > BKP_MANTER) { const v = lista.shift(); await apagarNoDrive(v.j); await apagarNoDrive(v.x); }
    S.backup = { ...c, arquivos: lista, ultima: Date.now() }; gravar(doc(db, 'config', 'backup'), S.backup).catch(onErroEscrita);
    if (interativo) toast('Backup salvo no Google Drive');
  } catch (e) { BKP.erro = e.message === 'drive-desconectado' ? 'A autorização do Google expirou. Toque em Fazer backup agora.' : e.message; if (interativo) toast(BKP.erro, 'err'); }
  finally { BKP.rodando = false; refreshView(); }
}
const backupAtrasado = () => driveConfigurado() && !somenteLeitura() && (!S.backup || !S.backup.ultima || Date.now() - S.backup.ultima > (BKP_DIAS + 1) * 864e5);
/* Verifica a cada 10 minutos: se o Google estiver conectado e o último backup tiver mais de 7 dias, faz sozinho. */
setInterval(() => { if (S.user && tokenTem(DRIVE_SCOPE) && (!S.backup?.ultima || Date.now() - S.backup.ultima > BKP_DIAS * 864e5)) fazerBackup(false); }, 10 * 60000);
setTimeout(() => { if (S.user && tokenTem(DRIVE_SCOPE) && (!S.backup?.ultima || Date.now() - S.backup.ultima > BKP_DIAS * 864e5)) fazerBackup(false); }, 20000);
ACOES.backup = () => fazerBackup(true);
function cardBackup() {
  if (!driveConfigurado()) return '';
  const c = S.backup || {}; const ult = c.ultima ? quando(c.ultima) : 'nunca';
  return `<div class="card"><h2>Backup automático no Drive</h2>
    <p class="muted small">Toda semana o app salva uma cópia completa (arquivo .json para restaurar e planilha .xlsx para consultar) na pasta SFG Financeiro / Backups do seu Google Drive e guarda as ${BKP_MANTER} mais recentes. Acontece sozinho quando o app está aberto e conectado ao Google.</p>
    <div class="kv"><span>Último backup</span><b class="${backupAtrasado() ? 'neg' : ''}">${ult}</b></div><div class="kv"><span>Cópias guardadas</span><b>${(c.arquivos || []).length}</b></div>
    ${BKP.erro ? `<p class="erro small">${esc(BKP.erro)}</p>` : ''}
    <div class="acts left rw"><button class="btn primary" data-act="backup" ${BKP.rodando ? 'disabled' : ''}>${BKP.rodando ? 'Salvando…' : 'Fazer backup agora'}</button>${c.link ? `<a class="btn" href="${esc(c.link)}" target="_blank" rel="noopener">Abrir pasta de backups</a>` : ''}</div>
    <p class="muted small">Para restaurar: baixe o .json da pasta e use "Restaurar backup .json" acima.</p></div>`;
}
