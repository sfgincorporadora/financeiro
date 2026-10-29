// Leitor de documentos (carregado só quando o usuário usa "Ler boleto ou nota").
// Extrai o texto de PDFs e lê códigos de barras e QR codes de fotos.
import * as pdfjs from 'pdfjs-dist/legacy/build/pdf.mjs';
import { MultiFormatReader, BinaryBitmap, HybridBinarizer, HTMLCanvasElementLuminanceSource, DecodeHintType, BarcodeFormat } from '@zxing/library';

pdfjs.GlobalWorkerOptions.workerSrc = new URL('pdf.worker.min.mjs', document.baseURI).href;

async function textoPDF(buf) {
  const doc = await pdfjs.getDocument({ data: new Uint8Array(buf), isEvalSupported: false }).promise;
  const paginas = [];
  for (let p = 1; p <= Math.min(doc.numPages, 6); p++) {
    const pg = await doc.getPage(p); const tc = await pg.getTextContent();
    const linhas = new Map();
    for (const it of tc.items) { if (!it.str) continue; const y = Math.round(it.transform[5] / 3); const l = linhas.get(y) || []; l.push([it.transform[4], it.str]); linhas.set(y, l); }
    paginas.push([...linhas.entries()].sort((a, b) => b[0] - a[0]).map(([, l]) => l.sort((a, b) => a[0] - b[0]).map(x => x[1]).join(' ').replace(/\s+/g, ' ').trim()).filter(Boolean).join('\n'));
  }
  return paginas.join('\n');
}

function leitor() {
  const r = new MultiFormatReader(); const h = new Map();
  h.set(DecodeHintType.POSSIBLE_FORMATS, [BarcodeFormat.ITF, BarcodeFormat.CODE_128, BarcodeFormat.QR_CODE]);
  h.set(DecodeHintType.TRY_HARDER, true); h.set(DecodeHintType.ALLOWED_LENGTHS, [44, 48]);
  r.setHints(h); return r;
}
async function codigoImagem(file) {
  const url = URL.createObjectURL(file);
  try {
    const img = await new Promise((res, rej) => { const i = new Image(); i.onload = () => res(i); i.onerror = () => rej(new Error('Não consegui abrir a imagem')); i.src = url; });
    const r = leitor(); const achados = [];
    for (const larg of [1600, 1100, 2200]) for (const rot of [0, 90, 270, 180]) {
      const s = Math.min(1, larg / Math.max(img.naturalWidth, img.naturalHeight)); const w = Math.round(img.naturalWidth * s), hh = Math.round(img.naturalHeight * s);
      const c = document.createElement('canvas'); const vira = rot === 90 || rot === 270; c.width = vira ? hh : w; c.height = vira ? w : hh;
      const g = c.getContext('2d'); g.translate(c.width / 2, c.height / 2); g.rotate(rot * Math.PI / 180); g.drawImage(img, -w / 2, -hh / 2, w, hh);
      try { const res = r.decode(new BinaryBitmap(new HybridBinarizer(new HTMLCanvasElementLuminanceSource(c)))); if (res && res.getText()) { achados.push(res.getText()); return achados[0]; } } catch (e) { /* tenta a próxima */ }
    }
    return null;
  } finally { URL.revokeObjectURL(url); }
}
window.SFGLeitor = { textoPDF, codigoImagem };
