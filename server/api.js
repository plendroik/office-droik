// LibreOffice tabanlı dönüştürme API'si.
// Hem Vite geliştirme sunucusuna hem de üretim sunucusuna (start.js) middleware olarak takılır.
import { spawn } from 'node:child_process';
import { existsSync } from 'node:fs';
import { mkdtemp, readdir, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import express from 'express';

const MAX_UPLOAD = 200 * 1024 * 1024;
const TIMEOUT_MS = 120_000;

// Windows'ta soffice.com konsol sürümüdür ve dönüştürme bitene kadar bekler; soffice.exe beklemeden dönebilir.
const SOFFICE_CANDIDATES = [
  process.env.SOFFICE_PATH,
  'C:\\Program Files\\LibreOffice\\program\\soffice.com',
  'C:\\Program Files\\LibreOffice\\program\\soffice.exe',
  'C:\\Program Files (x86)\\LibreOffice\\program\\soffice.com',
  '/Applications/LibreOffice.app/Contents/MacOS/soffice',
  '/usr/bin/soffice',
  '/usr/bin/libreoffice',
  '/snap/bin/libreoffice',
];

const TARGETS = {
  pdf: 'pdf',
  docx: 'docx:MS Word 2007 XML',
  xlsx: 'xlsx:Calc MS Excel 2007 XML',
  pptx: 'pptx:Impress MS PowerPoint 2007 XML',
};

export function findSoffice() {
  return SOFFICE_CANDIDATES.find((p) => p && existsSync(p)) ?? null;
}

class HttpError extends Error {
  constructor(status, message) {
    super(message);
    this.status = status;
  }
}

// soffice.com çıkınca soffice.bin bir süre daha yaşayabilir; yalnızca üst süreci öldürmek yetmez.
function killTree(child) {
  if (process.platform === 'win32') spawn('taskkill', ['/pid', String(child.pid), '/T', '/F'], { windowsHide: true });
  else child.kill('SIGKILL');
}

function run(cmd, args) {
  return new Promise((resolve, reject) => {
    const child = spawn(cmd, args, { windowsHide: true });
    let stderr = '';
    child.stderr.on('data', (d) => (stderr += d));
    const timer = setTimeout(() => {
      killTree(child);
      reject(new HttpError(504, 'LibreOffice zaman aşımına uğradı'));
    }, TIMEOUT_MS);
    child.on('error', (err) => {
      clearTimeout(timer);
      reject(err);
    });
    child.on('close', (code) => {
      clearTimeout(timer);
      if (code === 0) resolve();
      else reject(new HttpError(500, `LibreOffice hata kodu ${code}: ${stderr.trim()}`));
    });
  });
}

async function convert(input, name, to) {
  const soffice = findSoffice();
  if (!soffice) {
    throw new HttpError(503, 'LibreOffice bulunamadı. Kurun veya SOFFICE_PATH ortam değişkenini ayarlayın.');
  }
  const dir = await mkdtemp(path.join(tmpdir(), 'office-droik-'));
  try {
    const ext = path.extname(name).toLowerCase() || '.bin';
    const inFile = path.join(dir, `input${ext}`);
    const outDir = path.join(dir, 'out');
    await writeFile(inFile, input);

    // Her dönüşüm kendi profilini kullanır. Ortak profilde, kapanmakta olan önceki soffice.bin'e
    // bağlanan yeni işlem süresiz takılabiliyor; ayrı profil ~1-2 sn ek maliyetle bunu önler
    // ve dönüşümlerin paralel çalışmasına izin verir.
    const profile = pathToFileURL(path.join(dir, 'profile')).href;
    const args = ['--headless', '--norestore', `-env:UserInstallation=${profile}`];
    // HTML varsayılan olarak Writer/Web ile açılır; .docx'e yazabilmek için Writer filtresi gerekir.
    if (ext === '.html' || ext === '.htm') args.push('--infilter=HTML (StarWriter)');
    args.push('--convert-to', TARGETS[to], '--outdir', outDir, inFile);

    await run(soffice, args);
    const produced = await readdir(outDir).catch(() => []);
    if (produced.length === 0) throw new HttpError(500, 'Dönüştürme çıktı üretmedi');
    return await readFile(path.join(outDir, produced[0]));
  } finally {
    // soffice.bin profil dosyalarını kısa süre kilitli tutabilir; silme başarısız olursa geçici klasörde kalır.
    await rm(dir, { recursive: true, force: true, maxRetries: 5, retryDelay: 200 }).catch(() => {});
  }
}

function readBody(req) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    let size = 0;
    req.on('data', (chunk) => {
      size += chunk.length;
      if (size > MAX_UPLOAD) {
        reject(new HttpError(413, 'Dosya çok büyük'));
        req.destroy();
        return;
      }
      chunks.push(chunk);
    });
    req.on('end', () => resolve(Buffer.concat(chunks)));
    req.on('error', reject);
  });
}

function sendJson(res, status, body) {
  res.statusCode = status;
  res.setHeader('Content-Type', 'application/json; charset=utf-8');
  res.end(JSON.stringify(body));
}

/** `/api` altına takılır: GET /status, POST /convert?to=pdf&name=dosya.docx */
export function apiMiddleware() {
  return async (req, res, next) => {
    const url = new URL(req.url, 'http://localhost');
    try {
      if (url.pathname === '/status' && req.method === 'GET') {
        return sendJson(res, 200, { soffice: findSoffice() !== null });
      }
      if (url.pathname === '/convert' && req.method === 'POST') {
        const to = url.searchParams.get('to');
        const name = url.searchParams.get('name') ?? 'dosya';
        if (!to || !(to in TARGETS)) throw new HttpError(400, `Desteklenmeyen hedef biçim: ${to}`);
        const output = await convert(await readBody(req), name, to);
        res.statusCode = 200;
        res.setHeader('Content-Type', 'application/octet-stream');
        return res.end(output);
      }
      next();
    } catch (err) {
      sendJson(res, err.status ?? 500, { error: err.message });
    }
  };
}

// pdf.js'in CMap, standart font ve wasm dosyaları (Asyaca metinler ve bazı görüntü biçimleri için gerekli).
export function pdfjsAssetsMiddleware() {
  const root = path.join(path.dirname(fileURLToPath(import.meta.url)), '..', 'node_modules', 'pdfjs-dist');
  const router = express.Router();
  for (const dir of ['cmaps', 'standard_fonts', 'wasm']) {
    router.use(`/${dir}`, express.static(path.join(root, dir), { immutable: true, maxAge: '30d' }));
  }
  return router;
}
