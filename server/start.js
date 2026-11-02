// Üretim sunucusu: `npm run build` sonrası dist/ klasörünü ve API'yi tek porttan sunar.
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import express from 'express';
import { apiMiddleware, findSoffice, pdfjsAssetsMiddleware } from './api.js';

const root = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const port = Number(process.env.PORT ?? 4173);

const app = express();
app.use('/api', apiMiddleware());
app.use('/pdfjs', pdfjsAssetsMiddleware());
app.use(express.static(path.join(root, 'dist')));

app.listen(port, '127.0.0.1', () => {
  console.log(`office-droik çalışıyor: http://localhost:${port}`);
  console.log(findSoffice() ? `LibreOffice: ${findSoffice()}` : 'Uyarı: LibreOffice bulunamadı, dönüştürme özellikleri kapalı.');
});
