# office-droik

Word, Excel, PowerPoint ve PDF dosyalarını açıp düzenleyen açık kaynak ofis uygulaması.
Tarayıcıda çalışır, dosyalarınız bilgisayarınızdan çıkmaz.

| Biçim | Görüntüleme | Düzenleme |
|---|---|---|
| **PDF** | Tüm sayfalar, metin seçme/kopyalama, yakınlaştırma | Metin ekleme (Türkçe karakter destekli), vurgulama |
| **Excel** (.xlsx, .xls, .ods, .csv) | Biçimler, birleştirilmiş hücreler, dondurulmuş bölmeler, çoklu sayfa | Hücre düzenleme, formüller (hesap motoru dahil), biçimlendirme |
| **Word** (.docx, .doc, .odt, .rtf) | Başlıklar, listeler, tablolar, resimler | Zengin metin düzenleme, tablolar, hizalama, renk |
| **PowerPoint** (.pptx, .ppt, .odp) | LibreOffice ile birebir görünüm veya hızlı önizleme | Henüz yok |

Her belge **PDF olarak dışa aktarılabilir**.

## Kurulum

Gerekenler:

- [Node.js](https://nodejs.org) 20 veya üstü
- [LibreOffice](https://www.libreoffice.org) (önerilir). Eski biçimleri açmak, Word olarak kaydetmek ve PDF dışa aktarmak için gerekir. Bulunamazsa .docx/.xlsx/.pptx/.pdf yine açılır.

```bash
git clone https://github.com/plendroik/office-droik.git
cd office-droik
npm install
npm run dev
```

Tarayıcıda http://localhost:5173 adresini açın.

Üretim sürümü için:

```bash
npm run build
npm start        # http://localhost:4173
```

LibreOffice standart dışı bir yerdeyse yolunu belirtin:

```bash
SOFFICE_PATH="/opt/libreoffice/program/soffice" npm start
```

## Kullanım

- **Aç**: `Ctrl+O`, ya da dosyaları pencereye sürükleyip bırakın. Birden fazla dosya sekmelerde açılır.
- **Kaydet**: `Ctrl+S`. Chrome ve Edge'de dosya yerinde güncellenir, diğer tarayıcılarda indirilir.
- **Farklı kaydet**: `Ctrl+Shift+S`

## Nasıl çalışır

```
Tarayıcı (React)                          Node sunucusu
├─ PDF    → pdf.js + pdf-lib              └─ /api/convert → LibreOffice (headless)
├─ Excel  → Univer + ExcelJS                  .doc/.xls/.ppt → .docx/.xlsx/.pptx
├─ Word   → TipTap + mammoth                  HTML → .docx (Word kaydetme)
└─ Sunum  → LibreOffice PDF / pptx-preview    her biçim → .pdf
```

- `server/api.js`: LibreOffice dönüştürme API'si. Her dönüşüm kendi geçici profilini kullanır.
- `src/lib/xlsx.ts`: ExcelJS ↔ Univer dönüştürücü (değerler, formüller, stiller, birleştirmeler, sütun genişlikleri).
- `src/editors/`: Her biçimin editörü.

## Bilinen sınırlamalar

- **Excel**: Grafikler, resimler, koşullu biçimlendirme ve veri doğrulama kaydederken korunmaz.
- **Word**: Belge HTML üzerinden düzenlenir. Sayfa düzeni, üst/alt bilgi, yazı tipleri ve karmaşık biçimler sadeleşebilir.
- **PowerPoint**: Şimdilik yalnızca görüntüleme.
- **PDF**: Var olan metin düzenlenemez, üzerine ekleme yapılır. Döndürülmüş sayfalarda eklenen metin döndürülmez.

## Lisans

[MPL-2.0](LICENSE)
