import { useEffect, useImperativeHandle, useMemo, useRef, useState } from 'react';
import { init } from 'pptx-preview';
import { convert, MIME } from '../lib/files';
import { PdfEditor } from './PdfEditor';
import type { EditorProps } from './types';

type Mode = 'fast' | 'hq';

export function SlideViewer({ file, sofficeAvailable, onDirty, ref }: EditorProps) {
  // pptx-preview yer tutucu yazı boyutlarını her zaman doğru çözemiyor; LibreOffice varsa birebir görünüm varsayılan.
  const [mode, setMode] = useState<Mode>(sofficeAvailable ? 'hq' : 'fast');
  const [pdf, setPdf] = useState<ArrayBuffer | null>(null);
  const [error, setError] = useState<string | null>(null);
  const hostRef = useRef<HTMLDivElement>(null);

  useImperativeHandle(ref, () => ({
    // Sunum düzenleme henüz yok; kaydetme dosyanın kendisini yazar (farklı kaydet / PDF dışa aktarım için).
    async save() {
      return new Blob([file.data], { type: MIME.pptx });
    },
  }));

  useEffect(() => {
    if (mode !== 'fast') return;
    const host = hostRef.current!;
    // Liste modunda height kaydırılabilir görünüm alanının yüksekliğidir; slayt oranı dosyadan gelir.
    const previewer = init(host, { width: Math.min(host.clientWidth - 48, 1100), height: host.clientHeight - 48, mode: 'list' });
    previewer.preview(file.data.slice(0)).catch((err: Error) => setError(err.message));
    return () => {
      previewer.destroy();
      host.replaceChildren();
    };
  }, [file.data, mode]);

  useEffect(() => {
    if (mode !== 'hq' || pdf) return;
    convert(file.data, file.name, 'pdf')
      .then(setPdf)
      .catch((err: Error) => {
        setError(err.message);
        setMode('fast');
      });
  }, [mode, pdf, file.data, file.name]);

  const pdfFile = useMemo(() => (pdf ? { ...file, data: pdf, kind: 'pdf' as const } : null), [pdf, file]);

  return (
    <div className="slide-viewer">
      <div className="toolbar sub">
        <button type="button" className={mode === 'fast' ? 'tool active' : 'tool'} onClick={() => setMode('fast')}>
          Hızlı önizleme
        </button>
        <button
          type="button"
          className={mode === 'hq' ? 'tool active' : 'tool'}
          disabled={!sofficeAvailable}
          title={sofficeAvailable ? 'LibreOffice ile birebir görünüm' : 'LibreOffice bulunamadı'}
          onClick={() => setMode('hq')}
        >
          Yüksek kalite (LibreOffice)
        </button>
        <span className="muted">Sunum düzenleme sonraki sürümde.</span>
      </div>
      {error && <div className="editor-error">{error}</div>}
      {mode === 'fast' && <div ref={hostRef} className="slide-host" />}
      {mode === 'hq' && (pdfFile ? <PdfEditor file={pdfFile} sofficeAvailable={sofficeAvailable} onDirty={onDirty} readOnly /> : <p className="muted pad">LibreOffice ile dönüştürülüyor…</p>)}
    </div>
  );
}
