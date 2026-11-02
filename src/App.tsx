import { lazy, Suspense, useCallback, useEffect, useRef, useState, type DragEvent } from 'react';
import type { EditorHandle } from './editors/types';
import {
  ACCEPT, convert, hasFsAccess, isSupported, MIME, openFile, pickFiles, saveBlob, serverStatus, withExt,
  type DocKind, type OpenedFile,
} from './lib/files';

interface Tab {
  file: OpenedFile;
  dirty: boolean;
}

interface Status {
  text: string;
  error?: boolean;
}

const KIND_BADGE: Record<DocKind, { letter: string; label: string }> = {
  pdf: { letter: 'P', label: 'PDF' },
  sheet: { letter: 'X', label: 'Excel' },
  doc: { letter: 'W', label: 'Word' },
  slides: { letter: 'S', label: 'Sunum' },
};

// Editörler ayrı paketlerde: Univer tek başına ~10 MB, yalnızca tablo açılınca yüklenir.
const EDITORS = {
  pdf: lazy(() => import('./editors/PdfEditor').then((m) => ({ default: m.PdfEditor }))),
  sheet: lazy(() => import('./editors/SheetEditor').then((m) => ({ default: m.SheetEditor }))),
  doc: lazy(() => import('./editors/DocEditor').then((m) => ({ default: m.DocEditor }))),
  slides: lazy(() => import('./editors/SlideViewer').then((m) => ({ default: m.SlideViewer }))),
};

type PickedFile = { file: File; handle?: FileSystemFileHandle };

export function App() {
  const [tabs, setTabs] = useState<Tab[]>([]);
  const [activeId, setActiveId] = useState<string | null>(null);
  const [status, setStatus] = useState<Status | null>(null);
  const [busy, setBusy] = useState(false);
  const [soffice, setSoffice] = useState<boolean | null>(null);
  const [dragOver, setDragOver] = useState(false);
  const editors = useRef(new Map<string, EditorHandle>());
  const inputRef = useRef<HTMLInputElement>(null);

  const active = tabs.find((t) => t.file.id === activeId) ?? null;

  useEffect(() => {
    serverStatus().then((s) => setSoffice(s.soffice));
  }, []);

  const addFiles = useCallback(async (items: PickedFile[]) => {
    for (const { file, handle } of items) {
      if (!isSupported(file.name)) {
        setStatus({ text: `Desteklenmeyen dosya: ${file.name}`, error: true });
        continue;
      }
      setBusy(true);
      setStatus({ text: `${file.name} açılıyor…` });
      try {
        const opened = await openFile(file, handle);
        setTabs((prev) => [...prev, { file: opened, dirty: false }]);
        setActiveId(opened.id);
        setStatus(opened.convertedFrom ? { text: `${opened.convertedFrom} → ${opened.name} olarak dönüştürüldü` } : null);
      } catch (err) {
        setStatus({ text: (err as Error).message, error: true });
      } finally {
        setBusy(false);
      }
    }
  }, []);

  const openDialog = useCallback(async () => {
    if (!hasFsAccess) {
      inputRef.current?.click();
      return;
    }
    try {
      await addFiles(await pickFiles());
    } catch (err) {
      if ((err as DOMException).name !== 'AbortError') setStatus({ text: (err as Error).message, error: true });
    }
  }, [addFiles]);

  const updateTab = (id: string, patch: (t: Tab) => Tab) => setTabs((prev) => prev.map((t) => (t.file.id === id ? patch(t) : t)));

  const save = useCallback(async (saveAs: boolean) => {
    if (!active) return;
    const editor = editors.current.get(active.file.id);
    if (!editor) return;
    setBusy(true);
    setStatus({ text: 'Kaydediliyor…' });
    try {
      const blob = await editor.save();
      const result = await saveBlob(blob, active.file.name, active.file.handle, saveAs);
      if (!result) {
        setStatus(null);
        return;
      }
      updateTab(active.file.id, (t) => ({ dirty: false, file: { ...t.file, name: result.name, handle: result.handle ?? t.file.handle } }));
      setStatus({ text: `Kaydedildi: ${result.name}` });
    } catch (err) {
      setStatus({ text: `Kaydedilemedi: ${(err as Error).message}`, error: true });
    } finally {
      setBusy(false);
    }
  }, [active]);

  const exportPdf = useCallback(async () => {
    if (!active) return;
    const editor = editors.current.get(active.file.id);
    if (!editor) return;
    setBusy(true);
    setStatus({ text: 'PDF oluşturuluyor…' });
    try {
      const pdf = await convert(await editor.save(), active.file.name, 'pdf');
      const result = await saveBlob(new Blob([pdf], { type: MIME.pdf }), withExt(active.file.name, 'pdf'), undefined, true);
      setStatus(result ? { text: `PDF kaydedildi: ${result.name}` } : null);
    } catch (err) {
      setStatus({ text: `PDF oluşturulamadı: ${(err as Error).message}`, error: true });
    } finally {
      setBusy(false);
    }
  }, [active]);

  const closeTab = (id: string) => {
    const tab = tabs.find((t) => t.file.id === id);
    if (tab?.dirty && !window.confirm(`${tab.file.name} kaydedilmedi. Yine de kapatılsın mı?`)) return;
    const index = tabs.findIndex((t) => t.file.id === id);
    const rest = tabs.filter((t) => t.file.id !== id);
    setTabs(rest);
    editors.current.delete(id);
    if (activeId === id) setActiveId(rest[Math.min(index, rest.length - 1)]?.file.id ?? null);
  };

  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (!(e.ctrlKey || e.metaKey)) return;
      const key = e.key.toLowerCase();
      if (key === 's') {
        e.preventDefault();
        void save(e.shiftKey);
      } else if (key === 'o') {
        e.preventDefault();
        void openDialog();
      }
    };
    window.addEventListener('keydown', onKey, true);
    return () => window.removeEventListener('keydown', onKey, true);
  }, [save, openDialog]);

  useEffect(() => {
    const onBeforeUnload = (e: BeforeUnloadEvent) => {
      if (tabs.some((t) => t.dirty)) e.preventDefault();
    };
    window.addEventListener('beforeunload', onBeforeUnload);
    return () => window.removeEventListener('beforeunload', onBeforeUnload);
  }, [tabs]);

  const onDrop = (e: DragEvent) => {
    e.preventDefault();
    setDragOver(false);
    // Handle'lar drop olayı sırasında senkron istenmeli; sonradan erişilemez.
    const pending = [...e.dataTransfer.items]
      .filter((item) => item.kind === 'file')
      .map(async (item): Promise<PickedFile | null> => {
        const file = item.getAsFile();
        if (!file) return null;
        const getHandle = (item as DataTransferItem & { getAsFileSystemHandle?: () => Promise<FileSystemHandle | null> }).getAsFileSystemHandle;
        const handle = getHandle ? await getHandle.call(item).catch(() => null) : null;
        return { file, handle: handle?.kind === 'file' ? (handle as FileSystemFileHandle) : undefined };
      });
    Promise.all(pending).then((items) => addFiles(items.filter((i): i is PickedFile => i !== null)));
  };

  return (
    <div
      className={dragOver ? 'app drag-over' : 'app'}
      onDragOver={(e) => {
        e.preventDefault();
        setDragOver(true);
      }}
      onDragLeave={(e) => e.currentTarget === e.target && setDragOver(false)}
      onDrop={onDrop}
    >
      <header className="toolbar main">
        <span className="brand">office-droik</span>
        <button type="button" className="tool" onClick={openDialog} title="Dosya aç (Ctrl+O)">Aç</button>
        <button type="button" className="tool" onClick={() => save(false)} disabled={!active || busy} title="Kaydet (Ctrl+S)">Kaydet</button>
        <button type="button" className="tool" onClick={() => save(true)} disabled={!active || busy} title="Farklı kaydet (Ctrl+Shift+S)">Farklı kaydet</button>
        <button
          type="button"
          className="tool"
          onClick={exportPdf}
          disabled={!active || active.file.kind === 'pdf' || !soffice || busy}
          title={soffice ? 'PDF olarak dışa aktar' : 'PDF dışa aktarımı için LibreOffice gerekli'}
        >
          PDF'e aktar
        </button>
        <span className="spacer" />
        {status && <span className={status.error ? 'status error' : 'status'}>{status.text}</span>}
        {soffice === false && <span className="status error" title="Eski biçimler, Word kaydetme ve PDF dışa aktarma kapalı">LibreOffice yok</span>}
        <input
          ref={inputRef}
          type="file"
          multiple
          accept={ACCEPT}
          hidden
          onChange={(e) => {
            void addFiles([...(e.target.files ?? [])].map((file) => ({ file })));
            e.target.value = '';
          }}
        />
      </header>

      {tabs.length > 0 && (
        <nav className="tabs">
          {tabs.map(({ file, dirty }) => (
            <div key={file.id} className={file.id === activeId ? 'tab active' : 'tab'} onClick={() => setActiveId(file.id)} title={file.name}>
              <span className={`badge badge-${file.kind}`}>{KIND_BADGE[file.kind].letter}</span>
              <span className="tab-name">{file.name}</span>
              {dirty && <span className="dirty" title="Kaydedilmemiş değişiklikler">●</span>}
              <button
                type="button"
                className="tab-close"
                title="Kapat"
                onClick={(e) => {
                  e.stopPropagation();
                  closeTab(file.id);
                }}
              >
                ×
              </button>
            </div>
          ))}
        </nav>
      )}

      <main className="workspace">
        {tabs.length === 0 && <Welcome onOpen={openDialog} soffice={soffice} />}
        {tabs.map(({ file }) => {
          const Editor = EDITORS[file.kind];
          return (
            <section key={file.id} className="editor-pane" hidden={file.id !== activeId}>
              <Suspense fallback={<p className="muted pad">Editör yükleniyor…</p>}>
                <Editor
                  file={file}
                  sofficeAvailable={soffice === true}
                  onDirty={() => updateTab(file.id, (t) => (t.dirty ? t : { ...t, dirty: true }))}
                  ref={(handle: EditorHandle | null) => {
                    if (handle) editors.current.set(file.id, handle);
                  }}
                />
              </Suspense>
            </section>
          );
        })}
      </main>
      {dragOver && <div className="drop-hint">Açmak için bırakın</div>}
    </div>
  );
}

function Welcome({ onOpen, soffice }: { onOpen: () => void; soffice: boolean | null }) {
  return (
    <div className="welcome">
      <h1>office-droik</h1>
      <p>Word, Excel, PowerPoint ve PDF dosyalarını açın ve düzenleyin. Dosyalarınız bilgisayarınızdan çıkmaz.</p>
      <button type="button" className="primary" onClick={onOpen}>Dosya aç</button>
      <p className="muted">veya dosyaları bu pencereye sürükleyip bırakın</p>
      <ul className="formats">
        {(Object.keys(KIND_BADGE) as DocKind[]).map((kind) => (
          <li key={kind}>
            <span className={`badge badge-${kind}`}>{KIND_BADGE[kind].letter}</span>
            {KIND_BADGE[kind].label}
          </li>
        ))}
      </ul>
      <p className="muted small">
        {soffice === null ? 'LibreOffice denetleniyor…'
          : soffice ? 'LibreOffice bulundu: .doc, .xls, .ppt, .odt gibi biçimler ve PDF dışa aktarımı açık.'
          : 'LibreOffice bulunamadı: yalnızca .docx/.xlsx/.pptx/.pdf açılabilir, Word kaydetme kapalı.'}
      </p>
    </div>
  );
}
