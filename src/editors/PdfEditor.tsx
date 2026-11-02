import { useEffect, useImperativeHandle, useRef, useState, type PointerEvent as ReactPointerEvent } from 'react';
import * as pdfjs from 'pdfjs-dist';
import type { PDFPageProxy } from 'pdfjs-dist';
import workerUrl from 'pdfjs-dist/build/pdf.worker.min.mjs?url';
import 'pdfjs-dist/web/pdf_viewer.css';
import { BlendMode, PDFDocument, rgb } from 'pdf-lib';
import fontkit from '@pdf-lib/fontkit';
import fontUrl from 'dejavu-fonts-ttf/ttf/DejaVuSans.ttf?url';
import { MIME } from '../lib/files';
import type { EditorProps } from './types';

pdfjs.GlobalWorkerOptions.workerSrc = workerUrl;

// Ekrandaki metin kutuları PDF'e gömülecek fontla aynı fontla çizilir; böylece konumlar tutarlı kalır.
// DejaVu Sans Türkçe karakterlerin tamamını içerir (pdf-lib'in standart fontları ğ, ş, ı içermez).
const OVERLAY_FONT = 'DroikSans';
const fontLoaded = new FontFace(OVERLAY_FONT, `url(${fontUrl})`).load().then((face) => document.fonts.add(face));
const LINE_HEIGHT = 1.2;
// DejaVu Sans'ta satır kutusunun üstünden taban çizgisine olan mesafe (em cinsinden).
const BASELINE = 0.946;

const ZOOMS = [0.5, 0.75, 1, 1.25, 1.5, 2, 3];

type Tool = 'select' | 'text' | 'highlight';

// Koordinatlar sayfanın 1x ölçekli görünüm alanındadır (sol üst köşe orijin).
type Annot =
  | { id: string; page: number; kind: 'text'; x: number; y: number; size: number; color: string; text: string }
  | { id: string; page: number; kind: 'highlight'; x: number; y: number; w: number; h: number };

interface PdfEditorProps extends EditorProps {
  readOnly?: boolean;
}

function hexToRgb(hex: string) {
  const n = parseInt(hex.slice(1), 16);
  return rgb(((n >> 16) & 255) / 255, ((n >> 8) & 255) / 255, (n & 255) / 255);
}

export function PdfEditor({ file, onDirty, readOnly = false, ref }: PdfEditorProps) {
  const [pages, setPages] = useState<PDFPageProxy[]>([]);
  const [error, setError] = useState<string | null>(null);
  const [scale, setScale] = useState(1.25);
  const [tool, setTool] = useState<Tool>('select');
  const [textSize, setTextSize] = useState(14);
  const [textColor, setTextColor] = useState('#d00000');
  const [annots, setAnnots] = useState<Annot[]>([]);
  const [focusId, setFocusId] = useState<string | null>(null);

  useEffect(() => {
    setPages([]);
    setError(null);
    let cancelled = false;
    const task = pdfjs.getDocument({
      // pdf.js tamponu worker'a devreder; asıl veriyi kaydetme için korumak adına kopyalıyoruz.
      data: new Uint8Array(file.data.slice(0)),
      cMapUrl: '/pdfjs/cmaps/',
      cMapPacked: true,
      standardFontDataUrl: '/pdfjs/standard_fonts/',
      wasmUrl: '/pdfjs/wasm/',
    });
    task.onPassword = (update: (pw: string) => void, reason: number) => {
      const pw = window.prompt(reason === 2 ? 'Parola yanlış, tekrar deneyin:' : 'Bu PDF parola korumalı. Parola:');
      if (pw === null) task.destroy();
      else update(pw);
    };
    task.promise
      .then((doc) => Promise.all(Array.from({ length: doc.numPages }, (_, i) => doc.getPage(i + 1))))
      .then((loaded) => !cancelled && setPages(loaded))
      .catch((err: Error) => !cancelled && setError(err.message));
    return () => {
      cancelled = true;
      task.destroy();
    };
  }, [file.data]);

  useImperativeHandle(ref, () => ({
    async save() {
      if (annots.length === 0) return new Blob([file.data], { type: MIME.pdf });
      const pdf = await PDFDocument.load(file.data, { ignoreEncryption: true });
      pdf.registerFontkit(fontkit);
      const font = await pdf.embedFont(await (await fetch(fontUrl)).arrayBuffer(), { subset: true });
      const pdfPages = pdf.getPages();

      for (const a of annots) {
        const target = pdfPages[a.page];
        const viewport = pages[a.page].getViewport({ scale: 1 });
        if (a.kind === 'highlight') {
          const [x1, y1] = viewport.convertToPdfPoint(a.x, a.y);
          const [x2, y2] = viewport.convertToPdfPoint(a.x + a.w, a.y + a.h);
          target.drawRectangle({
            x: Math.min(x1, x2),
            y: Math.min(y1, y2),
            width: Math.abs(x2 - x1),
            height: Math.abs(y2 - y1),
            color: rgb(1, 0.87, 0),
            opacity: 0.4,
            blendMode: BlendMode.Multiply,
          });
        } else {
          a.text.split('\n').forEach((line, i) => {
            const [x, y] = viewport.convertToPdfPoint(a.x, a.y + (BASELINE + LINE_HEIGHT * i) * a.size);
            target.drawText(line, { x, y, size: a.size, font, color: hexToRgb(a.color) });
          });
        }
      }
      return new Blob([(await pdf.save()) as Uint8Array<ArrayBuffer>], { type: MIME.pdf });
    },
  }), [annots, pages, file.data]);

  const changeAnnots = (update: (prev: Annot[]) => Annot[]) => {
    setAnnots(update);
    onDirty();
  };

  const zoomIndex = ZOOMS.indexOf(scale);

  return (
    <div className="pdf-editor">
      <div className="toolbar sub">
        {!readOnly && (
          <>
            <button type="button" className={tool === 'select' ? 'tool active' : 'tool'} onClick={() => setTool('select')} title="Metin seç ve kopyala">
              ⌶ Seç
            </button>
            <button type="button" className={tool === 'text' ? 'tool active' : 'tool'} onClick={() => setTool('text')} title="Tıkladığın yere metin ekle">
              T Metin ekle
            </button>
            <button type="button" className={tool === 'highlight' ? 'tool active' : 'tool'} onClick={() => setTool('highlight')} title="Sürükleyerek vurgula">
              ▬ Vurgula
            </button>
            {tool === 'text' && (
              <>
                <select value={textSize} onChange={(e) => setTextSize(Number(e.target.value))} title="Yazı boyutu">
                  {[8, 10, 12, 14, 18, 24, 32].map((s) => (
                    <option key={s} value={s}>{s} pt</option>
                  ))}
                </select>
                <label className="tool color" title="Yazı rengi">
                  A
                  <input type="color" value={textColor} onChange={(e) => setTextColor(e.target.value)} />
                </label>
              </>
            )}
            <span className="sep" />
          </>
        )}
        <button type="button" className="tool" disabled={zoomIndex <= 0} onClick={() => setScale(ZOOMS[zoomIndex - 1])} title="Uzaklaştır">
          −
        </button>
        <span className="zoom-label">{Math.round(scale * 100)}%</span>
        <button type="button" className="tool" disabled={zoomIndex >= ZOOMS.length - 1} onClick={() => setScale(ZOOMS[zoomIndex + 1])} title="Yakınlaştır">
          +
        </button>
        <span className="sep" />
        <span className="muted">{pages.length ? `${pages.length} sayfa` : ''}</span>
        {annots.length > 0 && <span className="muted"> · {annots.length} ekleme (kaydedince PDF'e yazılır)</span>}
      </div>
      {error && <div className="editor-error">PDF açılamadı: {error}</div>}
      <div className="pdf-scroll">
        {pages.map((page, i) => (
          <PageView
            key={i}
            page={page}
            index={i}
            scale={scale}
            tool={readOnly ? 'select' : tool}
            readOnly={readOnly}
            annots={annots.filter((a) => a.page === i)}
            focusId={focusId}
            onAdd={(a) => {
              changeAnnots((prev) => [...prev, a]);
              if (a.kind === 'text') setFocusId(a.id);
            }}
            onText={(id, text) => changeAnnots((prev) => prev.map((a) => (a.id === id && a.kind === 'text' ? { ...a, text } : a)))}
            onRemove={(id) => changeAnnots((prev) => prev.filter((a) => a.id !== id))}
            newText={{ size: textSize, color: textColor }}
          />
        ))}
      </div>
    </div>
  );
}

interface PageViewProps {
  page: PDFPageProxy;
  index: number;
  scale: number;
  tool: Tool;
  readOnly: boolean;
  annots: Annot[];
  focusId: string | null;
  newText: { size: number; color: string };
  onAdd(a: Annot): void;
  onText(id: string, text: string): void;
  onRemove(id: string): void;
}

function PageView({ page, index, scale, tool, readOnly, annots, focusId, newText, onAdd, onText, onRemove }: PageViewProps) {
  const wrapRef = useRef<HTMLDivElement>(null);
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const textRef = useRef<HTMLDivElement>(null);
  const [visible, setVisible] = useState(false);
  const [draft, setDraft] = useState<{ x0: number; y0: number; x: number; y: number } | null>(null);
  const viewport = page.getViewport({ scale });

  // Sayfalar ekrana yaklaştıkça çizilir; yüzlerce sayfalık PDF'ler de hızlı açılır.
  useEffect(() => {
    const el = wrapRef.current!;
    const observer = new IntersectionObserver(([entry]) => entry.isIntersecting && setVisible(true), { rootMargin: '800px 0px' });
    observer.observe(el);
    return () => observer.disconnect();
  }, []);

  useEffect(() => {
    if (!visible) return;
    const canvas = canvasRef.current!;
    const dpr = window.devicePixelRatio || 1;
    canvas.width = Math.floor(viewport.width * dpr);
    canvas.height = Math.floor(viewport.height * dpr);
    const renderTask = page.render({ canvas, viewport, transform: dpr !== 1 ? [dpr, 0, 0, dpr, 0, 0] : undefined });
    renderTask.promise.catch(() => {});

    const container = textRef.current!;
    container.replaceChildren();
    const textLayer = new pdfjs.TextLayer({ textContentSource: page.streamTextContent(), container, viewport });
    textLayer.render().catch(() => {});
    return () => {
      renderTask.cancel();
      textLayer.cancel();
    };
    // viewport her render'da yeni nesne; ölçek değişimi yeterli tetikleyici.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [visible, scale, page]);

  const toPage = (e: ReactPointerEvent) => {
    const rect = wrapRef.current!.getBoundingClientRect();
    return { x: (e.clientX - rect.left) / scale, y: (e.clientY - rect.top) / scale };
  };

  const onPointerDown = (e: ReactPointerEvent) => {
    if (e.target !== e.currentTarget) return;
    const p = toPage(e);
    if (tool === 'text') {
      // Varsayılan davranış odağı tıklanan katmana taşır ve yeni kutu boş kalıp silinirdi.
      e.preventDefault();
      onAdd({ id: crypto.randomUUID(), page: index, kind: 'text', x: p.x, y: p.y - newText.size * 0.6, size: newText.size, color: newText.color, text: '' });
    } else if (tool === 'highlight') {
      e.currentTarget.setPointerCapture(e.pointerId);
      setDraft({ x0: p.x, y0: p.y, x: p.x, y: p.y });
    }
  };
  const onPointerMove = (e: ReactPointerEvent) => {
    if (draft) setDraft({ ...draft, ...toPage(e) });
  };
  const onPointerUp = () => {
    if (!draft) return;
    const w = Math.abs(draft.x - draft.x0);
    const h = Math.abs(draft.y - draft.y0);
    if (w > 2 && h > 2) {
      onAdd({ id: crypto.randomUUID(), page: index, kind: 'highlight', x: Math.min(draft.x, draft.x0), y: Math.min(draft.y, draft.y0), w, h });
    }
    setDraft(null);
  };

  const pageStyle = {
    width: viewport.width,
    height: viewport.height,
    '--scale-factor': scale,
    '--user-unit': 1,
    '--total-scale-factor': scale,
    '--scale-round-x': '1px',
    '--scale-round-y': '1px',
  } as React.CSSProperties;

  return (
    <div ref={wrapRef} className="pdf-page" style={pageStyle} data-page={index + 1}>
      <canvas ref={canvasRef} style={{ width: viewport.width, height: viewport.height }} />
      <div ref={textRef} className="textLayer" />
      <div
        className={`annot-layer tool-${tool}`}
        onPointerDown={onPointerDown}
        onPointerMove={onPointerMove}
        onPointerUp={onPointerUp}
      >
        {annots.map((a) =>
          a.kind === 'highlight' ? (
            <div key={a.id} className="annot annot-highlight" style={{ left: a.x * scale, top: a.y * scale, width: a.w * scale, height: a.h * scale }}>
              {!readOnly && <button type="button" className="annot-del" title="Sil" onClick={() => onRemove(a.id)}>×</button>}
            </div>
          ) : (
            <div key={a.id} className="annot annot-text-wrap" style={{ left: a.x * scale, top: a.y * scale }}>
              <div
                className="annot-text"
                contentEditable={!readOnly}
                suppressContentEditableWarning
                spellCheck={false}
                style={{ fontFamily: OVERLAY_FONT, fontSize: a.size * scale, lineHeight: LINE_HEIGHT, color: a.color }}
                ref={(el) => {
                  if (el && a.id === focusId && document.activeElement !== el && !a.text) {
                    fontLoaded.finally(() => el.focus({ preventScroll: true }));
                  }
                }}
                onBlur={(e) => {
                  const text = e.currentTarget.innerText.replace(/\n$/, '');
                  if (!text.trim()) onRemove(a.id);
                  else if (text !== a.text) onText(a.id, text);
                }}
              >
                {a.text}
              </div>
              {!readOnly && <button type="button" className="annot-del" title="Sil" onMouseDown={(e) => e.preventDefault()} onClick={() => onRemove(a.id)}>×</button>}
            </div>
          ),
        )}
        {draft && (
          <div
            className="annot annot-highlight draft"
            style={{
              left: Math.min(draft.x, draft.x0) * scale,
              top: Math.min(draft.y, draft.y0) * scale,
              width: Math.abs(draft.x - draft.x0) * scale,
              height: Math.abs(draft.y - draft.y0) * scale,
            }}
          />
        )}
      </div>
    </div>
  );
}
