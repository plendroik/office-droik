import { useEffect, useImperativeHandle, useState } from 'react';
import { EditorContent, useEditor, type Editor } from '@tiptap/react';
import StarterKit from '@tiptap/starter-kit';
import { TableKit } from '@tiptap/extension-table';
import TextAlign from '@tiptap/extension-text-align';
import { TextStyleKit } from '@tiptap/extension-text-style';
import Image from '@tiptap/extension-image';
import mammoth from 'mammoth';
import { convert, MIME } from '../lib/files';
import type { EditorProps } from './types';

// LibreOffice'in HTML'i Word'e çevirirken tabloları ve resimleri düzgün yerleştirmesi için.
const EXPORT_CSS = `
  body { font-family: Calibri, Arial, sans-serif; font-size: 11pt; }
  table { border-collapse: collapse; }
  td, th { border: 1px solid #000; padding: 4px; vertical-align: top; }
  img { max-width: 100%; }
`;

function exportHtml(body: string): string {
  return `<!DOCTYPE html><html><head><meta charset="utf-8"><style>${EXPORT_CSS}</style></head><body>${body}</body></html>`;
}

export function DocEditor({ file, sofficeAvailable, onDirty, ref }: EditorProps) {
  const [loading, setLoading] = useState(true);
  const [warnings, setWarnings] = useState<string[]>([]);

  const editor = useEditor({
    extensions: [
      StarterKit,
      TableKit.configure({ table: { resizable: true } }),
      TextAlign.configure({ types: ['heading', 'paragraph'] }),
      TextStyleKit,
      Image.configure({ allowBase64: true }),
    ],
    shouldRerenderOnTransaction: true,
    onUpdate: ({ transaction }) => {
      if (transaction.docChanged) onDirty();
    },
  });

  useEffect(() => {
    if (!editor) return;
    let cancelled = false;
    setLoading(true);
    mammoth
      .convertToHtml({ arrayBuffer: file.data }, { styleMap: ['u => u'] })
      .then((result) => {
        if (cancelled) return;
        editor.commands.setContent(result.value || '<p></p>', { emitUpdate: false });
        setWarnings(result.messages.filter((m) => m.type === 'warning').map((m) => m.message));
      })
      .catch((err: Error) => !cancelled && setWarnings([`Dosya okunamadı: ${err.message}`]))
      .finally(() => !cancelled && setLoading(false));
    return () => {
      cancelled = true;
    };
  }, [editor, file.data]);

  useImperativeHandle(ref, () => ({
    async save() {
      if (!editor) throw new Error('Editör hazır değil');
      if (!sofficeAvailable) throw new Error('Word olarak kaydetmek için LibreOffice gerekli');
      const docx = await convert(exportHtml(editor.getHTML()), 'belge.html', 'docx');
      return new Blob([docx], { type: MIME.docx });
    },
  }));

  return (
    <div className="doc-editor">
      {editor && <DocToolbar editor={editor} />}
      {warnings.length > 0 && (
        <div className="editor-note">
          Bazı biçimlendirmeler aktarılamadı ({warnings.length} uyarı). Karmaşık sayfa düzenleri sadeleşmiş olabilir.
        </div>
      )}
      <div className="doc-scroll">
        <div className="doc-page">{loading ? <p className="muted">Yükleniyor…</p> : <EditorContent editor={editor} />}</div>
      </div>
    </div>
  );
}

function DocToolbar({ editor }: { editor: Editor }) {
  const chain = () => editor.chain().focus();
  const btn = (label: string, title: string, active: boolean, run: () => void) => (
    <button type="button" title={title} className={active ? 'tool active' : 'tool'} onMouseDown={(e) => e.preventDefault()} onClick={run}>
      {label}
    </button>
  );
  const blockValue = editor.isActive('heading', { level: 1 }) ? 'h1'
    : editor.isActive('heading', { level: 2 }) ? 'h2'
    : editor.isActive('heading', { level: 3 }) ? 'h3'
    : 'p';

  return (
    <div className="toolbar sub">
      {btn('↶', 'Geri al (Ctrl+Z)', false, () => chain().undo().run())}
      {btn('↷', 'Yinele (Ctrl+Y)', false, () => chain().redo().run())}
      <span className="sep" />
      <select
        value={blockValue}
        onChange={(e) => {
          const v = e.target.value;
          if (v === 'p') chain().setParagraph().run();
          else chain().toggleHeading({ level: Number(v[1]) as 1 | 2 | 3 }).run();
        }}
      >
        <option value="p">Normal metin</option>
        <option value="h1">Başlık 1</option>
        <option value="h2">Başlık 2</option>
        <option value="h3">Başlık 3</option>
      </select>
      <span className="sep" />
      {btn('K', 'Kalın (Ctrl+B)', editor.isActive('bold'), () => chain().toggleBold().run())}
      {btn('İ', 'İtalik (Ctrl+I)', editor.isActive('italic'), () => chain().toggleItalic().run())}
      {btn('A̲', 'Altı çizili (Ctrl+U)', editor.isActive('underline'), () => chain().toggleUnderline().run())}
      {btn('S̶', 'Üstü çizili', editor.isActive('strike'), () => chain().toggleStrike().run())}
      <label className="tool color" title="Yazı rengi">
        A
        <input
          type="color"
          value={(editor.getAttributes('textStyle').color as string | undefined) ?? '#000000'}
          onChange={(e) => chain().setColor(e.target.value).run()}
        />
      </label>
      <span className="sep" />
      {btn('⯇', 'Sola hizala', editor.isActive({ textAlign: 'left' }), () => chain().setTextAlign('left').run())}
      {btn('≡', 'Ortala', editor.isActive({ textAlign: 'center' }), () => chain().setTextAlign('center').run())}
      {btn('⯈', 'Sağa hizala', editor.isActive({ textAlign: 'right' }), () => chain().setTextAlign('right').run())}
      {btn('☰', 'İki yana yasla', editor.isActive({ textAlign: 'justify' }), () => chain().setTextAlign('justify').run())}
      <span className="sep" />
      {btn('•', 'Madde işaretli liste', editor.isActive('bulletList'), () => chain().toggleBulletList().run())}
      {btn('1.', 'Numaralı liste', editor.isActive('orderedList'), () => chain().toggleOrderedList().run())}
      <span className="sep" />
      {btn('⊞', 'Tablo ekle (3×3)', false, () => chain().insertTable({ rows: 3, cols: 3, withHeaderRow: true }).run())}
      {editor.isActive('table') && (
        <>
          {btn('+Satır', 'Alta satır ekle', false, () => chain().addRowAfter().run())}
          {btn('+Sütun', 'Sağa sütun ekle', false, () => chain().addColumnAfter().run())}
          {btn('−Satır', 'Satırı sil', false, () => chain().deleteRow().run())}
          {btn('−Sütun', 'Sütunu sil', false, () => chain().deleteColumn().run())}
          {btn('✕Tablo', 'Tabloyu sil', false, () => chain().deleteTable().run())}
        </>
      )}
    </div>
  );
}
