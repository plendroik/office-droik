export type DocKind = 'pdf' | 'sheet' | 'doc' | 'slides';

interface FormatInfo {
  kind: DocKind;
  /** Eski/yabancı biçimler açılmadan önce LibreOffice ile bu biçime çevrilir. */
  convertTo?: 'docx' | 'xlsx' | 'pptx';
}

const FORMATS: Record<string, FormatInfo> = {
  pdf: { kind: 'pdf' },
  xlsx: { kind: 'sheet' },
  xls: { kind: 'sheet', convertTo: 'xlsx' },
  ods: { kind: 'sheet', convertTo: 'xlsx' },
  csv: { kind: 'sheet', convertTo: 'xlsx' },
  docx: { kind: 'doc' },
  doc: { kind: 'doc', convertTo: 'docx' },
  odt: { kind: 'doc', convertTo: 'docx' },
  rtf: { kind: 'doc', convertTo: 'docx' },
  pptx: { kind: 'slides' },
  ppt: { kind: 'slides', convertTo: 'pptx' },
  odp: { kind: 'slides', convertTo: 'pptx' },
};

export const ACCEPT = Object.keys(FORMATS).map((ext) => `.${ext}`).join(',');

export const MIME: Record<string, string> = {
  pdf: 'application/pdf',
  xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
};

export interface OpenedFile {
  id: string;
  name: string;
  kind: DocKind;
  data: ArrayBuffer;
  /** Dosya File System Access API ile açıldıysa yerinde kaydetmek için kullanılır. */
  handle?: FileSystemFileHandle;
  /** Dosya açılırken başka biçimden çevrildiyse kaynak adı. */
  convertedFrom?: string;
}

export function extOf(name: string): string {
  const dot = name.lastIndexOf('.');
  return dot === -1 ? '' : name.slice(dot + 1).toLowerCase();
}

export function withExt(name: string, ext: string): string {
  const dot = name.lastIndexOf('.');
  return `${dot === -1 ? name : name.slice(0, dot)}.${ext}`;
}

export function isSupported(name: string): boolean {
  return extOf(name) in FORMATS;
}

export async function serverStatus(): Promise<{ soffice: boolean }> {
  try {
    const res = await fetch('/api/status');
    return await res.json();
  } catch {
    return { soffice: false };
  }
}

export async function convert(data: BlobPart, name: string, to: 'pdf' | 'docx' | 'xlsx' | 'pptx'): Promise<ArrayBuffer> {
  const res = await fetch(`/api/convert?to=${to}&name=${encodeURIComponent(name)}`, {
    method: 'POST',
    body: new Blob([data]),
  });
  if (!res.ok) {
    const body = await res.json().catch(() => ({ error: res.statusText }));
    throw new Error(body.error ?? 'Dönüştürme başarısız');
  }
  return res.arrayBuffer();
}

export async function openFile(file: File, handle?: FileSystemFileHandle): Promise<OpenedFile> {
  const info = FORMATS[extOf(file.name)];
  if (!info) throw new Error(`Desteklenmeyen dosya türü: ${file.name}`);
  const id = crypto.randomUUID();
  const data = await file.arrayBuffer();
  if (!info.convertTo) return { id, name: file.name, kind: info.kind, data, handle };
  // Çevrilen dosyayı asıl dosyanın üzerine yazmamak için handle bırakılır.
  return {
    id,
    name: withExt(file.name, info.convertTo),
    kind: info.kind,
    data: await convert(data, file.name, info.convertTo),
    convertedFrom: file.name,
  };
}

type PickerWindow = Window & {
  showOpenFilePicker?: (opts?: object) => Promise<FileSystemFileHandle[]>;
  showSaveFilePicker?: (opts?: object) => Promise<FileSystemFileHandle>;
};
const pickerWindow = window as PickerWindow;

export const hasFsAccess = typeof pickerWindow.showOpenFilePicker === 'function';

export async function pickFiles(): Promise<{ file: File; handle: FileSystemFileHandle }[]> {
  const handles = await pickerWindow.showOpenFilePicker!({
    multiple: true,
    types: [{ description: 'Belgeler', accept: { 'application/octet-stream': ACCEPT.split(',') } }],
  });
  return Promise.all(handles.map(async (handle) => ({ file: await handle.getFile(), handle })));
}

/**
 * Handle varsa yerinde yazar; yoksa (veya saveAs ise) kaydetme penceresi açar,
 * o da yoksa indirme olarak kaydeder. Yeni handle'ı döndürür.
 */
export async function saveBlob(
  blob: Blob,
  name: string,
  handle: FileSystemFileHandle | undefined,
  saveAs: boolean,
): Promise<{ name: string; handle?: FileSystemFileHandle } | null> {
  let target = saveAs ? undefined : handle;
  if (!target && pickerWindow.showSaveFilePicker) {
    try {
      target = await pickerWindow.showSaveFilePicker({ suggestedName: name });
    } catch (err) {
      if ((err as DOMException).name === 'AbortError') return null;
      throw err;
    }
  }
  if (target) {
    const writable = await target.createWritable();
    await writable.write(blob);
    await writable.close();
    return { name: target.name, handle: target };
  }
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = name;
  a.click();
  setTimeout(() => URL.revokeObjectURL(url), 10_000);
  return { name };
}
