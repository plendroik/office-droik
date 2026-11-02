import type { Ref } from 'react';
import type { OpenedFile } from '../lib/files';

export interface EditorHandle {
  /** Belgenin güncel hâlini kendi biçiminde (pdf/xlsx/docx/pptx) üretir. */
  save(): Promise<Blob>;
}

export interface EditorProps {
  file: OpenedFile;
  sofficeAvailable: boolean;
  onDirty(): void;
  ref?: Ref<EditorHandle>;
}
