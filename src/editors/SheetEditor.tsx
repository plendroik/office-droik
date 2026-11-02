import { useEffect, useImperativeHandle, useRef, useState } from 'react';
import { createUniver, LocaleType, mergeLocales, type FUniver } from '@univerjs/presets';
import { UniverSheetsCorePreset } from '@univerjs/preset-sheets-core';
import sheetsCoreEnUS from '@univerjs/preset-sheets-core/locales/en-US';
import '@univerjs/preset-sheets-core/lib/index.css';
import { univerToXlsx, xlsxToUniver } from '../lib/xlsx';
import type { EditorProps } from './types';

// CommandType.COMMAND; kullanıcı eylemleri bu türdendir, sonuçları ise MUTATION olarak gelir.
const COMMAND = 0;

export function SheetEditor({ file, onDirty, ref }: EditorProps) {
  const containerRef = useRef<HTMLDivElement>(null);
  const apiRef = useRef<FUniver | null>(null);
  const onDirtyRef = useRef(onDirty);
  onDirtyRef.current = onDirty;
  const [error, setError] = useState<string | null>(null);

  useImperativeHandle(ref, () => ({
    async save() {
      const workbook = apiRef.current?.getActiveWorkbook();
      if (!workbook) throw new Error('Çalışma kitabı hazır değil');
      return univerToXlsx(workbook);
    },
  }));

  useEffect(() => {
    let disposed = false;
    const { univer, univerAPI } = createUniver({
      locale: LocaleType.EN_US,
      locales: { [LocaleType.EN_US]: mergeLocales(sheetsCoreEnUS) },
      presets: [UniverSheetsCorePreset({ container: containerRef.current! })],
    });
    apiRef.current = univerAPI;

    xlsxToUniver(file.data, file.name)
      .then((snapshot) => {
        if (disposed) return;
        univerAPI.createWorkbook(snapshot);
        // Yükleme sırasında formül hesaplama ve hücre editörü kendi mutation'larını üretir;
        // bu yüzden yalnızca kullanıcı kaynaklı tablo komutları belgeyi "kaydedilmedi" yapar.
        univerAPI.addEvent(univerAPI.Event.CommandExecuted, (event) => {
          if (event.type === COMMAND && event.id.startsWith('sheet.command.')) onDirtyRef.current();
        });
      })
      .catch((err: Error) => !disposed && setError(err.message));

    return () => {
      disposed = true;
      apiRef.current = null;
      univer.dispose();
    };
    // Ad, "farklı kaydet" sonrası değişebilir; yeniden yükleme yalnızca veri değişince.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [file.data]);

  return (
    <div className="editor-fill">
      {error && <div className="editor-error">Dosya açılamadı: {error}</div>}
      <div ref={containerRef} className="univer-host" />
    </div>
  );
}
