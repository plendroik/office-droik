// ExcelJS çalışma kitabı <-> Univer anlık görüntüsü (IWorkbookData) dönüşümleri.
import ExcelJS from 'exceljs';
import type { ICellData, IColumnData, IRowData, IStyleData, IWorkbookData, IWorksheetData, Nullable } from '@univerjs/presets';
import type { FWorkbook } from '@univerjs/preset-sheets-core';

// Excel'in varsayılan Office teması; tema rengi indeksleri bu sıradadır (lt1, dk1, lt2, dk2, accent1-6).
const THEME_COLORS = ['FFFFFF', '000000', 'E7E6E6', '44546A', '4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47'];

const BORDER_STYLES = [
  'thin', 'hair', 'dotted', 'dashed', 'dashDot', 'dashDotDot', 'double',
  'medium', 'mediumDashed', 'mediumDashDot', 'mediumDashDotDot', 'slantDashDot', 'thick',
] as const;

const H_ALIGN: Record<string, number> = { left: 1, center: 2, centerContinuous: 2, right: 3, justify: 4, distributed: 6 };
const V_ALIGN: Record<string, number> = { top: 1, middle: 2, bottom: 3 };
const WRAP = 3;

const EXCEL_EPOCH = Date.UTC(1899, 11, 30);
const DAY_MS = 86_400_000;

// Excel sütun genişliği karakter cinsindendir; Univer piksel kullanır.
const charsToPx = (w: number) => Math.round(w * 7 + 5);
const pxToChars = (px: number) => Math.max(0, (px - 5) / 7);
const ptToPx = (pt: number) => Math.round((pt * 4) / 3);
const pxToPt = (px: number) => (px * 3) / 4;

function applyTint(hex: string, tint: number): string {
  const channels = [0, 2, 4].map((i) => parseInt(hex.slice(i, i + 2), 16));
  const tinted = channels.map((c) => Math.round(tint < 0 ? c * (1 + tint) : c + (255 - c) * tint));
  return tinted.map((c) => c.toString(16).padStart(2, '0')).join('');
}

function toRgb(color: Partial<ExcelJS.Color> | undefined): string | undefined {
  if (!color) return undefined;
  if (color.argb) return `#${color.argb.length === 8 ? color.argb.slice(2) : color.argb}`;
  if (color.theme !== undefined && THEME_COLORS[color.theme]) {
    const tint = (color as { tint?: number }).tint ?? 0;
    return `#${applyTint(THEME_COLORS[color.theme], tint)}`;
  }
  return undefined;
}

const toArgb = (rgb: Nullable<string>) => (rgb ? `FF${rgb.replace('#', '').toUpperCase()}` : undefined);

function excelStyleToUniver(style: Partial<ExcelJS.Style>): IStyleData | null {
  const s: IStyleData = {};
  const { font, fill, alignment, border, numFmt } = style;
  if (font) {
    if (font.bold) s.bl = 1;
    if (font.italic) s.it = 1;
    if (font.underline) s.ul = { s: 1 };
    if (font.strike) s.st = { s: 1 };
    if (font.size) s.fs = font.size;
    if (font.name) s.ff = font.name;
    const cl = toRgb(font.color);
    if (cl && cl !== '#000000') s.cl = { rgb: cl };
  }
  if (fill?.type === 'pattern' && fill.pattern === 'solid') {
    const bg = toRgb(fill.fgColor);
    if (bg) s.bg = { rgb: bg };
  }
  if (alignment) {
    if (alignment.horizontal && H_ALIGN[alignment.horizontal]) s.ht = H_ALIGN[alignment.horizontal];
    if (alignment.vertical && V_ALIGN[alignment.vertical]) s.vt = V_ALIGN[alignment.vertical];
    if (alignment.wrapText) s.tb = WRAP;
  }
  if (border) {
    const bd: NonNullable<IStyleData['bd']> = {};
    for (const [side, key] of [['top', 't'], ['bottom', 'b'], ['left', 'l'], ['right', 'r']] as const) {
      const b = border[side];
      const idx = b?.style ? BORDER_STYLES.indexOf(b.style as (typeof BORDER_STYLES)[number]) : -1;
      if (idx >= 0) bd[key] = { s: idx + 1, cl: { rgb: toRgb(b!.color) ?? '#000000' } };
    }
    if (Object.keys(bd).length) s.bd = bd;
  }
  if (numFmt && numFmt !== 'General') s.n = { pattern: numFmt };
  return Object.keys(s).length ? s : null;
}

function univerStyleToExcel(s: IStyleData): Partial<ExcelJS.Style> {
  const style: Partial<ExcelJS.Style> = {};
  const font: Partial<ExcelJS.Font> = {};
  if (s.bl) font.bold = true;
  if (s.it) font.italic = true;
  if (s.ul?.s) font.underline = true;
  if (s.st?.s) font.strike = true;
  if (s.fs) font.size = s.fs;
  if (s.ff) font.name = s.ff;
  if (s.cl?.rgb) font.color = { argb: toArgb(s.cl.rgb) };
  if (Object.keys(font).length) style.font = font;
  if (s.bg?.rgb) style.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: toArgb(s.bg.rgb) } };

  const alignment: Partial<ExcelJS.Alignment> = {};
  const h = Object.entries(H_ALIGN).find(([name, v]) => v === s.ht && name !== 'centerContinuous');
  if (h) alignment.horizontal = h[0] as ExcelJS.Alignment['horizontal'];
  const v = Object.entries(V_ALIGN).find(([, n]) => n === s.vt);
  if (v) alignment.vertical = v[0] as ExcelJS.Alignment['vertical'];
  if (s.tb === WRAP) alignment.wrapText = true;
  if (Object.keys(alignment).length) style.alignment = alignment;

  if (s.bd) {
    const border: Partial<ExcelJS.Borders> = {};
    for (const [key, side] of [['t', 'top'], ['b', 'bottom'], ['l', 'left'], ['r', 'right']] as const) {
      const b = s.bd[key];
      if (b && b.s > 0) border[side] = { style: BORDER_STYLES[b.s - 1], color: { argb: toArgb(b.cl?.rgb) ?? 'FF000000' } };
    }
    if (Object.keys(border).length) style.border = border;
  }
  if (s.n?.pattern) style.numFmt = s.n.pattern;
  return style;
}

function cellContent(cell: ExcelJS.Cell): Pick<ICellData, 'v' | 't' | 'f'> | null {
  const value = cell.value;
  if (value === null || value === undefined) return null;
  if (typeof value === 'number') return { v: value, t: 2 };
  if (typeof value === 'string') return { v: value, t: 1 };
  if (typeof value === 'boolean') return { v: value, t: 3 };
  if (value instanceof Date) return { v: (value.getTime() - EXCEL_EPOCH) / DAY_MS, t: 2 };
  if ('formula' in value || 'sharedFormula' in value) {
    // ExcelJS paylaşılan formülleri bu hücreye göre çevrilmiş olarak döndürür.
    const result = (value as ExcelJS.CellFormulaValue).result;
    const v = result instanceof Date ? (result.getTime() - EXCEL_EPOCH) / DAY_MS
      : typeof result === 'object' && result && 'error' in result ? result.error
      : result;
    return { f: `=${cell.formula}`, v: v as ICellData['v'] };
  }
  if ('richText' in value) return { v: value.richText.map((r) => r.text).join(''), t: 1 };
  if ('hyperlink' in value) return { v: String(value.text ?? value.hyperlink), t: 1 };
  if ('error' in value) return { v: value.error, t: 1 };
  return { v: String(cell.text), t: 1 };
}

function parseRef(ref: string): { row: number; col: number } {
  const m = /^\$?([A-Z]+)\$?(\d+)$/.exec(ref.toUpperCase())!;
  const col = [...m[1]].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
  return { row: Number(m[2]) - 1, col: col - 1 };
}

export async function xlsxToUniver(data: ArrayBuffer, name: string): Promise<Partial<IWorkbookData>> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(data);

  const styles: Record<string, IStyleData> = {};
  const styleIds = new Map<string, string>();
  const styleId = (s: IStyleData) => {
    const key = JSON.stringify(s);
    let id = styleIds.get(key);
    if (!id) {
      id = `s${styleIds.size}`;
      styleIds.set(key, id);
      styles[id] = s;
    }
    return id;
  };

  const sheets: Record<string, Partial<IWorksheetData>> = {};
  const sheetOrder: string[] = [];

  wb.eachSheet((ws, index) => {
    const id = `sheet-${index}`;
    sheetOrder.push(id);
    const cellData: Record<number, Record<number, ICellData>> = {};
    const rowData: Record<number, { h: number; hd?: 0 | 1 }> = {};
    const columnData: Record<number, { w: number; hd?: 0 | 1 }> = {};

    ws.eachRow({ includeEmpty: true }, (row, r) => {
      if (row.height || row.hidden) rowData[r - 1] = { h: ptToPx(row.height ?? 15), ...(row.hidden ? { hd: 1 as const } : {}) };
      row.eachCell({ includeEmpty: true }, (cell, c) => {
        if (cell.isMerged && cell.master !== cell) return;
        const content = cellContent(cell);
        const style = excelStyleToUniver(cell.style);
        if (!content && !style) return;
        (cellData[r - 1] ??= {})[c - 1] = { ...content, ...(style ? { s: styleId(style) } : {}) };
      });
    });

    for (let c = 1; c <= ws.columnCount; c++) {
      const col = ws.getColumn(c);
      if (col.width || col.hidden) columnData[c - 1] = { w: charsToPx(col.width ?? 8.43), ...(col.hidden ? { hd: 1 as const } : {}) };
    }

    const mergeData = (ws.model.merges ?? []).map((range: string) => {
      const [a, b] = range.split(':').map(parseRef);
      return { startRow: a.row, startColumn: a.col, endRow: b.row, endColumn: b.col };
    });

    const view = ws.views?.[0] as Partial<ExcelJS.WorksheetViewFrozen> | undefined;
    const xSplit = view?.state === 'frozen' ? view.xSplit ?? 0 : 0;
    const ySplit = view?.state === 'frozen' ? view.ySplit ?? 0 : 0;

    sheets[id] = {
      id,
      name: ws.name,
      tabColor: toRgb(ws.properties.tabColor) ?? '',
      hidden: ws.state === 'visible' ? 0 : 1,
      rowCount: Math.max(1000, ws.rowCount + 100),
      columnCount: Math.max(26, ws.columnCount + 10),
      cellData,
      rowData,
      columnData,
      mergeData,
      freeze: { xSplit, ySplit, startRow: ySplit, startColumn: xSplit },
    };
  });

  if (sheetOrder.length === 0) {
    sheetOrder.push('sheet-1');
    sheets['sheet-1'] = { id: 'sheet-1', name: 'Sayfa1' };
  }

  return { id: crypto.randomUUID(), name, sheetOrder, sheets, styles };
}

export async function univerToXlsx(fWorkbook: FWorkbook): Promise<Blob> {
  const snapshot = fWorkbook.save();
  const wb = new ExcelJS.Workbook();
  const resolveStyle = (s: ICellData['s']): IStyleData | null =>
    typeof s === 'string' ? snapshot.styles[s] ?? null : s ?? null;

  const facadeSheets = new Map(fWorkbook.getSheets().map((s) => [s.getSheetId(), s]));

  for (const sheetId of snapshot.sheetOrder) {
    const sheet = snapshot.sheets[sheetId];
    const fSheet = facadeSheets.get(sheetId);
    if (!sheet || !fSheet) continue;
    const ws = wb.addWorksheet(sheet.name, {
      state: sheet.hidden ? 'hidden' : 'visible',
      properties: sheet.tabColor ? { tabColor: { argb: toArgb(sheet.tabColor) } } : {},
    });

    // Formülleri facade'dan alıyoruz: anlık görüntüde paylaşılan formüller yalnızca kimlik (si)
    // olarak durabilir. Değerler ise anlık görüntüden gelir; facade'ın getValues()'u tarihleri
    // biçimlendirilmiş metin olarak döndürür.
    const lastRow = fSheet.getLastRow();
    const lastCol = fSheet.getLastColumn();
    const range = lastRow >= 0 && lastCol >= 0 ? fSheet.getRange(0, 0, lastRow + 1, lastCol + 1) : null;
    const formulas = range?.getFormulas() ?? [];

    for (const [r, row] of Object.entries(sheet.cellData ?? {})) {
      for (const [c, cellData] of Object.entries(row ?? {}) as [string, ICellData | undefined][]) {
        const ri = Number(r);
        const ci = Number(c);
        const cell = ws.getCell(ri + 1, ci + 1);
        const formula = formulas[ri]?.[ci];
        const value = cellData?.v ?? null;
        const plain = cellData?.t === 3 ? Boolean(value) : value;
        if (formula) cell.value = { formula: formula.replace(/^=/, ''), result: plain ?? undefined } as ExcelJS.CellFormulaValue;
        else if (plain !== null && plain !== '') cell.value = plain;
        const style = resolveStyle(cellData?.s);
        if (style) cell.style = univerStyleToExcel(style);
      }
    }

    for (const m of sheet.mergeData ?? []) {
      ws.mergeCells(m.startRow + 1, m.startColumn + 1, m.endRow + 1, m.endColumn + 1);
    }
    for (const [c, col] of Object.entries(sheet.columnData ?? {}) as [string, Partial<IColumnData>][]) {
      if (!col) continue;
      const column = ws.getColumn(Number(c) + 1);
      if (col.w) column.width = pxToChars(col.w);
      if (col.hd) column.hidden = true;
    }
    for (const [r, row] of Object.entries(sheet.rowData ?? {}) as [string, Partial<IRowData>][]) {
      if (!row) continue;
      const wsRow = ws.getRow(Number(r) + 1);
      if (row.h && row.h !== sheet.defaultRowHeight) wsRow.height = pxToPt(row.h);
      if (row.hd) wsRow.hidden = true;
    }
    const { xSplit = 0, ySplit = 0 } = sheet.freeze ?? {};
    if (xSplit > 0 || ySplit > 0) ws.views = [{ state: 'frozen', xSplit, ySplit }];
  }

  if (wb.worksheets.length === 0) wb.addWorksheet('Sayfa1');
  const buffer = await wb.xlsx.writeBuffer();
  return new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}
