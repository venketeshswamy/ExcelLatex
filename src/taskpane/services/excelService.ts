/**
 * ExcelKaTeX Excel Service
 *
 * Button behaviour (per user requirements):
 *   Insert Formula  → writes =MATH.KATEX(...) formula string into the active cell only.
 *                     No image is inserted. The formula drives the cell value.
 *   In-Cell Image   → tries modern Excel EntityCellValue (ExcelApi 1.16+) first;
 *                     falls back to floating shape if not supported.
 *   Floating Shape  → always inserts a floating PNG image shape anchored to the cell.
 *   Read Cell       → reads LaTeX from the active cell (formula, entity, or 📐 marker).
 */

import { compileLatex, RenderOptions } from '../../core/katexEngine';
import { buildKatexEntityCellValue, isEntityCellValueSupported } from '../../customfunctions/entityCellBuilder';

export { buildKatexEntityCellValue, isEntityCellValueSupported };

export interface ExcelService {
  insertFormulaToActiveCell(latex: string, options?: RenderOptions): Promise<void>;
  insertInCellImageToActiveCell(latex: string, options?: RenderOptions): Promise<void>;
  insertFloatingShapeToActiveCell(latex: string, options?: RenderOptions): Promise<void>;
  readActiveCellFormula(): Promise<string | null>;
}

/** Escapes double-quotes for use inside an Excel formula string literal. */
export function escapeFormulaString(str: string): string {
  return str.replace(/"/g, '""');
}

/**
 * Strips the "data:image/png;base64," header prefix from a data URL.
 * Excel's shapes.addImage() requires the raw Base64 bytes only.
 */
function toRawBase64(dataUrl: string): string {
  if (!dataUrl) return '';
  const idx = dataUrl.indexOf(',');
  return idx >= 0 ? dataUrl.slice(idx + 1) : dataUrl;
}

/** Formats the background parameter for Excel formula insertion. */
export function formatBackgroundParam(bg: any): string {
  if (bg === 0 || bg === '0') return '0';
  if (bg === 1 || bg === '1') return '1';
  if (bg === 2 || bg === '2') return '2';
  if (typeof bg === 'number') return String(bg);
  if (typeof bg === 'string') {
    const trimmed = bg.trim();
    const lower = trimmed.toLowerCase();
    if (lower === 'transparent') return '0';
    if (lower === 'white') return '1';
    if (lower === 'black') return '2';
    return `"${escapeFormulaString(trimmed)}"`;
  }
  return '0';
}

// ─── Button 1: Insert Formula ────────────────────────────────────────────────

/**
 * Writes the =MATH.KATEX(...) formula string into the active cell.
 * Does NOT insert any image — the formula result is displayed by Excel's engine.
 */
export async function insertFormulaToActiveCell(
  latex: string,
  options: RenderOptions = {}
): Promise<void> {
  if (typeof Excel === 'undefined') {
    console.warn('[ExcelKaTeX] Excel runtime not found.');
    return;
  }

  const escapedLatex = escapeFormulaString(latex);
  const isDefaultBg =
    options.background === undefined ||
    options.background === 'transparent' ||
    options.background === '0' ||
    options.background === 0;

  const isDefaultColor =
    options.color === undefined ||
    options.color === '#000000';

  const isDefaultFontSize =
    options.fontSize === undefined ||
    options.fontSize === 16;

  const isDefaultDisplayMode =
    options.displayMode === undefined ||
    options.displayMode === true;

  const isCustomOptions = !(
    isDefaultBg &&
    isDefaultColor &&
    isDefaultFontSize &&
    isDefaultDisplayMode
  );

  let formula = `=MATH.KATEX("${escapedLatex}")`;
  if (isCustomOptions) {
    const bg = formatBackgroundParam(options.background);
    const defaultText = (options.background === 2 || options.background === '2' || options.background === 'black' || options.background === '#000000')
      ? '#ffffff'
      : '#000000';
    const fg = `"${escapeFormulaString(options.color || defaultText)}"`;
    const sz = options.fontSize ?? 16;
    const dm = options.displayMode ?? true;
    formula = `=MATH.KATEX("${escapedLatex}", ${bg}, ${fg}, ${sz}, ${dm})`;
  }

  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.formulas = [[formula]];
    await context.sync();
  });
}

// ─── Button 2: In-Cell Image ─────────────────────────────────────────────────

/**
 * Attempts to insert an EntityCellValue (modern in-cell image, ExcelApi 1.16+).
 * Falls back to a floating shape if the current Excel build doesn't support it.
 */
export async function insertInCellImageToActiveCell(
  latex: string,
  options: RenderOptions = {}
): Promise<void> {
  if (typeof Excel === 'undefined') {
    console.warn('[ExcelKaTeX] Excel runtime not found.');
    return;
  }

  const supportsEntity = isEntityCellValueSupported();

  if (supportsEntity) {
    try {
      const renderResult = await compileLatex(latex, options);
      const entity = buildKatexEntityCellValue(latex, renderResult, options);

      await Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        // Set the entity value — both valuesAsJson and values for mock and live runtime compatibility
        (range as any).valuesAsJson = [[entity]];
        range.values = [[entity as any]];
        await context.sync();
      });
      return;
    } catch (entityErr) {
      console.warn('[ExcelKaTeX] EntityCellValue failed, falling back to floating shape:', entityErr);
    }
  }

  // Fallback: floating shape
  await insertFloatingShapeToActiveCell(latex, options);
}

// ─── Button 3: Floating Shape ─────────────────────────────────────────────────

/**
 * Renders the LaTeX formula to a high-DPI PNG and inserts it as a
 * floating image shape anchored to the top-left corner of the selected cell.
 * Shape dimensions are natural (no artificial max-width cap).
 */
export async function insertFloatingShapeToActiveCell(
  latex: string,
  options: RenderOptions = {}
): Promise<void> {
  if (typeof Excel === 'undefined') {
    console.warn('[ExcelKaTeX] Excel runtime not found.');
    return;
  }

  const renderResult = await compileLatex(latex, options);
  const rawBase64 = toRawBase64(renderResult.pngDataUrl);

  if (!rawBase64 || rawBase64.length < 10) {
    throw new Error(
      'Image generation failed — the rasterizer returned an empty result. ' +
      'Please check the browser console for details.'
    );
  }

  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = context.workbook.getSelectedRange();
    range.load(['left', 'top']);
    await context.sync();

    const shape = sheet.shapes.addImage(rawBase64);
    shape.left = range.left;
    shape.top  = range.top;
    // Use natural dimensions at 1× (the PNG is already rendered at 3× scale internally)
    shape.width  = renderResult.width;
    shape.height = renderResult.height;
    shape.altTextTitle       = `LaTeX: ${latex}`;
    shape.altTextDescription = latex;
    await context.sync();
  });
}

// ─── Button 4: Read Cell ──────────────────────────────────────────────────────

/**
 * Reads the LaTeX formula from the currently selected cell.
 * Handles: =MATH.KATEX(...) formula, EntityCellValue, and 📐 marker strings.
 */
export async function readActiveCellFormula(): Promise<string | null> {
  if (typeof Excel === 'undefined') {
    console.warn('[ExcelKaTeX] Excel runtime not found.');
    return null;
  }

  return await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load(['formulas', 'values']);
    await context.sync();

    // Check formula bar first
    const formula = range.formulas?.[0]?.[0];
    if (formula && typeof formula === 'string' && formula.startsWith('=')) {
      const match = formula.match(/^=MATH\.KATEX\s*\(\s*"((?:[^"]|"")*)"/i);
      if (match?.[1]) return match[1].replace(/""/g, '"');
      return formula;
    }

    const value = range.values?.[0]?.[0];

    // 📐 marker from custom function
    if (typeof value === 'string' && value.startsWith('📐 ')) {
      return value.slice(3).trim();
    }

    // EntityCellValue
    if (value && typeof value === 'object') {
      const latex = (value as any).properties?.latex?.basicValue;
      if (latex) return String(latex);
    }

    // Plain text
    if (value !== undefined && value !== null && value !== '') {
      return String(value);
    }

    return null;
  });
}

// ─── SharedRuntime onChanged Auto-Renderer ────────────────────────────────────

/**
 * Registers a worksheet.onChanged handler that automatically converts
 * "📐 <latex>" cells (produced by =MATH.KATEX custom function) to floating images.
 * Call this once from Office.onReady().
 */
export async function installKatexShapeAutoRenderer(
  defaultOptions: RenderOptions = {}
): Promise<void> {
  if (typeof Excel === 'undefined') return;

  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.onChanged.add(async (event) => {
        if (!event.address) return;
        try {
          await Excel.run(async (ctx2) => {
            const changedRange = ctx2.workbook.worksheets
              .getActiveWorksheet()
              .getRange(event.address);
            changedRange.load(['values', 'left', 'top']);
            await ctx2.sync();

            const cellVal = changedRange.values?.[0]?.[0];
            if (typeof cellVal !== 'string' || !cellVal.startsWith('📐 ')) return;

            const latex = cellVal.slice(3).trim();
            if (!latex) return;

            const renderResult = await compileLatex(latex, defaultOptions);
            const rawBase64 = toRawBase64(renderResult.pngDataUrl);
            if (!rawBase64 || rawBase64.length < 10) return;

            const ws2 = ctx2.workbook.worksheets.getActiveWorksheet();
            changedRange.load(['left', 'top']);
            await ctx2.sync();

            const shape = ws2.shapes.addImage(rawBase64);
            shape.left = changedRange.left;
            shape.top  = changedRange.top;
            shape.width  = renderResult.width;
            shape.height = renderResult.height;
            shape.altTextTitle = `LaTeX: ${latex}`;
            await ctx2.sync();
          });
        } catch (err) {
          console.error('[ExcelKaTeX] Auto-render error:', err);
        }
      });
      await context.sync();
    });
  } catch (err) {
    console.error('[ExcelKaTeX] installKatexShapeAutoRenderer failed:', err);
  }
}

export const excelService: ExcelService = {
  insertFormulaToActiveCell,
  insertInCellImageToActiveCell,
  insertFloatingShapeToActiveCell,
  readActiveCellFormula,
};
