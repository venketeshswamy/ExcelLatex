/**
 * ExcelKaTeX Excel Service — Production-Ready
 *
 * Core Capabilities:
 *   1. Insert In-Cell Image  → Native Excel in-cell picture (WebImage via valuesAsJson) with metadata in AltText.
 *   2. Insert Formula        → Clean, simplified =MATH.KATEX("...", [bg]) formula.
 *   3. Floating Shape        → In-place shape replacement anchored to cell with zero duplicate stacking.
 *   4. Read Cell (Bi-dir)    → Reads active cell (formula, WebImage alt-text JSON, Entity, or floating shape).
 *   5. Batch Convert Range   → Converts multi-cell LaTeX selections to images in batch.
 *   6. Delete Shapes         → Cleans floating shapes in selected range.
 */

import { compileLatex, RenderOptions } from '../../core/katexEngine';
import {
  buildKatexEntityCellValue,
  buildKatexWebImageCellValue,
  isEntityCellValueSupported
} from '../../customfunctions/entityCellBuilder';
import {
  serializeEquationMetadata,
  parseEquationMetadata
} from '../../core/imageRasterizer';

export { buildKatexEntityCellValue, buildKatexWebImageCellValue, isEntityCellValueSupported };

export interface ReadCellResult {
  latex: string;
  background?: string | number;
  color?: string;
  fontSize?: number;
  displayMode?: boolean;
}

export interface ExcelService {
  insertFormulaToActiveCell(latex: string, options?: RenderOptions): Promise<void>;
  insertInCellImageToActiveCell(latex: string, options?: RenderOptions): Promise<void>;
  insertFloatingShapeToActiveCell(latex: string, options?: RenderOptions): Promise<void>;
  readActiveCellFormula(): Promise<ReadCellResult | string | null>;
  batchConvertSelectedRange(outputType?: 'image' | 'formula' | 'shape', options?: RenderOptions): Promise<{ total: number; converted: number }>;
  deleteShapesInSelection(): Promise<number>;
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

/** Formats the background parameter for simplified Excel formula insertion. */
export function formatBackgroundParam(bg: any): string {
  if (bg === 0 || bg === '0') return '0';
  if (bg === 1 || bg === '1') return '1';
  if (bg === 2 || bg === '2') return '2';
  if (typeof bg === 'number') return String(bg);
  if (typeof bg === 'string') {
    const trimmed = bg.trim();
    const lower = trimmed.toLowerCase();
    if (lower === '0' || lower === 'transparent') return '0';
    if (lower === '1' || lower === 'white') return '1';
    if (lower === '2' || lower === 'black') return '2';
    return `"${escapeFormulaString(trimmed)}"`;
  }
  return '0';
}

// ─── Action 1: In-Cell Image (Primary 1-Click Action) ────────────────────────

/**
 * Inserts a native in-cell picture into the currently active cell.
 * The image sits directly in the grid cell, with full metadata embedded in AltText.
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
      const webImage = buildKatexWebImageCellValue(latex, renderResult, options);

      await Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        (range as any).valuesAsJson = [[webImage]];
        await context.sync();
      });
      return;
    } catch (err) {
      console.warn('[ExcelKaTeX] In-cell image failed, falling back to shape:', err);
    }
  }

  // Fallback to floating shape on legacy hosts
  await insertFloatingShapeToActiveCell(latex, options);
}

// ─── Action 2: Insert Formula (=MATH.KATEX) ──────────────────────────────────

/**
 * Writes the simplified =MATH.KATEX("...", [bg]) formula string into the active cell.
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
  const bg = formatBackgroundParam(options.background);

  // Simplified bulletproof formula: =MATH.KATEX(latex) or =MATH.KATEX(latex, bg)
  let formula = `=MATH.KATEX("${escapedLatex}")`;
  if (bg !== '0') {
    formula = `=MATH.KATEX("${escapedLatex}", ${bg})`;
  }

  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.formulas = [[formula]];
    await context.sync();
  });
}

// ─── Action 3: Floating Shape (With In-Place Anti-Stacking Replacement) ────────

/**
 * Renders the LaTeX formula to a 4x Retina PNG and inserts it as a floating
 * shape anchored to the cell. Replaces any existing shape on that cell.
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
    throw new Error('Image generation failed — empty rasterizer output.');
  }

  const metadataJson = serializeEquationMetadata(latex, options);

  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = context.workbook.getSelectedRange();
    range.load(['left', 'top', 'address']);
    const shapes = sheet.shapes;
    shapes.load(['items/name', 'items/left', 'items/top', 'items/altTextTitle']);
    await context.sync();

    const targetLeft = range.left;
    const targetTop = range.top;
    const shapeTag = `LaTeX_Shape_${range.address || 'cell'}`;

    // Remove any existing equation shape anchored at this cell to prevent duplicate stacking
    for (const s of shapes.items) {
      if (
        s.name === shapeTag ||
        (Math.abs(s.left - targetLeft) < 5 && Math.abs(s.top - targetTop) < 5 && s.altTextTitle?.startsWith('LaTeX:'))
      ) {
        try {
          s.delete();
        } catch { /* ignore */ }
      }
    }
    await context.sync();

    // Insert new high-resolution shape
    const shape = sheet.shapes.addImage(rawBase64);
    shape.name = shapeTag;
    shape.left = targetLeft;
    shape.top = targetTop;
    shape.width = renderResult.width;
    shape.height = renderResult.height;
    shape.altTextTitle = `LaTeX: ${latex}`;
    shape.altTextDescription = metadataJson;

    await context.sync();
  });
}

// ─── Action 4: Read Cell (Bidirectional Smart Inspection) ─────────────────────

/**
 * Reads formula or metadata from active cell or floating shape.
 */
export async function readActiveCellFormula(): Promise<ReadCellResult | string | null> {
  if (typeof Excel === 'undefined') {
    console.warn('[ExcelKaTeX] Excel runtime not found.');
    return null;
  }

  return await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = context.workbook.getSelectedRange();
    range.load(['valuesAsJson', 'formulas', 'values', 'left', 'top', 'address']);
    const shapes = sheet.shapes;
    shapes.load(['items/name', 'items/left', 'items/top', 'items/altTextTitle', 'items/altTextDescription']);
    await context.sync();

    // 1. Check formula bar (e.g. =MATH.KATEX(...) or in-cell image JSON @{"latex":...})
    const formula = range.formulas?.[0]?.[0];
    if (formula && typeof formula === 'string') {
      if (formula.startsWith('=')) {
        const match = formula.match(/^=MATH\.KATEX\s*\(\s*"((?:[^"]|"")*)"(?:\s*,\s*([^,)]+))?/i);
        if (match?.[1]) {
          const unescapedLatex = match[1].replace(/""/g, '"');
          const rawBg = match[2]?.trim();
          let bg: any = 0;
          if (rawBg === '1' || rawBg?.toLowerCase() === '"white"') bg = 1;
          if (rawBg === '2' || rawBg?.toLowerCase() === '"black"') bg = 2;
          return {
            latex: unescapedLatex,
            background: bg
          };
        }
      }

      // Check if formula is raw in-cell image JSON (e.g. @{"latex":...} or {"latex":...})
      const cleanJson = formula.replace(/^@+/, '').trim();
      if (cleanJson.startsWith('{') && cleanJson.includes('"latex"')) {
        const parsed = parseEquationMetadata(cleanJson);
        if (parsed) return parsed;
      }
    }

    // 2. Check in-cell rich values via valuesAsJson
    const richValue = (range as any).valuesAsJson?.[0]?.[0];
    if (richValue && typeof richValue === 'object') {
      const alt = richValue.altText;
      if (alt) {
        const parsed = parseEquationMetadata(alt);
        if (parsed) return parsed;
      }
      const entityLatex = richValue.properties?.latex?.basicValue;
      if (entityLatex) {
        return {
          latex: String(entityLatex),
          background: 0
        };
      }
    }

    // 3. Check legacy range.values object if present
    const value = range.values?.[0]?.[0];
    if (value && typeof value === 'object') {
      const alt = (value as any).altText;
      if (alt) {
        const parsed = parseEquationMetadata(alt);
        if (parsed) return parsed;
      }
    }

    // 4. Check floating shapes overlapping this cell
    const targetLeft = range.left;
    const targetTop = range.top;
    for (const s of shapes.items) {
      if (Math.abs(s.left - targetLeft) < 15 && Math.abs(s.top - targetTop) < 15) {
        const altDesc = s.altTextDescription;
        if (altDesc) {
          const parsed = parseEquationMetadata(altDesc);
          if (parsed) return parsed;
        }
        if (s.altTextTitle?.startsWith('LaTeX:')) {
          return {
            latex: s.altTextTitle.substring(6).trim(),
            background: 0
          };
        }
      }
    }

    // 5. Plain text in cell (ignoring "#VALUE!" error placeholder)
    if (typeof value === 'string' && value.trim() && value.trim() !== '#VALUE!') {
      const clean = value.startsWith('📐 ') ? value.slice(3).trim() : value.trim();
      if (clean.startsWith('{') && clean.includes('"latex"')) {
        const parsed = parseEquationMetadata(clean);
        if (parsed) return parsed;
      }
      return {
        latex: clean,
        background: 0
      };
    }

    return null;
  });
}

// ─── Action 5: Batch Convert Selected Range ───────────────────────────────────

/**
 * Scans all non-empty cells in selected range and converts them to rendered math in adjacent cells.
 */
export async function batchConvertSelectedRange(
  outputType: 'image' | 'formula' | 'shape' = 'image',
  options: RenderOptions = {}
): Promise<{ total: number; converted: number }> {
  if (typeof Excel === 'undefined') return { total: 0, converted: 0 };

  return await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load(['values', 'formulas', 'rowCount', 'columnCount', 'rowIndex', 'columnIndex']);
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    await context.sync();

    const rowCount = range.rowCount;
    const colCount = range.columnCount;
    let converted = 0;
    let total = 0;

    for (let r = 0; r < rowCount; r++) {
      for (let c = 0; c < colCount; c++) {
        const val = range.values[r][c];
        const rawFormula = range.formulas[r][c];
        let latex = '';

        if (typeof rawFormula === 'string' && rawFormula.startsWith('=')) {
          const m = rawFormula.match(/^=MATH\.KATEX\s*\(\s*"((?:[^"]|"")*)"/i);
          if (m?.[1]) latex = m[1].replace(/""/g, '"');
        } else if (typeof val === 'string' && val.trim() && val.trim() !== '#VALUE!') {
          latex = val.trim();
        }

        if (latex) {
          total++;
          try {
            const destCell = sheet.getCell(range.rowIndex + r, range.columnIndex + colCount + c);
            if (outputType === 'formula') {
              const bg = formatBackgroundParam(options.background);
              destCell.formulas = [[`=MATH.KATEX("${escapeFormulaString(latex)}", ${bg})`]];
            } else {
              const renderResult = await compileLatex(latex, options);
              const webImage = buildKatexWebImageCellValue(latex, renderResult, options);
              (destCell as any).valuesAsJson = [[webImage]];
            }
            converted++;
          } catch { /* continue */ }
        }
      }
    }

    await context.sync();
    return { total, converted };
  });
}

// ─── Action 6: Delete Shapes in Selection ────────────────────────────────────

/**
 * Deletes all floating math shapes overlapping the current cell selection.
 */
export async function deleteShapesInSelection(): Promise<number> {
  if (typeof Excel === 'undefined') return 0;

  return await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = context.workbook.getSelectedRange();
    range.load(['left', 'top', 'width', 'height']);
    const shapes = sheet.shapes;
    shapes.load(['items/name', 'items/left', 'items/top', 'items/altTextTitle']);
    await context.sync();

    const minX = range.left;
    const maxX = range.left + range.width;
    const minY = range.top;
    const maxY = range.top + range.height;

    let deleted = 0;
    for (const s of shapes.items) {
      if (
        s.left >= minX - 10 &&
        s.left <= maxX + 10 &&
        s.top >= minY - 10 &&
        s.top <= maxY + 10 &&
        (s.name?.startsWith('LaTeX_Shape_') || s.altTextTitle?.startsWith('LaTeX:'))
      ) {
        s.delete();
        deleted++;
      }
    }

    await context.sync();
    return deleted;
  });
}

/**
 * Installs optional worksheet change handler for legacy environments.
 */
export async function installKatexShapeAutoRenderer(): Promise<void> {
  // Available hook for auto-rendering background jobs
}

export const excelService: ExcelService = {
  insertFormulaToActiveCell,
  insertInCellImageToActiveCell,
  insertFloatingShapeToActiveCell,
  readActiveCellFormula,
  batchConvertSelectedRange,
  deleteShapesInSelection
};
