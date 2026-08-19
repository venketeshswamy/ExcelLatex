/**
 * Custom Excel Function: =MATH.KATEX(...)
 *
 * ARCHITECTURE NOTE (important):
 *   Custom Functions run in a sandboxed JS context. They CANNOT call Office.js
 *   APIs (Excel.run, sheet.shapes, etc.) directly — that is prohibited by the
 *   Office platform and will silently fail or error.
 *
 *   Therefore this function:
 *     1. Validates the LaTeX string.
 *     2. Returns an in-cell rich entity (Excel.EntityCellValue) on modern Excel (ExcelApi 1.16+).
 *     3. Or enqueues shape rendering to KaTeXShapeManager and returns a placeholder text.
 */

import { compileLatex } from '../core/katexEngine';
import { parseMathKatexParams } from './parameterParser';
import {
  buildKatexEntityCellValue,
  buildKatexWebImageCellValue,
  isEntityCellValueSupported,
  KatexEntityCellValue,
  KatexWebImageCellValue
} from './entityCellBuilder';
import { KaTeXShapeManager } from './shapeManager';

export { isEntityCellValueSupported } from './entityCellBuilder';

/**
 * Custom function error creator compatible with both Office runtime and Node/Vitest.
 */
export function createCustomFunctionsError(message: string, code = '#VALUE!'): Error {
  const CF = (globalThis as any).CustomFunctions;
  if (CF && CF.Error) {
    return new CF.Error(CF.ErrorCode?.invalidValue || code, message);
  }
  const err = new Error(message);
  err.name = 'CustomFunctions.Error';
  (err as any).code = code;
  return err;
}

/**
 * =MATH.KATEX() Custom Function.
 *
 * Dual-Engine Architecture:
 * - Modern Excel (ExcelApi 1.16+): Evaluates capability and returns in-cell rich image (Excel.EntityCellValue or WebImageCellValue).
 * - Legacy Excel / Shape Mode: Enqueues shape rendering to KaTeXShapeManager and returns [KaTeX Shape: <latex>].
 */
export async function mathKatex(
  latexString: string,
  background?: any,
  color?: any,
  fontSize?: any,
  displayMode?: any,
  outputMethod?: any,
  invocation?: any
): Promise<KatexEntityCellValue | KatexWebImageCellValue | string> {
  // Validate & parse all parameters (handles multi-slot invocation shifting and delimiter stripping)
  const validation = parseMathKatexParams(
    latexString,
    background,
    color,
    fontSize,
    displayMode,
    outputMethod,
    invocation
  );

  if (!validation.isValid || !validation.params) {
    throw createCustomFunctionsError(
      validation.error || 'Invalid LaTeX parameter.',
      validation.errorCode || '#VALUE!'
    );
  }

  const {
    latexString: latex,
    background: bg,
    color: fg,
    fontSize: size,
    displayMode: mode,
    outputMethod: method,
    cellAddress
  } = validation.params;

  const options = { background: bg, color: fg, fontSize: size, displayMode: mode };

  const supportsEntity = isEntityCellValueSupported();
  const shouldRenderCell = method === 'cell' || method === 'image' || (method === 'auto' && supportsEntity);

  if (shouldRenderCell) {
    try {
      const renderResult = await compileLatex(latex, options);
      if (method === 'image') {
        return buildKatexWebImageCellValue(latex, renderResult);
      }
      return buildKatexEntityCellValue(latex, renderResult, options);
    } catch (err: any) {
      throw createCustomFunctionsError(
        err?.message || 'Failed to render LaTeX equation.',
        '#VALUE!'
      );
    }
  }

  // Fallback / Shape routing: Enqueue shape into KaTeXShapeManager and return placeholder
  KaTeXShapeManager.enqueueShape(latex, cellAddress, options);
  return `[KaTeX Shape: ${latex}]`;
}

// Alias for backwards compatibility with tests and callers
export const mathKatexFunction = mathKatex;

// Auto-register with Office CustomFunctions runtime if present
if (
  typeof (globalThis as any).CustomFunctions !== 'undefined' &&
  typeof (globalThis as any).CustomFunctions.associate === 'function'
) {
  try {
    (globalThis as any).CustomFunctions.associate('KATEX', mathKatex);
    (globalThis as any).CustomFunctions.associate('MATH.KATEX', mathKatex);
  } catch {
    // Silent — already registered or in mock
  }
}
