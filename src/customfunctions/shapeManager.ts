/**
 * Asynchronous Floating Shape Fallback Queue Manager for ExcelKaTeX.
 * Manages queuing and batch placement of floating equation shapes (sheet.shapes.addImage)
 * anchored to specific worksheet cell coordinates when in-cell EntityCellValue is not supported or requested.
 */

import { compileLatex, RenderOptions } from '../core/katexEngine';

export interface ShapeQueueItem {
  latex: string;
  address: string;
  options?: RenderOptions;
  timestamp?: number;
}

/**
 * Parses an Excel range address string into worksheet name and cell reference.
 * Handles formats like 'Sheet1!A1', "'Sheet Name'!$B$5", "A1".
 */
export function parseAddress(fullAddress: string): { sheetName?: string; cellRef: string } {
  if (!fullAddress || typeof fullAddress !== 'string') {
    return { cellRef: 'A1' };
  }

  const trimmed = fullAddress.trim();
  const exclamationIdx = trimmed.lastIndexOf('!');

  if (exclamationIdx === -1) {
    return { cellRef: trimmed.replace(/\$/g, '') };
  }

  let sheetPart = trimmed.substring(0, exclamationIdx).trim();
  const cellPart = trimmed.substring(exclamationIdx + 1).trim();

  // Strip single quotes around sheet name if present
  if (sheetPart.startsWith("'") && sheetPart.endsWith("'")) {
    sheetPart = sheetPart.substring(1, sheetPart.length - 1).replace(/''/g, "'");
  }

  return {
    sheetName: sheetPart,
    cellRef: cellPart.replace(/\$/g, '')
  };
}

export class KaTeXShapeManager {
  private static queue: ShapeQueueItem[] = [];
  private static isProcessing = false;
  private static autoFlushTimer: ReturnType<typeof setTimeout> | null = null;
  private static autoFlushDelayMs = 50;

  /**
   * Enqueues an equation to be inserted as a floating shape anchored at the specified cell address.
   * Debounces queue auto-flush at 50ms to automatically process batched requests.
   */
  public static enqueueShape(
    latex: string,
    address: string = 'Sheet1!A1',
    options?: RenderOptions
  ): void {
    if (!latex || !latex.trim()) {
      return;
    }

    this.queue.push({
      latex: latex.trim(),
      address: address.trim() || 'Sheet1!A1',
      options: options || {},
      timestamp: Date.now()
    });

    // Schedule debounced auto-flush
    if (this.autoFlushTimer) {
      clearTimeout(this.autoFlushTimer);
    }
    this.autoFlushTimer = setTimeout(() => {
      this.autoFlushTimer = null;
      this.processQueue().catch((err) => {
        console.error('[KaTeXShapeManager] Error in debounced auto-flush:', err);
      });
    }, this.autoFlushDelayMs);
  }

  /**
   * Returns number of pending items in the floating shape queue.
   */
  public static getPendingCount(): number {
    return this.queue.length;
  }

  /**
   * Returns a copy of the current pending shape queue.
   */
  public static getQueue(): ShapeQueueItem[] {
    return [...this.queue];
  }

  /**
   * Clears all pending shape requests from the queue and cancels any pending auto-flush timer.
   */
  public static clear(): void {
    if (this.autoFlushTimer) {
      clearTimeout(this.autoFlushTimer);
      this.autoFlushTimer = null;
    }
    this.queue = [];
  }

  /**
   * Processes all queued shape requests asynchronously in Excel using a two-phase batch loading pattern:
   * Phase 1: Compile LaTeX and request property load on target cell ranges (`range.load(['left', 'top', 'width', 'height'])`)
   * Phase 2: `await context.sync()` to resolve dimensions from Excel host
   * Phase 3: Insert image shapes at resolved coordinates (`sheet.shapes.addImage(...)`)
   * Phase 4: Final `await context.sync()` to commit all shape insertions
   *
   * @returns The number of shapes successfully processed and inserted.
   */
  public static async processQueue(): Promise<number> {
    if (this.isProcessing || this.queue.length === 0) {
      return 0;
    }

    // Cancel pending auto-flush timer since we are processing immediately
    if (this.autoFlushTimer) {
      clearTimeout(this.autoFlushTimer);
      this.autoFlushTimer = null;
    }

    this.isProcessing = true;
    const itemsToProcess = [...this.queue];
    this.queue = [];

    try {
      if (typeof (globalThis as any).Excel !== 'undefined' && typeof (globalThis as any).Excel.run === 'function') {
        await (globalThis as any).Excel.run(async (context: any) => {
          const workbook = context.workbook;
          const preparedItems: Array<{
            item: ShapeQueueItem;
            render: any;
            targetSheet: any;
            range: any;
          }> = [];

          // Phase 1: Iterate queued items and call range.load(['left', 'top', 'width', 'height'])
          for (const item of itemsToProcess) {
            try {
              const render = await compileLatex(item.latex, item.options);
              const { sheetName, cellRef } = parseAddress(item.address);

              let targetSheet = workbook.worksheets.getActiveWorksheet();
              if (sheetName && typeof workbook.worksheets.getItem === 'function') {
                try {
                  const namedSheet = workbook.worksheets.getItem(sheetName);
                  if (namedSheet) {
                    targetSheet = namedSheet;
                  }
                } catch {
                  // Fall back to active worksheet
                }
              }

              let range = targetSheet.getRange(cellRef || item.address);
              if (typeof range.load === 'function') {
                range.load(['left', 'top', 'width', 'height']);
              }

              preparedItems.push({
                item,
                render,
                targetSheet,
                range
              });
            } catch (prepErr) {
              console.error(`[KaTeXShapeManager] Error preparing shape for ${item.address}:`, prepErr);
            }
          }

          // Phase 2: Batch sync to resolve range coordinates from Excel host
          await context.sync();

          // Phase 3: Iterate queued items with resolved dimensions and insert shapes
          for (const prep of preparedItems) {
            try {
              let rawBase64 = prep.render.pngDataUrl || '';
              if (rawBase64.includes(',')) {
                rawBase64 = rawBase64.split(',')[1];
              }
              const shape = prep.targetSheet.shapes.addImage(rawBase64);
              shape.left = prep.range.left ?? 0;
              shape.top = prep.range.top ?? 0;
              shape.width = prep.render.width;
              shape.height = prep.render.height;
              shape.altTextTitle = `LaTeX: ${prep.item.latex}`;
              shape.altTextDescription = `Rendered LaTeX equation: ${prep.item.latex}`;
            } catch (insertErr) {
              console.error(`[KaTeXShapeManager] Error inserting shape for ${prep.item.address}:`, insertErr);
            }
          }

          // Phase 4: Final context sync to commit all created shapes
          await context.sync();
        });
      }

      return itemsToProcess.length;
    } finally {
      this.isProcessing = false;
    }
  }

  /**
   * Returns whether the manager is currently processing the queue.
   */
  public static getProcessingStatus(): boolean {
    return this.isProcessing;
  }
}
