import { describe, it, expect, beforeEach } from 'vitest';
import { KaTeXShapeManager, parseAddress } from '../../src/customfunctions/shapeManager';
import { resetOfficeMock, getMockExcelState } from '../../src/mocks/officeMock';

describe('KaTeXShapeManager Unit Tests', () => {
  beforeEach(() => {
    resetOfficeMock();
    KaTeXShapeManager.clear();
  });

  describe('parseAddress', () => {
    it('parses simple cell reference', () => {
      expect(parseAddress('A1')).toEqual({ cellRef: 'A1' });
      expect(parseAddress('$B$5')).toEqual({ cellRef: 'B5' });
    });

    it('parses worksheet and cell reference', () => {
      expect(parseAddress('Sheet1!C10')).toEqual({ sheetName: 'Sheet1', cellRef: 'C10' });
      expect(parseAddress('Data!$D$12')).toEqual({ sheetName: 'Data', cellRef: 'D12' });
    });

    it('parses quoted worksheet names with spaces and special characters', () => {
      expect(parseAddress("'Q1 Results'!$AA$100")).toEqual({
        sheetName: 'Q1 Results',
        cellRef: 'AA100'
      });
      expect(parseAddress("'John''s Sheet'!B2")).toEqual({
        sheetName: "John's Sheet",
        cellRef: 'B2'
      });
    });

    it('handles empty or invalid inputs gracefully', () => {
      expect(parseAddress('')).toEqual({ cellRef: 'A1' });
      expect(parseAddress(undefined as any)).toEqual({ cellRef: 'A1' });
    });
  });

  describe('Queue Operations', () => {
    it('starts with an empty queue', () => {
      expect(KaTeXShapeManager.getPendingCount()).toBe(0);
      expect(KaTeXShapeManager.getQueue()).toEqual([]);
    });

    it('enqueues single and multiple items', () => {
      KaTeXShapeManager.enqueueShape('x = 1', 'Sheet1!A1');
      expect(KaTeXShapeManager.getPendingCount()).toBe(1);

      KaTeXShapeManager.enqueueShape('y = 2', 'Sheet1!A2', { fontSize: 20 });
      expect(KaTeXShapeManager.getPendingCount()).toBe(2);

      const queue = KaTeXShapeManager.getQueue();
      expect(queue[0].latex).toBe('x = 1');
      expect(queue[1].latex).toBe('y = 2');
      expect(queue[1].options?.fontSize).toBe(20);
    });

    it('ignores empty latex strings during enqueuing', () => {
      KaTeXShapeManager.enqueueShape('', 'Sheet1!A1');
      KaTeXShapeManager.enqueueShape('   ', 'Sheet1!A2');
      expect(KaTeXShapeManager.getPendingCount()).toBe(0);
    });

    it('clears queue successfully', () => {
      KaTeXShapeManager.enqueueShape('a + b = c', 'Sheet1!B2');
      KaTeXShapeManager.enqueueShape('d + e = f', 'Sheet1!B3');
      expect(KaTeXShapeManager.getPendingCount()).toBe(2);

      KaTeXShapeManager.clear();
      expect(KaTeXShapeManager.getPendingCount()).toBe(0);
    });
  });

  describe('processQueue', () => {
    it('returns 0 when queue is empty', async () => {
      const processed = await KaTeXShapeManager.processQueue();
      expect(processed).toBe(0);
    });

    it('processes queued shapes and inserts images into worksheet', async () => {
      const state = getMockExcelState();
      const sheet = state.workbook.worksheets.getActiveWorksheet();
      const cell = sheet.getRange('B4');
      cell.left = 100;
      cell.top = 150;

      KaTeXShapeManager.enqueueShape('\\int_0^1 x dx', 'Sheet1!B4', { fontSize: 16 });
      expect(KaTeXShapeManager.getPendingCount()).toBe(1);

      const processedCount = await KaTeXShapeManager.processQueue();
      expect(processedCount).toBe(1);
      expect(KaTeXShapeManager.getPendingCount()).toBe(0);

      expect(sheet.shapes.items.length).toBe(1);
      const shape = sheet.shapes.items[0];
      expect(shape.left).toBe(100);
      expect(shape.top).toBe(150);
      expect(shape.altTextTitle).toBe('LaTeX: \\int_0^1 x dx');
      expect(shape.width).toBeGreaterThan(0);
      expect(shape.height).toBeGreaterThan(0);
    });

    it('processes batch queue items sequentially', async () => {
      const state = getMockExcelState();
      const sheet = state.workbook.worksheets.getActiveWorksheet();

      for (let i = 1; i <= 5; i++) {
        KaTeXShapeManager.enqueueShape(`E_${i} = m c^2`, `Sheet1!A${i}`);
      }

      expect(KaTeXShapeManager.getPendingCount()).toBe(5);
      const count = await KaTeXShapeManager.processQueue();
      expect(count).toBe(5);
      expect(KaTeXShapeManager.getPendingCount()).toBe(0);
      expect(sheet.shapes.items.length).toBe(5);
    });
  });
});
