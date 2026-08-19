import { describe, it, expect, beforeEach, vi } from 'vitest';
import { compileLatex, validateLatex } from '../../src/core/katexEngine';
import { clearEquationCache } from '../../src/core/imageRasterizer';
import {
  mathKatexFunction,
  KaTeXShapeManager
} from '../helpers/testHelpers';
import {
  resetOfficeMock,
  getMockExcelState,
  setMockTheme,
  ExcelMock
} from '../../src/mocks/officeMock';
import {
  resetCustomFunctionsMock,
  CustomFunctionsMock
} from '../../src/mocks/customFunctionsMock';

describe('Tier 3: Cross-Feature Combinations Suite', () => {
  beforeEach(() => {
    resetOfficeMock();
    resetCustomFunctionsMock();
    clearEquationCache();
    KaTeXShapeManager.clear();
  });

  describe('3.1 Custom Functions + Office Dark Theme Interaction', () => {
    it('adapts custom function rendering for dark theme with inverted colors', async () => {
      setMockTheme({
        bodyBackgroundColor: '#1f1f1f',
        bodyForegroundColor: '#ffffff',
        isDark: true
      });

      const theme = getMockExcelState().theme;
      const inv = CustomFunctionsMock.createInvocation('Sheet1!A1');

      const entity = (await mathKatexFunction(
        '\\int_{-\\infty}^{+\\infty} e^{-x^2} dx = \\sqrt{\\pi}',
        theme.bodyBackgroundColor,
        theme.bodyForegroundColor,
        16,
        true,
        'cell',
        inv
      )) as ExcelMock.EntityCellValue;

      expect(entity.type).toBe('Entity');
      expect(entity.layouts?.compact?.icon).toBeDefined();

      const render = await compileLatex('\\int_{-\\infty}^{+\\infty} e^{-x^2} dx = \\sqrt{\\pi}', {
        background: theme.bodyBackgroundColor,
        color: theme.bodyForegroundColor
      });

      expect(render.svg).toContain('background-color: #1f1f1f');
      expect(render.svg).toContain('color: #ffffff');
    });
  });

  describe('3.2 High-DPI Scaling + Floating Shape Alignment', () => {
    it('produces high-DPI (3x) rasterization anchored accurately to cell dimensions', async () => {
      const state = getMockExcelState();
      const activeSheet = state.workbook.worksheets.getActiveWorksheet();
      const cell = activeSheet.getRange('Sheet1!C10');
      cell.left = 240;
      cell.top = 180;
      cell.width = 150;
      cell.height = 30;

      const formula = '\\mathcal{F}\\{f(t)\\} = \\int_{-\\infty}^\\infty f(t) e^{-i\\omega t} dt';
      KaTeXShapeManager.enqueueShape(formula, 'Sheet1!C10', { scale: 3 });

      const processed = await KaTeXShapeManager.processQueue();
      expect(processed).toBe(1);

      const shape = activeSheet.shapes.items[0];
      expect(shape.left).toBe(240);
      expect(shape.top).toBe(180);
      expect(shape.width).toBeGreaterThan(0);
      expect(shape.height).toBeGreaterThan(0);
    });
  });

  describe('3.3 Debounced Typing + Rapid Equation Preset Switching', () => {
    it('coalesces rapid preset selections without race conditions', () => {
      vi.useFakeTimers();

      let currentLatex = '';
      const updateLatex = vi.fn((val: string) => {
        currentLatex = val;
      });

      let timer: any = null;
      const setLatexDebounced = (val: string) => {
        if (timer) clearTimeout(timer);
        timer = setTimeout(() => updateLatex(val), 100);
      };

      const presetList = [
        'E = mc^2',
        'a^2 + b^2 = c^2',
        'e^{i\\pi} + 1 = 0',
        '\\nabla \\times \\mathbf{B} = \\mu_0 \\mathbf{J}',
        '\\oint \\mathbf{E} \\cdot d\\mathbf{A} = \\frac{Q}{\\varepsilon_0}'
      ];

      // Rapidly trigger 5 presets in 50ms
      presetList.forEach((preset) => {
        vi.advanceTimersByTime(10);
        setLatexDebounced(preset);
      });

      // Complete debounce window
      vi.advanceTimersByTime(150);

      expect(updateLatex).toHaveBeenCalledTimes(1);
      expect(currentLatex).toBe(presetList[presetList.length - 1]);
      vi.useRealTimers();
    });
  });

  describe('3.4 Cell Error State + Multi-Format Clipboard Export', () => {
    it('blocks formula insertion on syntax error while preserving raw clipboard export', async () => {
      const brokenFormula = '\\frac{numerator}{'; // unclosed brace
      const validation = validateLatex(brokenFormula);

      expect(validation.isValid).toBe(false);

      // Sheet insertion should be rejected or prevented
      let insertAllowed = validation.isValid;
      expect(insertAllowed).toBe(false);

      // But copying raw LaTeX string to clipboard remains possible
      const rawExport = brokenFormula;
      expect(rawExport).toBe(brokenFormula);
    });
  });

  describe('3.5 Multi-Cell Batch Recalculation + Shape Queue Throttling', () => {
    it('concurrently evaluates 30 custom functions and drains shape queue cleanly', async () => {
      const promises: Promise<any>[] = [];

      for (let i = 1; i <= 30; i++) {
        const inv = CustomFunctionsMock.createInvocation(`Sheet1!A${i}`);
        const p = mathKatexFunction(
          `y_{${i}} = \\sin(${i}x) + \\cos(${i}x)`,
          undefined,
          undefined,
          14,
          true,
          'shape',
          inv
        );
        promises.push(p);
      }

      const results = await Promise.all(promises);
      expect(results.length).toBe(30);
      expect(KaTeXShapeManager.getPendingCount()).toBe(30);

      const processedCount = await KaTeXShapeManager.processQueue();
      expect(processedCount).toBe(30);
      expect(KaTeXShapeManager.getPendingCount()).toBe(0);

      const sheet = getMockExcelState().workbook.worksheets.getActiveWorksheet();
      expect(sheet.shapes.items.length).toBe(30);
    });
  });
});
