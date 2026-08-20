import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import {
  mathKatex,
  isEntityCellValueSupported
} from '../src/customfunctions/mathKatex';
import {
  parseMathKatexParams,
  stripLatexDelimiters,
  sanitizeColor,
  sanitizeFontSize,
  sanitizeOutputMethod
} from '../src/customfunctions/parameterParser';
import {
  KatexEntityCellValue
} from '../src/customfunctions/entityCellBuilder';
import {
  KaTeXShapeManager,
  parseAddress
} from '../src/customfunctions/shapeManager';
import {
  compileLatex,
  validateLatex
} from '../src/core/katexEngine';
import {
  clearEquationCache
} from '../src/core/imageRasterizer';
import {
  insertFormulaToActiveCell,
  insertInCellImageToActiveCell,
  escapeFormulaString
} from '../src/taskpane/services/excelService';
import {
  resetOfficeMock,
  setMockRequirementSupported,
  getMockExcelState
} from '../src/mocks/officeMock';
import {
  resetCustomFunctionsMock,
  CustomFunctionsMock
} from '../src/mocks/customFunctionsMock';

describe('Empirical Challenger Harness: M2 & M3 Deep Verification', () => {
  beforeEach(() => {
    resetOfficeMock();
    resetCustomFunctionsMock();
    KaTeXShapeManager.clear();
    clearEquationCache();
  });

  afterEach(() => {
    vi.restoreAllMocks();
  });

  // =========================================================================
  // Challenge 1: Dual-Engine Requirement Detection & Output Routing
  // =========================================================================
  describe('1. Dual-Engine Capability Detection & Routing Toggles', () => {
    it('accurately detects ExcelApi 1.16 capability across various Office context states', () => {
      // Modern Excel: isSetSupported('ExcelApi', 1.16) returns true
      setMockRequirementSupported('ExcelApi', 1.16, true);
      expect(isEntityCellValueSupported()).toBe(true);

      // Legacy Excel: isSetSupported('ExcelApi', 1.16) returns false
      setMockRequirementSupported('ExcelApi', 1.16, false);
      expect(isEntityCellValueSupported()).toBe(false);

      // Office global missing
      const oldOffice = (globalThis as any).Office;
      try {
        delete (globalThis as any).Office;
        expect(isEntityCellValueSupported()).toBe(false);
      } finally {
        (globalThis as any).Office = oldOffice;
      }
    });

    it('returns rich EntityCellValue in modern mode with auto routing', async () => {
      setMockRequirementSupported('ExcelApi', 1.16, true);
      const inv = CustomFunctionsMock.createInvocation('Sheet1!A1');
      const res = await mathKatex('\\frac{-b \\pm \\sqrt{b^2 - 4ac}}{2a}', undefined, undefined, 18, true, 'auto', inv);

      expect(typeof res).toBe('object');
      const entity = res as KatexEntityCellValue;
      expect(entity.type).toBe('Entity');
      expect(entity.text).toBe('[Math: \\frac{-b \\pm \\sqrt{b^2 - 4ac}}{2a}]');
      expect(entity.properties?.latex?.basicValue).toBe('\\frac{-b \\pm \\sqrt{b^2 - 4ac}}{2a}');
      expect(entity.properties?.fontSize?.basicValue).toBe(18);
      expect(entity.properties?.displayMode?.basicValue).toBe(true);
      expect(entity.properties?.image?.address).toMatch(/^data:image\/png;base64,/);
      expect(entity.properties?.svg?.basicValue).toContain('<svg');
      expect(KaTeXShapeManager.getPendingCount()).toBe(0);
    });

    it('falls back to shape queue in legacy mode with auto routing', async () => {
      setMockRequirementSupported('ExcelApi', 1.16, false);
      const inv = CustomFunctionsMock.createInvocation('Sheet1!B5');
      const res = await mathKatex('\\sum_{i=1}^n i', undefined, undefined, undefined, undefined, 'auto', inv);

      expect(typeof res).toBe('string');
      expect(res).toBe('[KaTeX Shape: \\sum_{i=1}^n i]');
      expect(KaTeXShapeManager.getPendingCount()).toBe(1);
      const queued = KaTeXShapeManager.getQueue()[0];
      expect(queued.address).toBe('Sheet1!B5');
      expect(queued.latex).toBe('\\sum_{i=1}^n i');
    });

    it('respects explicit outputMethod="cell" even if capability mock is toggled false', async () => {
      setMockRequirementSupported('ExcelApi', 1.16, false);
      const inv = CustomFunctionsMock.createInvocation('Sheet1!C3');
      const res = await mathKatex('x^2 + y^2 = r^2', undefined, undefined, undefined, undefined, 'cell', inv);

      expect(typeof res).toBe('object');
      expect((res as KatexEntityCellValue).type).toBe('Entity');
      expect(KaTeXShapeManager.getPendingCount()).toBe(0);
    });

    it('respects explicit outputMethod="shape" even if capability mock is toggled true', async () => {
      setMockRequirementSupported('ExcelApi', 1.16, true);
      const inv = CustomFunctionsMock.createInvocation('Sheet1!D4');
      const res = await mathKatex('\\oint_C \\mathbf{F} \\cdot d\\mathbf{r}', undefined, undefined, undefined, undefined, 'shape', inv);

      expect(typeof res).toBe('string');
      expect(res).toBe('[KaTeX Shape: \\oint_C \\mathbf{F} \\cdot d\\mathbf{r}]');
      expect(KaTeXShapeManager.getPendingCount()).toBe(1);
    });

    it('correctly handles invocation passed in 6th argument position (standard Excel CF behavior)', async () => {
      setMockRequirementSupported('ExcelApi', 1.16, false);
      const inv = CustomFunctionsMock.createInvocation('CustomSheet!K10');
      // When user enters =MATH.KATEX("E=mc^2"), Excel calls mathKatex("E=mc^2", undefined, undefined, undefined, undefined, inv)
      const res = await mathKatex('E = mc^2', undefined, undefined, undefined, undefined, inv);

      expect(typeof res).toBe('string');
      expect(res).toBe('[KaTeX Shape: E = mc^2]');
      expect(KaTeXShapeManager.getPendingCount()).toBe(1);
      expect(KaTeXShapeManager.getQueue()[0].address).toBe('CustomSheet!K10');
    });
  });

  // =========================================================================
  // Challenge 2: Parameter Parsing & Delimiter Stripping Stress
  // =========================================================================
  describe('2. Parameter Parsing, Delimiters & Escaping Stress', () => {
    it('strips all standard LaTeX delimiter variations cleanly', () => {
      const cases = [
        { raw: '$$E = mc^2$$', expected: 'E = mc^2' },
        { raw: '  $$  \\alpha + \\beta  $$  ', expected: '\\alpha + \\beta' },
        { raw: '\\[ \\int x dx \\]', expected: '\\int x dx' },
        { raw: '  \\[  \\sum_{i=1}^n i^2  \\]  \n', expected: '\\sum_{i=1}^n i^2' },
        { raw: '\\( a^2 + b^2 = c^2 \\)', expected: 'a^2 + b^2 = c^2' },
        { raw: '  \\(  x \\le y  \\)  ', expected: 'x \\le y' },
        { raw: '$x + y = z$', expected: 'x + y = z' },
        { raw: '  $  \\lim_{x \\to 0} \\frac{\\sin x}{x} = 1  $  ', expected: '\\lim_{x \\to 0} \\frac{\\sin x}{x} = 1' },
      ];

      for (const { raw, expected } of cases) {
        expect(stripLatexDelimiters(raw)).toBe(expected);
        const parsed = parseMathKatexParams(raw);
        expect(parsed.isValid).toBe(true);
        expect(parsed.params?.latexString).toBe(expected);
      }
    });

    it('preserves inner dollar signs and math constructs when stripping delimiters', () => {
      const latexWithEscapedDollar = '$$\\text{Cost is } \\$ 100$$';
      expect(stripLatexDelimiters(latexWithEscapedDollar)).toBe('\\text{Cost is } \\$ 100');

      const innerMatrix = '\\[ \\begin{pmatrix} 1 & 0 \\\\ 0 & 1 \\end{pmatrix} \\]';
      expect(stripLatexDelimiters(innerMatrix)).toBe('\\begin{pmatrix} 1 & 0 \\\\ 0 & 1 \\end{pmatrix}');
    });

    it('preserves intentional LaTeX space commands (\\ , \\quad, \\qquad, \\text{ })', () => {
      const spaceExpr = 'a \\ b \\quad c \\qquad d \\text{ hello world } e';
      const stripped = stripLatexDelimiters(`$$${spaceExpr}$$`);
      expect(stripped).toBe(spaceExpr);

      const val = validateLatex(stripped);
      expect(val.isValid).toBe(true);
    });

    it('sanitizes background and foreground color strings across all valid CSS color formats', () => {
      expect(sanitizeColor('#fff', 'transparent')).toBe('#fff');
      expect(sanitizeColor('#107c41', 'transparent')).toBe('#107c41');
      expect(sanitizeColor('#107c4180', 'transparent')).toBe('#107c4180');
      expect(sanitizeColor('rgb(255, 0, 0)', 'transparent')).toBe('rgb(255, 0, 0)');
      expect(sanitizeColor('rgba(255, 0, 0, 0.8)', 'transparent')).toBe('rgba(255, 0, 0, 0.8)');
      expect(sanitizeColor('hsl(120, 100%, 50%)', 'transparent')).toBe('hsl(120, 100%, 50%)');
      expect(sanitizeColor('red', 'transparent')).toBe('red');
      expect(sanitizeColor('transparent', '#000')).toBe('transparent');
      expect(sanitizeColor('currentColor', '#000')).toBe('currentColor');

      // Invalid / injection color strings fall back to default
      expect(sanitizeColor('javascript:alert(1)', 'transparent')).toBe('transparent');
      expect(sanitizeColor('<style>body{color:red}</style>', '#000')).toBe('#000');
      expect(sanitizeColor('invalidColor()', '#000000')).toBe('#000000');
      expect(sanitizeColor('   ', '#000000')).toBe('#000000');
    });

    it('clamps font sizes strictly to [4, 1000] and falls back to 16 for invalid numbers', () => {
      expect(sanitizeFontSize(undefined)).toBe(16);
      expect(sanitizeFontSize(null as any)).toBe(16);
      expect(sanitizeFontSize(NaN)).toBe(16);
      expect(sanitizeFontSize(-5)).toBe(16);
      expect(sanitizeFontSize(0)).toBe(16);
      expect(sanitizeFontSize(1)).toBe(4);
      expect(sanitizeFontSize(3.9)).toBe(4);
      expect(sanitizeFontSize(4)).toBe(4);
      expect(sanitizeFontSize(24)).toBe(24);
      expect(sanitizeFontSize(1000)).toBe(1000);
      expect(sanitizeFontSize(1001)).toBe(1000);
      expect(sanitizeFontSize(50000)).toBe(1000);
    });

    it('sanitizes outputMethod with case-insensitivity and default fallback', () => {
      expect(sanitizeOutputMethod('auto')).toBe('auto');
      expect(sanitizeOutputMethod('AUTO')).toBe('auto');
      expect(sanitizeOutputMethod('cell')).toBe('cell');
      expect(sanitizeOutputMethod('CELL')).toBe('cell');
      expect(sanitizeOutputMethod('shape')).toBe('shape');
      expect(sanitizeOutputMethod('SHAPE')).toBe('shape');
      expect(sanitizeOutputMethod('unknown')).toBe('auto');
      expect(sanitizeOutputMethod('')).toBe('auto');
      expect(sanitizeOutputMethod(undefined)).toBe('auto');
    });
  });

  // =========================================================================
  // Challenge 3: Error Surfacing & #VALUE! Reporting
  // =========================================================================
  describe('3. Error Surfacing & #VALUE! Reporting', () => {
    it('surfaces #VALUE! on empty or whitespace equations', async () => {
      const emptySamples = ['', '   ', '\t', '\n\r', '  \t  \n'];

      for (const sample of emptySamples) {
        const parsed = parseMathKatexParams(sample);
        expect(parsed.isValid).toBe(false);
        expect(parsed.errorCode).toBe('#VALUE!');

        try {
          await mathKatex(sample);
          expect.unreachable('Should have thrown');
        } catch (err: any) {
          expect((err as any).code).toBe('#VALUE!');
        }
      }
    });

    it('surfaces #VALUE! on non-string inputs', async () => {
      const nonStrings = [null, undefined, 123 as any, false as any, {} as any];

      for (const sample of nonStrings) {
        const parsed = parseMathKatexParams(sample);
        expect(parsed.isValid).toBe(false);
        expect(parsed.errorCode).toBe('#VALUE!');

        try {
          await mathKatex(sample);
          expect.unreachable('Should have thrown');
        } catch (err: any) {
          expect((err as any).code).toBe('#VALUE!');
        }
      }
    });

    it('surfaces #VALUE! on syntax errors without unhandled exceptions', async () => {
      const brokenSyntax = [
        '\\frac{1}{',
        '\\sqrt[',
        '\\begin{matrix} 1 & 2',
        '\\begin{pmatrix} 1 & 2 \\end{bmatrix}',
        '\\unknownCommand{xyz}',
        'x^2^3',
        '\\left( x + y'
      ];

      for (const broken of brokenSyntax) {
        const parsed = parseMathKatexParams(broken);
        expect(parsed.isValid).toBe(false);
        expect(parsed.errorCode).toBe('#VALUE!');
        expect(parsed.error).toBeDefined();

        try {
          await mathKatex(broken);
          expect.unreachable('Should have thrown');
        } catch (err: any) {
          expect((err as any).code).toBe('#VALUE!');
        }
      }
    });

    it('surfaces #VALUE! on XSS and script tags', async () => {
      const xssInputs = [
        '<script>alert("xss")</script>',
        '<img src=x onerror=alert(1)>',
        'javascript:void(0)'
      ];

      for (const xss of xssInputs) {
        const parsed = parseMathKatexParams(xss);
        expect(parsed.isValid).toBe(false);
        expect(parsed.errorCode).toBe('#VALUE!');

        try {
          await mathKatex(xss);
          expect.unreachable('Should have thrown');
        } catch (err: any) {
          expect((err as any).code).toBe('#VALUE!');
        }
      }
    });
  });

  // =========================================================================
  // Challenge 4: Shape Manager Queue, Deduplication & Concurrency
  // =========================================================================
  describe('4. Shape Manager Queue, Debouncing & Concurrency', () => {
    it('correctly parses Excel cell addresses with various sheet syntax', () => {
      expect(parseAddress('Sheet1!A1')).toEqual({ sheetName: 'Sheet1', cellRef: 'A1' });
      expect(parseAddress("'Sales Report'!$B$10")).toEqual({ sheetName: 'Sales Report', cellRef: 'B10' });
      expect(parseAddress("'O''Reilly Data'!$AA$100")).toEqual({ sheetName: "O'Reilly Data", cellRef: 'AA100' });
      expect(parseAddress('C15')).toEqual({ cellRef: 'C15' });
      expect(parseAddress('$D$20')).toEqual({ cellRef: 'D20' });
      expect(parseAddress('')).toEqual({ cellRef: 'A1' });
    });

    it('batches 50 rapid shape queue additions and drains them in a single batch', async () => {
      for (let i = 1; i <= 50; i++) {
        KaTeXShapeManager.enqueueShape(`x_{${i}} + y_{${i}} = z_{${i}}`, `Sheet1!A${i}`);
      }

      expect(KaTeXShapeManager.getPendingCount()).toBe(50);
      const queue = KaTeXShapeManager.getQueue();
      expect(queue.length).toBe(50);
      expect(queue[0].address).toBe('Sheet1!A1');
      expect(queue[49].address).toBe('Sheet1!A50');

      const processedCount = await KaTeXShapeManager.processQueue();
      expect(processedCount).toBe(50);
      expect(KaTeXShapeManager.getPendingCount()).toBe(0);

      const state = getMockExcelState();
      const sheet = state.workbook.worksheets.getActiveWorksheet();
      expect(sheet.shapes.items.length).toBe(50);
    });

    it('prevents concurrent double-draining races via isProcessing lock', async () => {
      KaTeXShapeManager.enqueueShape('a + b = c', 'Sheet1!A1');
      KaTeXShapeManager.enqueueShape('d + e = f', 'Sheet1!A2');

      const [count1, count2] = await Promise.all([
        KaTeXShapeManager.processQueue(),
        KaTeXShapeManager.processQueue()
      ]);

      expect(count1 + count2).toBe(2);
      expect(KaTeXShapeManager.getPendingCount()).toBe(0);
    });

    it('attaches alt-text metadata to inserted shapes for deduplication and accessibility', async () => {
      const latex = '\\int_0^\\infty e^{-x^2} dx = \\frac{\\sqrt{\\pi}}{2}';
      KaTeXShapeManager.enqueueShape(latex, 'Sheet1!E10');
      await KaTeXShapeManager.processQueue();

      const state = getMockExcelState();
      const sheet = state.workbook.worksheets.getActiveWorksheet();
      expect(sheet.shapes.items.length).toBe(1);

      const shape = sheet.shapes.items[0];
      expect(shape.altTextTitle).toBe(`LaTeX: ${latex}`);
      expect(shape.altTextDescription).toBe(`Rendered LaTeX equation: ${latex}`);
      expect(shape.width).toBeGreaterThan(0);
      expect(shape.height).toBeGreaterThan(0);
    });
  });

  // =========================================================================
  // Challenge 5: Taskpane Excel Service Formula Formatting & Quotes
  // =========================================================================
  describe('5. Taskpane Excel Service & Formula Generation', () => {
    it('escapes internal double quotes correctly for Excel formulas', () => {
      expect(escapeFormulaString('E = mc^2')).toBe('E = mc^2');
      expect(escapeFormulaString('\\text{"Hello World"}')).toBe('\\text{""Hello World""}');
      expect(escapeFormulaString('a = "test", b = "demo"')).toBe('a = ""test"", b = ""demo""');
    });

    it('generates compact formula =MATH.KATEX("...") for default options', async () => {
      await insertFormulaToActiveCell('E = mc^2');

      const state = getMockExcelState();
      const range = state.workbook.getSelectedRange();
      expect(range.formulas[0][0]).toBe('=MATH.KATEX("E = mc^2")');
    });

    it('generates clean formula with background when non-default background is specified', async () => {
      await insertFormulaToActiveCell('\\int x dx', {
        background: 1
      });

      const state = getMockExcelState();
      const range = state.workbook.getSelectedRange();
      expect(range.formulas[0][0]).toBe('=MATH.KATEX("\\int x dx", 1)');
    });

    it('handles In-Cell Image insertion with dual values and valuesAsJson assignment', async () => {
      setMockRequirementSupported('ExcelApi', 1.16, true);
      await insertInCellImageToActiveCell('\\cos(\\theta)^2 + \\sin(\\theta)^2 = 1');

      const state = getMockExcelState();
      const range = state.workbook.getSelectedRange();
      expect((range as any).valuesAsJson).toBeDefined();
      expect(range.values).toBeDefined();
      expect(range.values[0][0].altText).toContain('cos');
      expect(range.values[0][0].altText).toContain('sin');
    });

    it('falls back to floating shape insertion when In-Cell Image is used on legacy host', async () => {
      setMockRequirementSupported('ExcelApi', 1.16, false);
      await insertInCellImageToActiveCell('F = ma');

      const state = getMockExcelState();
      const sheet = state.workbook.worksheets.getActiveWorksheet();
      expect(sheet.shapes.items.length).toBe(1);
      expect(sheet.shapes.items[0].altTextTitle).toBe('LaTeX: F = ma');
    });
  });

  // =========================================================================
  // Challenge 6: High-Throughput & Extreme Payload Stress
  // =========================================================================
  describe('6. High-Throughput & Stress Harness', () => {
    it('handles 100 parallel calls to mathKatex without memory or promise leakage', async () => {
      setMockRequirementSupported('ExcelApi', 1.16, true);
      const promises: Promise<any>[] = [];

      for (let i = 0; i < 100; i++) {
        const inv = CustomFunctionsMock.createInvocation(`Sheet1!A${i + 1}`);
        promises.push(
          mathKatex(`f_{${i}}(x) = x^{${i % 10}} + ${i}`, undefined, undefined, 16, true, 'auto', inv)
        );
      }

      const results = await Promise.all(promises);
      expect(results.length).toBe(100);
      for (const res of results) {
        expect((res as KatexEntityCellValue).type).toBe('Entity');
      }
    });

    it('handles extremely long multi-line LaTeX with alignment and comments', async () => {
      const multiLine = `\\begin{aligned}
        \\nabla \\times \\mathbf{E} &= -\\frac{\\partial \\mathbf{B}}{\\partial t} \\\\
        \\nabla \\times \\mathbf{B} &= \\mu_0 \\mathbf{J} + \\mu_0 \\epsilon_0 \\frac{\\partial \\mathbf{E}}{\\partial t} \\\\
        \\nabla \\cdot \\mathbf{E} &= \\frac{\\rho}{\\epsilon_0} \\\\
        \\nabla \\cdot \\mathbf{B} &= 0
      \\end{aligned}`;

      const res = await compileLatex(multiLine);
      expect(res.html).toContain('katex');
      expect(res.width).toBeGreaterThan(80);
      expect(res.height).toBeGreaterThan(30);
      expect(res.pngDataUrl).toMatch(/^data:image\/png;base64,/);
    });
  });
});
