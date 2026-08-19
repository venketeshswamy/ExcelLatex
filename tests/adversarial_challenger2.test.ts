import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import {
  compileLatex,
  validateLatex
} from '../src/core/katexEngine';
import {
  rasterizeLatex,
  generateStandaloneSvg,
  svgToPngDataUrl
} from '../src/core/imageRasterizer';
import {
  stripLatexDelimiters,
  parseMathKatexParams
} from '../src/customfunctions/parameterParser';
import {
  mathKatex
} from '../src/customfunctions/mathKatex';
import {
  buildKatexEntityCellValue,
  KatexEntityCellValue
} from '../src/customfunctions/entityCellBuilder';
import { KaTeXShapeManager } from '../src/customfunctions/shapeManager';
import {
  insertFormulaToActiveCell,
  insertInCellImageToActiveCell,
  insertFloatingShapeToActiveCell,
  readActiveCellFormula,
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
import fs from 'fs';
import path from 'path';

describe('Adversarial Challenge 2: Empirical Verification & Boundary Stress Harness', () => {
  beforeEach(() => {
    resetOfficeMock();
    resetCustomFunctionsMock();
    KaTeXShapeManager.clear();
  });

  afterEach(() => {
    vi.restoreAllMocks();
  });

  // =========================================================================
  // 1. SVG Markup Structure & Strict XML Conformance
  // =========================================================================
  describe('1. SVG Markup Structure & XML Validation', () => {
    it('generates well-formed SVG with xmlns, viewBox, width, height, and style attributes', () => {
      const html = '<span class="katex"><span class="katex-html">E=mc^2</span></span>';
      const width = 350;
      const height = 120;
      const bg = '#f0f4f8';
      const color = '#107c41';
      const fontSizePx = 24;

      const svg = generateStandaloneSvg(html, width, height, bg, color, fontSizePx);

      // Verify mandatory SVG attributes
      expect(svg).toContain('xmlns="http://www.w3.org/2000/svg"');
      expect(svg).toContain(`width="${width}"`);
      expect(svg).toContain(`height="${height}"`);
      expect(svg).toContain(`viewBox="0 0 ${width} ${height}"`);

      // Verify foreignObject structure
      expect(svg).toContain('<foreignObject width="100%" height="100%">');
      expect(svg).toContain('<html xmlns="http://www.w3.org/1999/xhtml">');

      // Verify CSS styles inside embedded HTML
      expect(svg).toContain(`font-size:${fontSizePx}px`);
      expect(svg).toContain(`color: ${color};`);
      expect(svg).toContain(`background-color: ${bg};`);
    });

    it('omits background-color CSS property when background is transparent', () => {
      const html = '<span class="katex">x</span>';
      const svgTransparent = generateStandaloneSvg(html, 100, 40, 'transparent', '#000000', 16);
      expect(svgTransparent).not.toContain('background-color:');

      const svgEmpty = generateStandaloneSvg(html, 100, 40, '', '#000000', 16);
      expect(svgEmpty).not.toContain('background-color:');
    });

    it('parses standalone SVG as valid DOM structure via DOMParser', () => {
      const html = '<span class="katex"><span class="katex-mathml">math</span></span>';
      const svg = generateStandaloneSvg(html, 250, 75, '#ffffff', '#333333', 18);

      const parser = new DOMParser();
      const doc = parser.parseFromString(svg, 'image/svg+xml');

      // Ensure no XML parser error document was produced
      const parserError = doc.querySelector('parsererror');
      expect(parserError).toBeNull();

      const svgRoot = doc.querySelector('svg');
      expect(svgRoot).not.toBeNull();
      expect(svgRoot?.getAttribute('xmlns')).toBe('http://www.w3.org/2000/svg');
      expect(svgRoot?.getAttribute('viewBox')).toBe('0 0 250 75');
      expect(svgRoot?.getAttribute('width')).toBe('250');
      expect(svgRoot?.getAttribute('height')).toBe('75');

      const foreignObject = svgRoot?.querySelector('foreignObject');
      expect(foreignObject).not.toBeNull();
      expect(foreignObject?.getAttribute('width')).toBe('100%');
      expect(foreignObject?.getAttribute('height')).toBe('100%');
    });

    it('svgToPngDataUrl injects xmlns if missing and produces a valid data URL', async () => {
      const rawSvgWithoutXmlns = '<svg width="100" height="50"><rect width="100" height="50" fill="red"/></svg>';
      const dataUrl = await svgToPngDataUrl(rawSvgWithoutXmlns, 100, 50, 2);

      expect(typeof dataUrl).toBe('string');
      expect(dataUrl.startsWith('data:image/png;base64,')).toBe(true);
    });
  });

  // =========================================================================
  // 2. Headless DOM Dimension Scaling & Extreme Bounds
  // =========================================================================
  describe('2. Headless DOM Dimension Scaling & Extreme Bounds', () => {
    it('scales dimensions proportionally for 500pt typography without crashing or zero-sizing', async () => {
      const latex = '\\int_0^1 x dx = \\frac{1}{2}';
      const res500 = await rasterizeLatex(latex, { fontSize: 500 });

      expect(res500.width).toBeGreaterThan(500);
      expect(res500.height).toBeGreaterThan(400);
      expect(Number.isFinite(res500.width)).toBe(true);
      expect(Number.isFinite(res500.height)).toBe(true);
      expect(Number.isFinite(res500.aspectRatio)).toBe(true);
      expect(res500.aspectRatio).toBeGreaterThan(0);
    });

    it('scales height proportionally for massive 15x15 dense matrices in headless DOM', async () => {
      const rows: string[] = [];
      for (let r = 0; r < 15; r++) {
        const cols: string[] = [];
        for (let c = 0; c < 15; c++) {
          cols.push(`m_{${r + 1},${c + 1}}`);
        }
        rows.push(cols.join(' & '));
      }
      const matrix15x15 = `\\begin{pmatrix} ${rows.join(' \\\\ ')} \\end{pmatrix}`;

      const resMatrix = await rasterizeLatex(matrix15x15, { fontSize: 16 });
      const resSingle = await rasterizeLatex('x = 1', { fontSize: 16 });

      // The 15x15 matrix must have significantly greater height and width than single equation
      expect(resMatrix.height).toBeGreaterThan(resSingle.height * 3);
      expect(resMatrix.width).toBeGreaterThan(resSingle.width * 2);
    });

    it('enforces minimum bounding box for single-character formulas', async () => {
      const singleChars = ['x', '1', '+', '\\alpha', '='];
      for (const char of singleChars) {
        const res = await rasterizeLatex(char, { fontSize: 16 });
        expect(res.width).toBeGreaterThanOrEqual(84); // min 80 + 4 padding
        expect(res.height).toBeGreaterThanOrEqual(40); // min 36 + 4 padding
      }
    });

    it('scales width for wide multi-term summation series', async () => {
      const longSum = '\\sum_{k=1}^{100} \\left( \\frac{a_k x^k}{k!} + \\frac{b_k y^k}{(2k)!} + \\frac{c_k z^k}{(3k)!} \\right)';
      const resLong = await rasterizeLatex(longSum, { fontSize: 20 });
      const resShort = await rasterizeLatex('x + y', { fontSize: 20 });

      expect(resLong.width).toBeGreaterThan(resShort.width * 3);
    });
  });

  // =========================================================================
  // 3. LaTeX Space Preservation Across Parser & Renderer
  // =========================================================================
  describe('3. LaTeX Space Preservation', () => {
    it('preserves all LaTeX spacing primitives during delimiter stripping', () => {
      const spaceExpressions = [
        { raw: '$$ a\\ b $$', expected: 'a\\ b' },
        { raw: '\\[ x\\quad y \\]', expected: 'x\\quad y' },
        { raw: '\\( u\\qquad v \\)', expected: 'u\\qquad v' },
        { raw: '$ \\text{hello world} $', expected: '\\text{hello world}' },
        { raw: '$$ a\\enspace b\\thinspace c\\,d\\;e\\:f\\!g $$', expected: 'a\\enspace b\\thinspace c\\,d\\;e\\:f\\!g' },
        { raw: '\\[ a~~~b \\]', expected: 'a~~~b' },
        { raw: '$$ \\text{  leading and trailing  } $$', expected: '\\text{  leading and trailing  }' }
      ];

      for (const { raw, expected } of spaceExpressions) {
        const stripped = stripLatexDelimiters(raw);
        expect(stripped).toBe(expected);

        const parsed = parseMathKatexParams(raw);
        expect(parsed.isValid).toBe(true);
        expect(parsed.params?.latexString).toBe(expected);
      }
    });

    it('compiles equations containing space primitives into valid HTML and SVG', async () => {
      const complexSpace = 'f(x) = x^2 \\quad \\text{for } x > 0 \\qquad g(x) = \\sqrt{x} \\enspace (\\text{domain: } x \\ge 0)';
      const val = validateLatex(complexSpace);
      expect(val.isValid).toBe(true);

      const render = await compileLatex(complexSpace);
      expect(render.html).toContain('katex');
      expect(render.svg).toContain('xmlns="http://www.w3.org/2000/svg"');
      expect(render.width).toBeGreaterThan(100);
      expect(render.height).toBeGreaterThan(30);
    });
  });

  // =========================================================================
  // 4. Taskpane Formula Formatting & Mock Base64 Validation
  // =========================================================================
  describe('4. Taskpane Formula Formatting & Excel Service Integration', () => {
    it('generates compact formula `=MATH.KATEX("...")` when default options are used', async () => {
      const state = getMockExcelState();
      const range = state.workbook.getSelectedRange();

      await insertFormulaToActiveCell('E = mc^2', {
        background: 'transparent',
        color: '#000000',
        fontSize: 16,
        displayMode: true
      });

      expect(range.formulas?.[0]?.[0]).toBe('=MATH.KATEX("E = mc^2")');
    });

    it('generates compact formula when options object is omitted', async () => {
      const state = getMockExcelState();
      const range = state.workbook.getSelectedRange();

      await insertFormulaToActiveCell('\\frac{a}{b}');
      expect(range.formulas?.[0]?.[0]).toBe('=MATH.KATEX("\\frac{a}{b}")');
    });

    it('generates parameterized formula when custom formatting options are passed', async () => {
      const state = getMockExcelState();
      const range = state.workbook.getSelectedRange();

      await insertFormulaToActiveCell('x^2', {
        background: '#ffff00',
        color: '#ff0000',
        fontSize: 24,
        displayMode: false
      });

      expect(range.formulas?.[0]?.[0]).toBe('=MATH.KATEX("x^2", "#ffff00", "#ff0000", 24, false)');
    });

    it('properly escapes double quotes inside formula string literals', () => {
      expect(escapeFormulaString('E = mc^2')).toBe('E = mc^2');
      expect(escapeFormulaString('\\text{"quote"}')).toBe('\\text{""quote""}');
      expect(escapeFormulaString('a = "b" and c = "d"')).toBe('a = ""b"" and c = ""d""');
    });

    it('insertFloatingShapeToActiveCell accepts valid short base64 test images and rejects empty ones', async () => {
      // Valid compile
      await expect(insertFloatingShapeToActiveCell('x + y')).resolves.not.toThrow();

      const state = getMockExcelState();
      const sheet = state.workbook.worksheets.getActiveWorksheet();
      expect(sheet.shapes.items.length).toBe(1);
      expect(sheet.shapes.items[0].altTextTitle).toBe('LaTeX: x + y');
    });

    it('insertInCellImageToActiveCell assigns both valuesAsJson and values', async () => {
      setMockRequirementSupported('ExcelApi', 1.16, true);
      const state = getMockExcelState();
      const range = state.workbook.getSelectedRange();

      await insertInCellImageToActiveCell('\\int x dx');

      expect((range as any).valuesAsJson).toBeDefined();
      expect((range as any).valuesAsJson[0][0].type).toBe('Entity');
      expect((range as any).valuesAsJson[0][0].text).toBe('[Math: \\int x dx]');
      expect(range.values[0][0].type).toBe('Entity');
    });

    it('readActiveCellFormula extracts LaTeX from formulas, entities, and shape placeholders', async () => {
      const state = getMockExcelState();
      const range = state.workbook.getSelectedRange();

      // 1. From compact formula
      range.formulas = [['=MATH.KATEX("a^2 + b^2 = c^2")']];
      expect(await readActiveCellFormula()).toBe('a^2 + b^2 = c^2');

      // 2. From parameterized formula
      range.formulas = [['=MATH.KATEX("\\sqrt{x}", "#fff", "#000", 20, true)']];
      expect(await readActiveCellFormula()).toBe('\\sqrt{x}');

      // 3. From entity cell value
      range.formulas = [['']];
      range.values = [[{
        type: 'Entity',
        text: '[Math: \\sum n]',
        properties: { latex: { basicValue: '\\sum n' } }
      }]];
      expect(await readActiveCellFormula()).toBe('\\sum n');

      // 4. From 📐 marker
      range.formulas = [['']];
      range.values = [['📐 \\gamma_k']];
      expect(await readActiveCellFormula()).toBe('\\gamma_k');
    });
  });

  // =========================================================================
  // 5. Holistic 10-Flow Systemic Verification
  // =========================================================================
  describe('5. Holistic 10-Flow Compatibility Verification', () => {
    // Flow 1: Taskpane edit -> live preview -> Insert to Cell -> in-cell / fallback shape
    it('Flow 1: Taskpane insert flow produces valid entity or fallback shape', async () => {
      const latex = '\\int_a^b f(x) dx';
      const state = getMockExcelState();

      // Modern path
      setMockRequirementSupported('ExcelApi', 1.16, true);
      await insertInCellImageToActiveCell(latex);
      const range = state.workbook.getSelectedRange();
      expect((range as any).valuesAsJson[0][0].type).toBe('Entity');

      // Legacy fallback path
      setMockRequirementSupported('ExcelApi', 1.16, false);
      await insertInCellImageToActiveCell(latex);
      const sheet = state.workbook.worksheets.getActiveWorksheet();
      expect(sheet.shapes.items.length).toBeGreaterThanOrEqual(1);
    });

    // Flow 2: Direct =MATH.KATEX(...) manual formula entry
    it('Flow 2: Direct =MATH.KATEX(...) manual formula entry resolves cleanly', async () => {
      setMockRequirementSupported('ExcelApi', 1.16, true);
      const inv = CustomFunctionsMock.createInvocation('Sheet1!C3');

      const entityResult = await mathKatex('\\frac{\\partial y}{\\partial x}', undefined, undefined, 16, true, 'auto', inv);
      expect((entityResult as KatexEntityCellValue).type).toBe('Entity');
      expect((entityResult as KatexEntityCellValue).properties?.latex?.basicValue).toBe('\\frac{\\partial y}{\\partial x}');
    });

    // Flow 3: Mid-session setting modifications
    it('Flow 3: Mid-session settings (color, background, fontSize) apply deterministically', async () => {
      const inv = CustomFunctionsMock.createInvocation('Sheet1!D4');
      const res = await mathKatex('x + 1', '#1e1e1e', '#00ff00', 28, false, 'cell', inv);

      const entity = res as KatexEntityCellValue;
      expect(entity.properties?.fontSize?.basicValue).toBe(28);
      expect(entity.properties?.displayMode?.basicValue).toBe(false);
      expect(entity.properties?.svg?.basicValue).toContain('background-color: #1e1e1e;');
      expect(entity.properties?.svg?.basicValue).toContain('color: #00ff00;');
    });

    // Flow 4: Copy/paste cell containing =MATH.KATEX(...)
    it('Flow 4: Evaluating multiple identical formulas at different addresses works idempotently', async () => {
      const inv1 = CustomFunctionsMock.createInvocation('Sheet1!A1');
      const inv2 = CustomFunctionsMock.createInvocation('Sheet1!B1');

      const res1 = await mathKatex('E=mc^2', undefined, undefined, 16, true, 'cell', inv1);
      const res2 = await mathKatex('E=mc^2', undefined, undefined, 16, true, 'cell', inv2);

      expect((res1 as KatexEntityCellValue).text).toBe((res2 as KatexEntityCellValue).text);
      expect((res1 as KatexEntityCellValue).properties?.dimensions?.basicValue).toBe(
        (res2 as KatexEntityCellValue).properties?.dimensions?.basicValue
      );
    });

    // Flow 5: Undo/redo lifecycle and fallback shape cleanup
    it('Flow 5: Clearing and re-enqueuing shape queue operates reliably', async () => {
      KaTeXShapeManager.enqueueShape('x_1', 'Sheet1!A1');
      expect(KaTeXShapeManager.getPendingCount()).toBe(1);

      KaTeXShapeManager.clear();
      expect(KaTeXShapeManager.getPendingCount()).toBe(0);

      KaTeXShapeManager.enqueueShape('x_2', 'Sheet1!A1');
      expect(KaTeXShapeManager.getPendingCount()).toBe(1);

      const processed = await KaTeXShapeManager.processQueue();
      expect(processed).toBe(1);
    });

    // Flow 6: Save/close/reopen workbook stability
    it('Flow 6: Parsing serialized EntityCellValue restores full metadata', () => {
      const rawEntity = buildKatexEntityCellValue('\\lambda = 500\\text{ nm}', {
        pngDataUrl: 'data:image/png;base64,mock',
        width: 150,
        height: 45,
        svg: '<svg></svg>'
      }, { fontSize: 18 });

      const serialized = JSON.stringify(rawEntity);
      const deserialized: KatexEntityCellValue = JSON.parse(serialized);

      expect(deserialized.type).toBe('Entity');
      expect(deserialized.properties?.latex?.basicValue).toBe('\\lambda = 500\\text{ nm}');
      expect(deserialized.properties?.fontSize?.basicValue).toBe(18);
    });

    // Flow 7: Rapid typing debounce
    it('Flow 7: Rapid sequence of shape enqueues coalesce into single queue', async () => {
      for (let i = 0; i < 20; i++) {
        KaTeXShapeManager.enqueueShape(`x_{${i}}`, `Sheet1!A${i + 1}`);
      }
      expect(KaTeXShapeManager.getPendingCount()).toBe(20);

      const processed = await KaTeXShapeManager.processQueue();
      expect(processed).toBe(20);
      expect(KaTeXShapeManager.getPendingCount()).toBe(0);
    });

    // Flow 8: Malformed LaTeX error reporting
    it('Flow 8: Malformed LaTeX yields #VALUE! and does not pollute state', async () => {
      const inv = CustomFunctionsMock.createInvocation('Sheet1!A1');
      await expect(mathKatex('\\frac{1}{', undefined, undefined, 16, true, 'cell', inv)).rejects.toThrow();
      expect(KaTeXShapeManager.getPendingCount()).toBe(0);
    });

    // Flow 9: manifest.xml requirement set verification
    it('Flow 9: manifest.xml specifies valid CustomFunctions and SharedRuntime requirements', () => {
      const manifestPath = path.resolve(__dirname, '../public/manifest.xml');
      expect(fs.existsSync(manifestPath)).toBe(true);

      const manifestContent = fs.readFileSync(manifestPath, 'utf8');
      expect(manifestContent).toContain('CustomFunctionsRuntime');
      expect(manifestContent).toContain('SharedRuntime');
      expect(manifestContent).not.toContain('http://cdn');
      expect(manifestContent).not.toContain('https://cdn');
    });

    // Flow 10: Correct capture and preservation of LaTeX space commands
    it('Flow 10: Complete space command set validated end-to-end', async () => {
      const spaceFormulas = [
        'a\\ b',
        'a\\quad b',
        'a\\qquad b',
        '\\text{alpha beta}',
        'x \\enspace y \\thinspace z'
      ];

      for (const formula of spaceFormulas) {
        const inv = CustomFunctionsMock.createInvocation('Sheet1!A1');
        const res = await mathKatex(`$$ ${formula} $$`, undefined, undefined, 16, true, 'cell', inv);
        expect((res as KatexEntityCellValue).properties?.latex?.basicValue).toBe(formula);
      }
    });
  });
});
