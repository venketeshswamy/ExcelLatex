import { describe, it, expect, beforeEach } from 'vitest';
import { compileLatex, compileLatexToHtml, validateLatex } from '../../src/core/katexEngine';
import { generateStandaloneSvg } from '../../src/core/imageRasterizer';
import {
  mathKatexFunction,
  KaTeXShapeManager
} from '../helpers/testHelpers';
import {
  resetOfficeMock,
  ExcelMock
} from '../../src/mocks/officeMock';
import {
  resetCustomFunctionsMock,
  CustomFunctionsMock
} from '../../src/mocks/customFunctionsMock';

describe('Tier 2: Boundary, Edge & Corner Cases Suite', () => {
  beforeEach(() => {
    resetOfficeMock();
    resetCustomFunctionsMock();
    KaTeXShapeManager.clear();
  });

  describe('2.1 Empty, Whitespace & Comment Inputs', () => {
    it('handles empty string without unhandled exceptions', () => {
      const validation = validateLatex('');
      expect(validation.isValid).toBe(false);
      expect(validation.error).toBeDefined();

      const html = compileLatexToHtml('');
      expect(html).toBe('');
    });

    it('handles pure whitespace and tabs/newlines', () => {
      const whitespaceString = '   \t\r\n   ';
      const validation = validateLatex(whitespaceString);
      expect(validation.isValid).toBe(false);
    });

    it('handles LaTeX comment-only input', () => {
      const commentInput = '% This is a pure latex comment';
      const validation = validateLatex(commentInput);
      expect(typeof validation.isValid).toBe('boolean');
    });

    it('rejects empty input in mathKatexFunction with CustomFunctions.Error', async () => {
      await expect(mathKatexFunction('')).rejects.toThrow();
      await expect(mathKatexFunction('   ')).rejects.toThrow();
    });
  });

  describe('2.2 Malformed & Broken LaTeX Syntax', () => {
    it('handles missing fraction denominator (\\frac{1})', () => {
      const res = validateLatex('\\frac{1}');
      expect(res.isValid).toBe(false);
      expect(res.error).toContain('KaTeX parse error');

      const safeHtml = compileLatexToHtml('\\frac{1}', { throwOnError: false });
      expect(safeHtml).toBeDefined();
      expect(safeHtml.length).toBeGreaterThan(0);
    });

    it('handles mismatched environments (\\begin{matrix} ... \\end{pmatrix})', () => {
      const res = validateLatex('\\begin{matrix} 1 & 2 \\end{pmatrix}');
      expect(res.isValid).toBe(false);
      expect(res.error).toBeDefined();
    });

    it('handles unclosed braces and brackets', () => {
      const unclosedBrace = '\\int_{0}^{\\infty x dx';
      expect(validateLatex(unclosedBrace).isValid).toBe(false);

      const unclosedBracket = '\\sqrt[3';
      expect(validateLatex(unclosedBracket).isValid).toBe(false);
    });

    it('handles unknown control sequences gracefully', () => {
      const unknown = '\\nonExistentMacroName{foo}';
      const res = validateLatex(unknown);
      expect(res.isValid).toBe(false);
      expect(res.error).toContain('Undefined control sequence');
    });
  });

  describe('2.3 Deep Structural Nesting & High Complexity', () => {
    it('handles 20-level deeply nested fractions', () => {
      let nestedFrac = 'x';
      for (let i = 0; i < 20; i++) {
        nestedFrac = `\\frac{1}{${nestedFrac}}`;
      }

      const validation = validateLatex(nestedFrac);
      expect(validation.isValid).toBe(true);
      expect(validation.complexity?.score).toBe('Advanced');
      expect(validation.complexity?.depth).toBeGreaterThanOrEqual(20);

      const html = compileLatexToHtml(nestedFrac);
      expect(html).toContain('katex');
    });

    it('handles 10-level nested radicals (\\sqrt{\\sqrt{...}})', () => {
      let nestedSqrt = 'x';
      for (let i = 0; i < 10; i++) {
        nestedSqrt = `\\sqrt{${nestedSqrt}}`;
      }

      const validation = validateLatex(nestedSqrt);
      expect(validation.isValid).toBe(true);
      expect(validation.complexity?.depth).toBeGreaterThanOrEqual(10);
    });

    it('handles 10x10 large matrix environments', () => {
      const rows = [];
      for (let r = 0; r < 10; r++) {
        const cols = [];
        for (let c = 0; c < 10; c++) {
          cols.push(`M_{${r},${c}}`);
        }
        rows.push(cols.join(' & '));
      }
      const matrixLatex = `\\begin{pmatrix} ${rows.join(' \\\\ ')} \\end{pmatrix}`;

      const validation = validateLatex(matrixLatex);
      expect(validation.isValid).toBe(true);
      expect(validation.complexity?.score).toBe('Advanced');
    });
  });

  describe('2.4 Extreme Typography, Colors & Font Sizes', () => {
    it('handles fontSize: 0 or negative font size with fallback', async () => {
      const res1 = await compileLatex('a + b = c', { fontSize: 0 });
      expect(res1.width).toBeGreaterThan(0);
      expect(res1.height).toBeGreaterThan(0);

      const res2 = await compileLatex('a + b = c', { fontSize: -10 });
      expect(res2.width).toBeGreaterThan(0);
    });

    it('handles extremely large fontSize (500pt)', async () => {
      const res = await compileLatex('E = mc^2', { fontSize: 500 });
      expect(res.width).toBeGreaterThan(500);
      expect(res.height).toBeGreaterThan(200);
    });

    it('handles RGBA, transparent, and non-standard color formats', async () => {
      const svg = generateStandaloneSvg(
        '<span>test</span>',
        100,
        30,
        'rgba(255, 255, 0, 0.3)',
        'rgba(0, 0, 128, 1)',
        14
      );
      expect(svg).toContain('background-color: rgba(255, 255, 0, 0.3)');
      expect(svg).toContain('color: rgba(0, 0, 128, 1)');
    });
  });

  describe('2.5 Long Payloads & Extreme String Boundaries', () => {
    it('handles long algebraic polynomial string (50+ terms)', () => {
      const terms = [];
      for (let i = 0; i < 50; i++) {
        terms.push(`a_{${i}} x^{${i}}`);
      }
      const longPoly = `P(x) = ${terms.join(' + ')}`;

      const validation = validateLatex(longPoly);
      expect(validation.isValid).toBe(true);
      expect(validation.complexity?.tokens).toBeGreaterThan(100);
    });

    it('handles multi-line equation alignment environments', () => {
      const alignLatex = `\\begin{aligned}
        \\nabla \\times \\mathbf{E} &= -\\frac{\\partial \\mathbf{B}}{\\partial t} \\\\
        \\nabla \\times \\mathbf{B} &= \\mu_0 \\mathbf{J} + \\mu_0 \\varepsilon_0 \\frac{\\partial \\mathbf{E}}{\\partial t} \\\\
        \\nabla \\cdot \\mathbf{E} &= \\frac{\\rho}{\\varepsilon_0} \\\\
        \\nabla \\cdot \\mathbf{B} &= 0
      \\end{aligned}`;

      const validation = validateLatex(alignLatex);
      expect(validation.isValid).toBe(true);
      expect(validation.complexity?.score).toBe('Advanced');
    });
  });

  describe('2.6 Custom Function Invocation & Special Address Anchors', () => {
    it('handles special worksheet names with spaces, apostrophes and dollar signs', async () => {
      const inv = CustomFunctionsMock.createInvocation("'Q3 Financials'!$AA$100");
      const entity = (await mathKatexFunction(
        '\\text{NPV} = \\sum \\frac{C_t}{(1+r)^t}',
        undefined,
        undefined,
        undefined,
        undefined,
        'cell',
        inv
      )) as ExcelMock.EntityCellValue;

      expect(entity.type).toBe('Entity');
    });

    it('handles undefined invocation object gracefully', async () => {
      const entity = (await mathKatexFunction('x = 1', undefined, undefined, undefined, undefined, 'cell', undefined)) as ExcelMock.EntityCellValue;
      expect(entity.type).toBe('Entity');
    });
  });
});
