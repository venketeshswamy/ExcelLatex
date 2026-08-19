import { describe, it, expect, beforeEach } from 'vitest';
import {
  compileLatex,
  compileLatexToHtml,
  compileLatexToMathML,
  validateLatex,
  calculateComplexity
} from '../../src/core/katexEngine';
import {
  rasterizeLatex,
  generateStandaloneSvg,
  getEquationCacheStats,
  clearEquationCache
} from '../../src/core/imageRasterizer';
import { macroRegistry } from '../../src/core/macros';
import { LRUCache } from '../../src/core/lruCache';

describe('Unit Test: Core KaTeX Local Math Engine', () => {
  beforeEach(() => {
    clearEquationCache();
  });

  describe('calculateComplexity()', () => {
    it('returns Simple score for empty or minimal expressions', () => {
      const emptyResult = calculateComplexity('');
      expect(emptyResult.score).toBe('Simple');
      expect(emptyResult.depth).toBe(0);
      expect(emptyResult.tokens).toBe(0);

      const simpleResult = calculateComplexity('x + y = z');
      expect(simpleResult.score).toBe('Simple');
      expect(simpleResult.depth).toBe(0);
    });

    it('identifies Moderate complexity with braces and basic fractions', () => {
      const modResult = calculateComplexity('\\frac{a+b}{c-d}');
      expect(['Moderate', 'Complex']).toContain(modResult.score);
      expect(modResult.depth).toBeGreaterThanOrEqual(1);
    });

    it('identifies Complex / Advanced score for matrices and nested environments', () => {
      const matrixLatex = `\\begin{pmatrix} a & b \\\\ c & d \\end{pmatrix}`;
      const res = calculateComplexity(matrixLatex);
      expect(['Complex', 'Advanced']).toContain(res.score);
      expect(res.depth).toBeGreaterThanOrEqual(2);
    });

    it('identifies Advanced score for deeply nested fractions and integrals', () => {
      const advancedLatex = `\\int_{-\\infty}^{+\\infty} \\frac{\\sqrt{1 + \\frac{x^2}{n}}}{1 + x^4} \\mathrm{d}x = \\frac{\\pi}{\\sqrt{2}}`;
      const res = calculateComplexity(advancedLatex);
      expect(res.score).toBe('Advanced');
      expect(res.tokens).toBeGreaterThan(15);
    });
  });

  describe('validateLatex()', () => {
    it('returns isValid: false with error message for empty input', () => {
      const res = validateLatex('');
      expect(res.isValid).toBe(false);
      expect(res.error).toBeDefined();
    });

    it('returns isValid: true and complexity metadata for valid LaTeX', () => {
      const res = validateLatex('E = mc^2');
      expect(res.isValid).toBe(true);
      expect(res.error).toBeUndefined();
      expect(res.complexity).toBeDefined();
      expect(['Simple', 'Moderate']).toContain(res.complexity?.score);
    });

    it('returns isValid: false with diagnostic message for broken LaTeX syntax', () => {
      const res = validateLatex('\\frac{1}');
      expect(res.isValid).toBe(false);
      expect(res.error).toContain('KaTeX parse error');
    });

    it('validates formulas containing standard macros', () => {
      const res = validateLatex('\\R \\times \\C \\implies \\grad f');
      expect(res.isValid).toBe(true);
    });
  });

  describe('compileLatexToHtml() & compileLatexToMathML()', () => {
    it('compiles basic algebra to HTML containing katex classes', () => {
      const html = compileLatexToHtml('a^2 + b^2 = c^2');
      expect(html).toContain('class="katex"');
      expect(html).toContain('katex-html');
    });

    it('supports displayMode true vs false', () => {
      const displayHtml = compileLatexToHtml('\\sum_{i=1}^n i', { displayMode: true });
      expect(displayHtml).toContain('katex-display');

      const inlineHtml = compileLatexToHtml('\\sum_{i=1}^n i', { displayMode: false });
      expect(inlineHtml).not.toContain('katex-display');
    });

    it('compiles to MathML when requested', () => {
      const mathml = compileLatexToMathML('\\sqrt{x}');
      expect(mathml).toContain('<math');
      expect(mathml).toContain('<msqrt>');
    });

    it('renders safe fallback markup when throwOnError is false', () => {
      const html = compileLatexToHtml('\\invalidmacro{x}', { throwOnError: false });
      expect(html).toBeDefined();
      expect(html.length).toBeGreaterThan(0);
    });

    it('throws error when throwOnError is true', () => {
      expect(() => {
        compileLatexToHtml('\\invalidmacro{x}', { throwOnError: true });
      }).toThrow();
    });
  });

  describe('rasterizeLatex() & Standalone SVG Generation', () => {
    it('generates standalone SVG with embedded foreignObject and styles', () => {
      const svg = generateStandaloneSvg('<span>KaTeX Markup</span>', 200, 60, '#ffffff', '#107c41', 16);
      expect(svg).toContain('<svg xmlns="http://www.w3.org/2000/svg"');
      expect(svg).toContain('width="200" height="60"');
      expect(svg).toContain('<foreignObject');
      expect(svg).toContain('background-color: #ffffff');
      expect(svg).toContain('color: #107c41');
      expect(svg).toContain('KaTeX Markup');
    });

    it('compiles LaTeX into full RenderResult with PNG and SVG data URLs', async () => {
      const result = await compileLatex('\\int_0^1 x^2 \\mathrm{d}x = \\frac{1}{3}', {
        background: '#ffffff',
        color: '#000000',
        fontSize: 18,
        displayMode: true
      });

      expect(result.html).toContain('class="katex"');
      expect(result.svg).toContain('<svg');
      expect(result.pngDataUrl).toBeDefined();
      expect(result.pngDataUrl.startsWith('data:image/')).toBe(true);
      expect(result.width).toBeGreaterThan(0);
      expect(result.height).toBeGreaterThan(0);
      expect(result.aspectRatio).toBeGreaterThan(0);
    });

    it('reuses cached results for identical inputs and options', async () => {
      const formula = '\\nabla \\times \\mathbf{B} = \\mu_0 \\mathbf{J}';
      const options = { fontSize: 16, color: '#201f1e' };

      const statsBefore = getEquationCacheStats();
      expect(statsBefore.size).toBe(0);

      const res1 = await rasterizeLatex(formula, options);
      const statsAfterFirst = getEquationCacheStats();
      expect(statsAfterFirst.size).toBe(1);
      expect(statsAfterFirst.misses).toBe(1);
      expect(statsAfterFirst.hits).toBe(0);

      const res2 = await rasterizeLatex(formula, options);
      const statsAfterSecond = getEquationCacheStats();
      expect(statsAfterSecond.size).toBe(1);
      expect(statsAfterSecond.hits).toBe(1);
      expect(res1).toBe(res2); // Reference equality from LRU cache
    });
  });

  describe('Macros Registry & Custom Expansion', () => {
    it('contains standard mathematical blackboard bold sets', () => {
      expect(macroRegistry.get('\\R')).toBeDefined();
      expect(macroRegistry.get('\\N')).toBeDefined();
      expect(macroRegistry.get('\\Z')).toBeDefined();
      expect(macroRegistry.get('\\C')).toBeDefined();
      expect(macroRegistry.get('\\Q')).toBeDefined();
    });

    it('allows registering and unregistering custom macros', () => {
      macroRegistry.register('\\myop', '\\operatorname{customOp}');
      expect(macroRegistry.get('\\myop')).toBeDefined();

      const html = compileLatexToHtml('\\myop(x)');
      expect(html).toContain('customOp');

      macroRegistry.unregister('\\myop');
      expect(macroRegistry.get('\\myop')).toBeUndefined();
    });
  });

  describe('LRUCache Engine Implementation', () => {
    it('evicts least recently used items when maxSize is exceeded', () => {
      const cache = new LRUCache<string, number>({ maxSize: 3 });
      cache.set('a', 1);
      cache.set('b', 2);
      cache.set('c', 3);

      expect(cache.get('a')).toBe(1); // 'a' accessed, 'b' is now oldest
      cache.set('d', 4); // should evict 'b'

      expect(cache.has('b')).toBe(false);
      expect(cache.has('a')).toBe(true);
      expect(cache.has('c')).toBe(true);
      expect(cache.has('d')).toBe(true);
      expect(cache.size()).toBe(3);
    });
  });
});
