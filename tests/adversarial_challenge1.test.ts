import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import {
  compileLatex,
  compileLatexToHtml,
  compileLatexToMathML,
  validateLatex,
  calculateComplexity
} from '../src/core/katexEngine';
import {
  rasterizeLatex,
  generateStandaloneSvg,
  svgToPngDataUrl,
  clearEquationCache,
  getEquationCacheStats
} from '../src/core/imageRasterizer';
import { LRUCache, createRenderCacheKey } from '../src/core/lruCache';
import { macroRegistry } from '../src/core/macros';
import {
  mathKatex,
  createCustomFunctionsError
} from '../src/customfunctions/mathKatex';
import {
  parseMathKatexParams,
  sanitizeBackground,
  sanitizeTextColor,
  sanitizeColor,
  sanitizeFontSize,
  sanitizeDisplayMode,
  sanitizeOutputMethod,
  isDarkColor,
  getDefaultTextColor,
  extractCellAddress,
  stripLatexDelimiters
} from '../src/customfunctions/parameterParser';
import { KatexEntityCellValue } from '../src/customfunctions/entityCellBuilder';
import { KaTeXShapeManager } from '../src/customfunctions/shapeManager';
import {
  resetOfficeMock,
  setMockRequirementSupported,
  getMockExcelState
} from '../src/mocks/officeMock';
import {
  resetCustomFunctionsMock,
  CustomFunctionsMock
} from '../src/mocks/customFunctionsMock';

describe('Adversarial Challenge 1: Core Rasterizer (R1) & Parameter Parser (R2) Stress Suite', () => {
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
  // 1. Extreme LaTeX Nesting & Recursive Structural Stress (R1)
  // =========================================================================
  describe('1. Extreme LaTeX Nesting & Recursive Structural Stress', () => {
    it('survives 10-level, 50-level, 100-level, and 150-level nested fractions', async () => {
      const levels = [10, 50, 100, 150];
      for (const lvl of levels) {
        let frac = '1';
        for (let i = 0; i < lvl; i++) {
          frac = `\\frac{1}{${frac} + ${i}}`;
        }
        const t0 = performance.now();
        const validation = validateLatex(frac);
        const t1 = performance.now();

        expect(validation.isValid).toBe(true);
        expect(validation.complexity?.depth).toBeGreaterThanOrEqual(Math.min(lvl, 4));
        expect(validation.complexity?.score).toBe('Advanced');
        expect(t1 - t0).toBeLessThan(2000);

        // Verify rasterization DOM-walk succeeds without call stack overflow
        const result = await compileLatex(frac);
        expect(result.html).toContain('katex');
        expect(result.pngDataUrl.startsWith('data:image/png')).toBe(true);
        expect(result.width).toBeGreaterThan(0);
        expect(result.height).toBeGreaterThan(0);
      }
    });

    it('survives 150-level nested square roots without stack overflow', async () => {
      let sqrt = 'x';
      for (let i = 0; i < 150; i++) {
        sqrt = `\\sqrt{${sqrt} + 1}`;
      }
      const validation = validateLatex(sqrt);
      expect(typeof validation.isValid).toBe('boolean');
      const html = compileLatexToHtml(sqrt, { throwOnError: false });
      expect(html).toContain('katex');

      const result = await compileLatex(sqrt);
      expect(result.pngDataUrl.startsWith('data:image/png')).toBe(true);
    });

    it('survives 80-level nested superscripts and subscripts', async () => {
      let superSub = 'x';
      for (let i = 0; i < 40; i++) {
        superSub = `{${superSub}}_{i_{${i}}}^{j^{${i}}}`;
      }
      const validation = validateLatex(superSub);
      expect(validation.isValid).toBe(true);
      const res = await compileLatex(superSub);
      expect(res.pngDataUrl.startsWith('data:image/png')).toBe(true);
    });

    it('survives deeply nested alternating structures (sqrt -> frac -> sum -> sqrt)', async () => {
      let expr = 'x';
      for (let i = 0; i < 15; i++) {
        expr = `\\sqrt{\\frac{\\sum_{k=1}^n ${expr}}{\\sqrt{k^2 + 1}}}`;
      }
      const validation = validateLatex(expr);
      expect(validation.isValid).toBe(true);
      const res = await compileLatex(expr);
      expect(res.pngDataUrl).toBeDefined();
      expect(res.pngDataUrl.startsWith('data:image/png')).toBe(true);
    });
  });

  // =========================================================================
  // 2. Complex & Adversarial Matrices (R1)
  // =========================================================================
  describe('2. Giant & Adversarial Matrices', () => {
    it('handles a 20x20 dense symbolic matrix compilation and rasterization', async () => {
      const rows: string[] = [];
      for (let r = 0; r < 20; r++) {
        const cols: string[] = [];
        for (let c = 0; c < 20; c++) {
          cols.push(`M_{${r + 1},${c + 1}}`);
        }
        rows.push(cols.join(' & '));
      }
      const matrixLatex = `\\begin{pmatrix} ${rows.join(' \\\\ ')} \\end{pmatrix}`;

      const t0 = performance.now();
      const validation = validateLatex(matrixLatex);
      const t1 = performance.now();

      expect(validation.isValid).toBe(true);
      expect(t1 - t0).toBeLessThan(2500);

      const render = await compileLatex(matrixLatex);
      expect(render.width).toBeGreaterThan(100);
      expect(render.height).toBeGreaterThan(100);
      expect(render.pngDataUrl.startsWith('data:image/png')).toBe(true);
    });

    it('handles jagged / asymmetric matrix rows gracefully', () => {
      const jagged = `\\begin{bmatrix} 1 & 2 & 3 \\\\ 4 & 5 \\\\ 6 & 7 & 8 & 9 & 10 \\end{bmatrix}`;
      const validation = validateLatex(jagged);
      expect(validation.isValid).toBe(true);
      const html = compileLatexToHtml(jagged);
      expect(html).toContain('katex');
    });

    it('handles matrix containing complex nested formulas in elements', async () => {
      const complexMatrix = `\\begin{pmatrix}
        \\int_0^\\infty e^{-x^2} dx & \\frac{\\partial^2 u}{\\partial t^2} \\\\
        \\sum_{n=1}^\\infty \\frac{1}{n^2} & \\sqrt{\\frac{a + b}{c - d}}
      \\end{pmatrix}`;

      const validation = validateLatex(complexMatrix);
      expect(validation.isValid).toBe(true);
      expect(validation.complexity?.score).toBe('Advanced');

      const res = await compileLatex(complexMatrix);
      expect(res.html).toContain('katex');
      expect(res.pngDataUrl.startsWith('data:image/png')).toBe(true);
    });

    it('supports diverse matrix environments: pmatrix, bmatrix, vmatrix, Vmatrix, Bmatrix, matrix, cases, aligned', () => {
      const envs = ['pmatrix', 'bmatrix', 'vmatrix', 'Vmatrix', 'Bmatrix', 'matrix', 'cases', 'aligned'];
      for (const env of envs) {
        const formula = `\\begin{${env}} a & b \\\\ c & d \\end{${env}}`;
        const validation = validateLatex(formula);
        expect(validation.isValid).toBe(true);
      }
    });
  });

  // =========================================================================
  // 3. Multi-Line Equations & Alignment Environments (R1)
  // =========================================================================
  describe('3. Multi-Line Equations & Alignment Environments', () => {
    it('handles aligned equations with multiple ampersands and linebreaks', async () => {
      const aligned = `\\begin{aligned}
        f(x) &= (x+a)(x+b) \\\\
             &= x^2 + (a+b)x + ab \\\\
             &= x^2 + 2ax + a^2
      \\end{aligned}`;
      const val = validateLatex(aligned);
      expect(val.isValid).toBe(true);
      const res = await compileLatex(aligned);
      expect(res.pngDataUrl.startsWith('data:image/png')).toBe(true);
      expect(res.height).toBeGreaterThan(40);
    });

    it('handles cases environment with piecewise mathematical conditions', async () => {
      const piecewise = `f(n) = \\begin{cases} 
        \\frac{n}{2} & \\text{if } n \\equiv 0 \\pmod{2} \\\\ 
        3n + 1 & \\text{if } n \\equiv 1 \\pmod{2} 
      \\end{cases}`;
      const val = validateLatex(piecewise);
      expect(val.isValid).toBe(true);
      const res = await compileLatex(piecewise);
      expect(res.pngDataUrl.startsWith('data:image/png')).toBe(true);
    });

    it('handles gather environment for centered multi-line equations', async () => {
      const gather = `\\begin{gathered}
        2x - 5y =  8 \\\\
        3x + 9y = -12
      \\end{gathered}`;
      const val = validateLatex(gather);
      expect(val.isValid).toBe(true);
      const res = await compileLatex(gather);
      expect(res.pngDataUrl.startsWith('data:image/png')).toBe(true);
    });
  });

  // =========================================================================
  // 4. Special Characters, Escaped Symbols & Multilingual Text (R1)
  // =========================================================================
  describe('4. Special Characters, Escaped Symbols & Multilingual Text', () => {
    it('handles escaped special TeX characters: %, &, _, #, $, {, }', async () => {
      const specials = '100\\% \\& \\#1 \\_ \\$50 \\{x \\mid x > 0\\}';
      const val = validateLatex(specials);
      expect(val.isValid).toBe(true);
      const res = await compileLatex(specials);
      expect(res.pngDataUrl.startsWith('data:image/png')).toBe(true);
    });

    it('handles accents, hats, vectors, and diacritics', async () => {
      const accents = '\\hat{a} + \\bar{b} + \\tilde{c} + \\vec{v} + \\dot{x} + \\ddot{y} + \\grave{z} + \\acute{w} + \\check{u} + \\breve{t}';
      const val = validateLatex(accents);
      expect(val.isValid).toBe(true);
      const res = await compileLatex(accents);
      expect(res.pngDataUrl.startsWith('data:image/png')).toBe(true);
    });

    it('handles big operators: int, iint, iiint, oint, sum, prod, coprod, bigotimes, bigoplus, bigcup, bigcap', async () => {
      const ops = '\\iint_D f(x,y)\\,dxdy + \\oint_C \\mathbf{F} \\cdot d\\mathbf{r} + \\prod_{i=1}^n i! + \\coprod_{j=1}^m A_j + \\bigotimes_{k=1}^p V_k + \\bigoplus W_l';
      const val = validateLatex(ops);
      expect(val.isValid).toBe(true);
      const res = await compileLatex(ops);
      expect(res.pngDataUrl.startsWith('data:image/png')).toBe(true);
    });

    it('handles complete Greek alphabet in upper and lowercase', async () => {
      const greek = '\\alpha \\beta \\gamma \\delta \\epsilon \\zeta \\eta \\theta \\iota \\kappa \\lambda \\mu \\nu \\xi \\pi \\rho \\sigma \\tau \\upsilon \\phi \\chi \\psi \\omega \\Gamma \\Delta \\Theta \\Lambda \\Xi \\Pi \\Sigma \\Phi \\Psi \\Omega';
      const val = validateLatex(greek);
      expect(val.isValid).toBe(true);
      const res = await compileLatex(greek);
      expect(res.html).toContain('katex');
      expect(res.width).toBeGreaterThan(0);
    });

    it('handles Hebrew and math set symbols: R, C, H, N, Z, Q, aleph, beth, gimel, daleth', () => {
      const symbols = '\\mathbb{R} \\mathbb{C} \\mathbb{H} \\mathbb{N} \\mathbb{Z} \\mathbb{Q} \\aleph \\beth \\gimel \\daleth';
      const val = validateLatex(symbols);
      expect(val.isValid).toBe(true);
    });

    it('handles multilingual text in \\text{}', async () => {
      const multilingual = '\\text{English} + \\text{Español: niño} + \\text{Français: été} + \\text{Deutsch: Größe} + \\text{日本語: 数式} + \\text{中文: 方程} + \\text{Русский: интеграл}';
      const val = validateLatex(multilingual);
      expect(val.isValid).toBe(true);
      const res = await compileLatex(multilingual);
      expect(res.html).toContain('katex');
      expect(res.pngDataUrl.startsWith('data:image/png')).toBe(true);
    });
  });

  // =========================================================================
  // 5. Zero Font Sizes, Clamping & Extreme Zoom Factors (R1 & R2)
  // =========================================================================
  describe('5. Zero Font Sizes, Clamping & Extreme Zoom Factors', () => {
    it('sanitizes and clamps font size edge cases: 0, negative, NaN, null, string, extremes', () => {
      // 0 or <= 0 -> default 16
      expect(sanitizeFontSize(0)).toBe(16);
      expect(sanitizeFontSize(-1)).toBe(16);
      expect(sanitizeFontSize(-999)).toBe(16);
      expect(sanitizeFontSize(0.0)).toBe(16);

      // Non-numeric / missing -> default 16
      expect(sanitizeFontSize(undefined)).toBe(16);
      expect(sanitizeFontSize(null)).toBe(16);
      expect(sanitizeFontSize('')).toBe(16);
      expect(sanitizeFontSize('   ')).toBe(16);
      expect(sanitizeFontSize('invalid')).toBe(16);
      expect(sanitizeFontSize(NaN)).toBe(16);

      // Clamping bounds [4, 1000]
      expect(sanitizeFontSize(0.5)).toBe(4);
      expect(sanitizeFontSize(1)).toBe(4);
      expect(sanitizeFontSize(2)).toBe(4);
      expect(sanitizeFontSize(3.9)).toBe(4);
      expect(sanitizeFontSize(4)).toBe(4);
      expect(sanitizeFontSize(12)).toBe(12);
      expect(sanitizeFontSize(24)).toBe(24);
      expect(sanitizeFontSize(72)).toBe(72);
      expect(sanitizeFontSize(500)).toBe(500);
      expect(sanitizeFontSize(1000)).toBe(1000);
      expect(sanitizeFontSize(1001)).toBe(1000);
      expect(sanitizeFontSize(50000)).toBe(1000);
      expect(sanitizeFontSize(Infinity)).toBe(1000);
      expect(sanitizeFontSize(-Infinity)).toBe(16);

      // String numeric inputs
      expect(sanitizeFontSize('18')).toBe(18);
      expect(sanitizeFontSize('2')).toBe(4);
      expect(sanitizeFontSize('2000')).toBe(1000);
    });

    it('handles rasterization with extreme font size boundaries (4pt to 1000pt)', async () => {
      const minRes = await compileLatex('x+1', { fontSize: 4 });
      expect(minRes.pngDataUrl.startsWith('data:image/png')).toBe(true);
      expect(minRes.width).toBeGreaterThan(0);
      expect(minRes.height).toBeGreaterThan(0);

      const maxRes = await compileLatex('x+1', { fontSize: 1000 });
      expect(maxRes.pngDataUrl.startsWith('data:image/png')).toBe(true);
      expect(maxRes.width).toBeGreaterThan(minRes.width);
      expect(maxRes.height).toBeGreaterThan(minRes.height);
    });

    it('supports extreme zoom / scaling factors (0.1x to 16x) without crash', async () => {
      const zoomFactors = [0.1, 0.5, 1, 2, 3, 4, 8, 16];
      for (const scale of zoomFactors) {
        const res = await compileLatex('E = mc^2', { scale });
        expect(res.pngDataUrl.startsWith('data:image/png')).toBe(true);
        expect(res.width).toBeGreaterThan(0);
        expect(res.height).toBeGreaterThan(0);
      }
    });
  });

  // =========================================================================
  // 6. Comprehensive Background Parameter Mapping (R2)
  // =========================================================================
  describe('6. Comprehensive Background Parameter Mapping (0, 1, 2, Strings, Hex, Named, Boundaries)', () => {
    it('maps numeric background inputs: 0 -> transparent, 1 -> #ffffff, 2 -> #000000, others -> transparent', () => {
      expect(sanitizeBackground(0)).toBe('transparent');
      expect(sanitizeBackground(1)).toBe('#ffffff');
      expect(sanitizeBackground(2)).toBe('#000000');
      // Non-standard numbers fall back to transparent default
      expect(sanitizeBackground(3)).toBe('transparent');
      expect(sanitizeBackground(-1)).toBe('transparent');
      expect(sanitizeBackground(999)).toBe('transparent');
    });

    it('maps string background codes: "0", "1", "2", "transparent", "white", "black"', () => {
      expect(sanitizeBackground('0')).toBe('transparent');
      expect(sanitizeBackground(' 0 ')).toBe('transparent');
      expect(sanitizeBackground('transparent')).toBe('transparent');
      expect(sanitizeBackground('TRANSPARENT')).toBe('transparent');

      expect(sanitizeBackground('1')).toBe('#ffffff');
      expect(sanitizeBackground(' 1 ')).toBe('#ffffff');
      expect(sanitizeBackground('white')).toBe('#ffffff');
      expect(sanitizeBackground('WHITE')).toBe('#ffffff');

      expect(sanitizeBackground('2')).toBe('#000000');
      expect(sanitizeBackground(' 2 ')).toBe('#000000');
      expect(sanitizeBackground('black')).toBe('#000000');
      expect(sanitizeBackground('BLACK')).toBe('#000000');
    });

    it('preserves valid custom hex, rgb, rgba, hsl, and named colors', () => {
      expect(sanitizeBackground('#123456')).toBe('#123456');
      expect(sanitizeBackground('#FFFFFF')).toBe('#FFFFFF');
      expect(sanitizeBackground('#000')).toBe('#000');
      expect(sanitizeBackground('rgba(255, 0, 0, 0.5)')).toBe('rgba(255, 0, 0, 0.5)');
      expect(sanitizeBackground('hsl(210, 100%, 50%)')).toBe('hsl(210, 100%, 50%)');
      expect(sanitizeBackground('navy')).toBe('navy');
      expect(sanitizeBackground('darkblue')).toBe('darkblue');
      expect(sanitizeBackground('yellow')).toBe('yellow');
    });

    it('safely falls back to default "transparent" for invalid, null, undefined, and malicious inputs', () => {
      expect(sanitizeBackground(undefined)).toBe('transparent');
      expect(sanitizeBackground(null)).toBe('transparent');
      expect(sanitizeBackground('')).toBe('transparent');
      expect(sanitizeBackground('   ')).toBe('transparent');
      expect(sanitizeBackground('\t\n')).toBe('transparent');
      expect(sanitizeBackground('invalid_color_xyz')).toBe('transparent');
      expect(sanitizeBackground('invalid-css-color!')).toBe('transparent');
      expect(sanitizeBackground('"><script>alert(1)</script>')).toBe('transparent');
      expect(sanitizeBackground('#gggggg')).toBe('transparent');
      expect(sanitizeBackground(true as any)).toBe('transparent');
      expect(sanitizeBackground(false as any)).toBe('transparent');
      expect(sanitizeBackground({} as any)).toBe('transparent');
      expect(sanitizeBackground([] as any)).toBe('transparent');
    });
  });

  // =========================================================================
  // 7. Adaptive Text Color Matrix (R2)
  // =========================================================================
  describe('7. Adaptive Text Color Matrix & Explicit Override Verification', () => {
    it('determines dark colors accurately via isDarkColor', () => {
      // Dark backgrounds -> true
      expect(isDarkColor('#000000')).toBe(true);
      expect(isDarkColor('#000')).toBe(true);
      expect(isDarkColor('black')).toBe(true);
      expect(isDarkColor('#111111')).toBe(true);
      expect(isDarkColor('#0f172a')).toBe(true);
      expect(isDarkColor('rgb(10, 10, 10)')).toBe(true);
      expect(isDarkColor('rgba(0, 0, 0, 1)')).toBe(true);
      expect(isDarkColor('navy')).toBe(true);
      expect(isDarkColor('darkblue')).toBe(true);
      expect(isDarkColor('midnightblue')).toBe(true);
      expect(isDarkColor('darkgreen')).toBe(true);
      expect(isDarkColor('darkred')).toBe(true);
      expect(isDarkColor('purple')).toBe(true);
      expect(isDarkColor('maroon')).toBe(true);

      // Light / transparent backgrounds -> false
      expect(isDarkColor('transparent')).toBe(false);
      expect(isDarkColor('currentColor')).toBe(false);
      expect(isDarkColor('#ffffff')).toBe(false);
      expect(isDarkColor('#fff')).toBe(false);
      expect(isDarkColor('white')).toBe(false);
      expect(isDarkColor('#f8fafc')).toBe(false);
      expect(isDarkColor('rgb(255, 255, 255)')).toBe(false);
      expect(isDarkColor('yellow')).toBe(false);
      expect(isDarkColor('cyan')).toBe(false);
    });

    it('computes default text color adaptively: white for dark bg, black for light/transparent bg', () => {
      expect(getDefaultTextColor('transparent')).toBe('#000000');
      expect(getDefaultTextColor('#ffffff')).toBe('#000000');
      expect(getDefaultTextColor('white')).toBe('#000000');
      expect(getDefaultTextColor('#000000')).toBe('#ffffff');
      expect(getDefaultTextColor('black')).toBe('#ffffff');
      expect(getDefaultTextColor('navy')).toBe('#ffffff');
    });

    it('applies adaptive text color in parseMathKatexParams when color is omitted', () => {
      // 0 (Transparent) -> #000000 text
      const p0 = parseMathKatexParams('x^2', 0);
      expect(p0.params?.background).toBe('transparent');
      expect(p0.params?.color).toBe('#000000');

      // 1 (White) -> #000000 text
      const p1 = parseMathKatexParams('x^2', 1);
      expect(p1.params?.background).toBe('#ffffff');
      expect(p1.params?.color).toBe('#000000');

      // 2 (Black) -> #ffffff text (Adaptive)
      const p2 = parseMathKatexParams('x^2', 2);
      expect(p2.params?.background).toBe('#000000');
      expect(p2.params?.color).toBe('#ffffff');

      // "2" / "black" -> #ffffff text
      const p2s = parseMathKatexParams('x^2', '2');
      expect(p2s.params?.background).toBe('#000000');
      expect(p2s.params?.color).toBe('#ffffff');

      const pBlack = parseMathKatexParams('x^2', 'black');
      expect(pBlack.params?.background).toBe('#000000');
      expect(pBlack.params?.color).toBe('#ffffff');

      // Dark custom hex -> #ffffff text
      const pDarkHex = parseMathKatexParams('x^2', '#102030');
      expect(pDarkHex.params?.background).toBe('#102030');
      expect(pDarkHex.params?.color).toBe('#ffffff');
    });

    it('honors explicit text color override over default on any background', () => {
      // Explicit green text on black background
      const p1 = parseMathKatexParams('x^2', 2, '#00ff00');
      expect(p1.params?.background).toBe('#000000');
      expect(p1.params?.color).toBe('#00ff00');

      // Explicit red text on white background
      const p2 = parseMathKatexParams('x^2', 1, '#ff0000');
      expect(p2.params?.background).toBe('#ffffff');
      expect(p2.params?.color).toBe('#ff0000');

      // Explicit yellow text on transparent background
      const p3 = parseMathKatexParams('x^2', 0, 'yellow');
      expect(p3.params?.background).toBe('transparent');
      expect(p3.params?.color).toBe('yellow');

      // Explicit black text on black background if explicitly requested
      const p4 = parseMathKatexParams('x^2', 2, '#000000');
      expect(p4.params?.background).toBe('#000000');
      expect(p4.params?.color).toBe('#000000');
    });

    it('falls back to adaptive text color when color is empty string, null, or whitespace', () => {
      const pEmptyBlack = parseMathKatexParams('x^2', 2, '');
      expect(pEmptyBlack.params?.color).toBe('#ffffff');

      const pNullBlack = parseMathKatexParams('x^2', 2, null);
      expect(pNullBlack.params?.color).toBe('#ffffff');

      const pSpaceBlack = parseMathKatexParams('x^2', 2, '   ');
      expect(pSpaceBlack.params?.color).toBe('#ffffff');

      const pEmptyWhite = parseMathKatexParams('x^2', 1, '');
      expect(pEmptyWhite.params?.color).toBe('#000000');
    });
  });

  // =========================================================================
  // 8. Omitted Parameters & Trailing Commas in Custom Function (R2)
  // =========================================================================
  describe('8. Omitted Parameters, Trailing Commas & Multi-Slot Shifting', () => {
    it('resiliently handles omitted arguments in all positions without #VALUE! error', () => {
      // 1 arg: =MATH.KATEX("\frac{2}{2}")
      const r1 = parseMathKatexParams('\\frac{2}{2}');
      expect(r1.isValid).toBe(true);
      expect(r1.params?.background).toBe('transparent');
      expect(r1.params?.color).toBe('#000000');
      expect(r1.params?.fontSize).toBe(16);

      // Trailing comma: =MATH.KATEX("\frac{2}{2}", )
      const r2 = parseMathKatexParams('\\frac{2}{2}', undefined);
      expect(r2.isValid).toBe(true);
      expect(r2.params?.background).toBe('transparent');

      // Trailing commas with empty strings: =MATH.KATEX("\frac{2}{2}", , )
      const r3 = parseMathKatexParams('\\frac{2}{2}', '', '');
      expect(r3.isValid).toBe(true);

      // Middle omitted: =MATH.KATEX("x^2", , , 24)
      const r4 = parseMathKatexParams('x^2', undefined, undefined, 24);
      expect(r4.isValid).toBe(true);
      expect(r4.params?.fontSize).toBe(24);
      expect(r4.params?.background).toBe('transparent');
      expect(r4.params?.color).toBe('#000000');

      // Middle omitted with 2: =MATH.KATEX("x^2", 2, , 20)
      const r5 = parseMathKatexParams('x^2', 2, undefined, 20);
      expect(r5.isValid).toBe(true);
      expect(r5.params?.background).toBe('#000000');
      expect(r5.params?.color).toBe('#ffffff');
      expect(r5.params?.fontSize).toBe(20);

      // Output method slot: =MATH.KATEX("x^2", , , , , "shape")
      const r6 = parseMathKatexParams('x^2', null, null, null, null, 'shape');
      expect(r6.isValid).toBe(true);
      expect(r6.params?.outputMethod).toBe('shape');
    });

    it('strips LaTeX delimiters ($$, \\[\\] , \\(\\), $) before parsing', () => {
      expect(stripLatexDelimiters('$$x^2 + y^2 = z^2$$')).toBe('x^2 + y^2 = z^2');
      expect(stripLatexDelimiters('\\[\\int f(x)dx\\]')).toBe('\\int f(x)dx');
      expect(stripLatexDelimiters('\\(\\alpha + \\beta\\)')).toBe('\\alpha + \\beta');
      expect(stripLatexDelimiters('$E = mc^2$')).toBe('E = mc^2');
      expect(stripLatexDelimiters('  $$ x = y $$  ')).toBe('x = y');
    });

    it('dynamically extracts invocation object across slots 2 to 7', () => {
      const inv = { address: 'Sheet1!B5' };

      // Invocation in slot 2: =MATH.KATEX("x^2")
      const resSlot2 = parseMathKatexParams('x^2', inv);
      expect(resSlot2.isValid).toBe(true);
      expect(resSlot2.params?.cellAddress).toBe('Sheet1!B5');
      expect(resSlot2.params?.background).toBe('transparent');

      // Invocation in slot 3: =MATH.KATEX("x^2", 2)
      const resSlot3 = parseMathKatexParams('x^2', 2, inv);
      expect(resSlot3.isValid).toBe(true);
      expect(resSlot3.params?.cellAddress).toBe('Sheet1!B5');
      expect(resSlot3.params?.background).toBe('#000000');
      expect(resSlot3.params?.color).toBe('#ffffff');

      // Invocation in slot 4: =MATH.KATEX("x^2", 1, "#107c41")
      const resSlot4 = parseMathKatexParams('x^2', 1, '#107c41', inv);
      expect(resSlot4.isValid).toBe(true);
      expect(resSlot4.params?.cellAddress).toBe('Sheet1!B5');
      expect(resSlot4.params?.background).toBe('#ffffff');
      expect(resSlot4.params?.color).toBe('#107c41');

      // Invocation in slot 5: =MATH.KATEX("x^2", 0, undefined, 20)
      const resSlot5 = parseMathKatexParams('x^2', 0, undefined, 20, inv);
      expect(resSlot5.isValid).toBe(true);
      expect(resSlot5.params?.cellAddress).toBe('Sheet1!B5');
      expect(resSlot5.params?.fontSize).toBe(20);

      // Invocation in slot 6: =MATH.KATEX("x^2", 1, undefined, 18, false)
      const resSlot6 = parseMathKatexParams('x^2', 1, undefined, 18, false, inv);
      expect(resSlot6.isValid).toBe(true);
      expect(resSlot6.params?.cellAddress).toBe('Sheet1!B5');
      expect(resSlot6.params?.displayMode).toBe(false);

      // Invocation in slot 7: =MATH.KATEX("x^2", 1, '#107c41', 20, false, 'cell')
      const resSlot7 = parseMathKatexParams('x^2', 1, '#107c41', 20, false, 'cell', inv);
      expect(resSlot7.isValid).toBe(true);
      expect(resSlot7.params?.cellAddress).toBe('Sheet1!B5');
      expect(resSlot7.params?.outputMethod).toBe('cell');
    });
  });

  // =========================================================================
  // 9. Core Rasterizer Canvas DOM-Walk & Off-Screen Probe (R1)
  // =========================================================================
  describe('9. Core Rasterizer DOM-Walk Probe & Canvas Operations', () => {
    it('positions probe off-screen with visibility:visible so DOM walk computes visible styles', async () => {
      const appendSpy = vi.spyOn(document.body, 'appendChild');
      const removeSpy = vi.spyOn(document.body, 'removeChild');

      await rasterizeLatex('E = mc^2', { fontSize: 18 });

      expect(appendSpy).toHaveBeenCalled();
      const probe = appendSpy.mock.calls[0][0] as HTMLElement;
      expect(probe.style.position).toBe('absolute');
      expect(probe.style.left).toBe('-99999px');
      expect(probe.style.top).toBe('0px');
      expect(probe.style.visibility).toBe('visible');
      expect(probe.style.opacity).toBe('1');
      expect(removeSpy).toHaveBeenCalledWith(probe);

      appendSpy.mockRestore();
      removeSpy.mockRestore();
    });

    it('invokes fillText for text nodes during DOM walk text rasterization', async () => {
      const dummyCanvas = document.createElement('canvas');
      const ctx = dummyCanvas.getContext('2d')!;

      await rasterizeLatex('a + b = c', { fontSize: 16 });

      expect(ctx.fillText).toHaveBeenCalled();
      const fillCalls = (ctx.fillText as any).mock.calls;
      const renderedTexts = fillCalls.map((call: any[]) => call[0]);
      expect(renderedTexts.some((t: string) => t.includes('a') || t.includes('b') || t.includes('+') || t.includes('='))).toBe(true);
    });

    it('handles transparent background by clearing canvas without filling opaque rect', async () => {
      const dummyCanvas = document.createElement('canvas');
      const ctx = dummyCanvas.getContext('2d')!;

      await rasterizeLatex('x + y', { background: 'transparent' });
      expect(ctx.clearRect).toHaveBeenCalled();
      expect(ctx.fillRect).not.toHaveBeenCalled();

      vi.clearAllMocks();
      await rasterizeLatex('y + z', { background: 0 });
      expect(ctx.clearRect).toHaveBeenCalled();
      expect(ctx.fillRect).not.toHaveBeenCalled();
    });

    it('handles solid background by filling canvas with specified color', async () => {
      const dummyCanvas = document.createElement('canvas');
      const ctx = dummyCanvas.getContext('2d')!;

      await rasterizeLatex('x + y', { background: '#ffffff' });
      expect(ctx.fillRect).toHaveBeenCalled();

      vi.clearAllMocks();
      await rasterizeLatex('k + 1', { background: '#000000' });
      expect(ctx.fillRect).toHaveBeenCalled();
      expect(ctx.fillStyle).toBe('#000000');
    });

    it('renders fraction lines, overlines, underlines with canvas line strokes', async () => {
      const dummyCanvas = document.createElement('canvas');
      const ctx = dummyCanvas.getContext('2d')!;

      await rasterizeLatex('\\frac{1}{2} + \\overline{AB} + \\underline{CD}', { fontSize: 16 });

      expect(ctx.beginPath).toHaveBeenCalled();
      expect(ctx.moveTo).toHaveBeenCalled();
      expect(ctx.lineTo).toHaveBeenCalled();
      expect(ctx.stroke).toHaveBeenCalled();
    });
  });

  // =========================================================================
  // 10. End-to-End =MATH.KATEX() Custom Function Evaluation (R1 & R2)
  // =========================================================================
  describe('10. End-to-End =MATH.KATEX() Custom Function Evaluation', () => {
    it('evaluates =MATH.KATEX("\\frac{2}{2}", 0) with transparent background & black text', async () => {
      const inv = CustomFunctionsMock.createInvocation('Sheet1!A1');
      const res = await mathKatex('\\frac{2}{2}', 0, undefined, 16, true, 'cell', inv);

      expect(typeof res).toBe('object');
      const entity = res as KatexEntityCellValue;
      expect(entity.type).toBe('Entity');
      expect(entity.text).toBe('[Math: \\frac{2}{2}]');
      expect(entity.properties?.image?.address.startsWith('data:image/png')).toBe(true);
    });

    it('evaluates =MATH.KATEX("\\frac{2}{2}", 1) with white background & black text', async () => {
      const inv = CustomFunctionsMock.createInvocation('Sheet1!A2');
      const res = await mathKatex('\\frac{2}{2}', 1, undefined, 16, true, 'cell', inv);

      expect(typeof res).toBe('object');
      const entity = res as KatexEntityCellValue;
      expect(entity.type).toBe('Entity');
      expect(entity.properties?.image?.address.startsWith('data:image/png')).toBe(true);
    });

    it('evaluates =MATH.KATEX("\\frac{2}{2}", 2) with black background & white text', async () => {
      const inv = CustomFunctionsMock.createInvocation('Sheet1!A3');
      const res = await mathKatex('\\frac{2}{2}', 2, undefined, 16, true, 'cell', inv);

      expect(typeof res).toBe('object');
      const entity = res as KatexEntityCellValue;
      expect(entity.type).toBe('Entity');
      expect(entity.properties?.image?.address.startsWith('data:image/png')).toBe(true);
      expect(entity.properties?.svg?.basicValue).toContain('background-color: #000000;');
      expect(entity.properties?.svg?.basicValue).toContain('color: #ffffff;');
    });

    it('evaluates omitted parameters in =MATH.KATEX("x^2", ) without #VALUE!', async () => {
      const inv = CustomFunctionsMock.createInvocation('Sheet1!A4');
      const res = await mathKatex('x^2', undefined, undefined, undefined, undefined, undefined, inv);
      expect(typeof res).toBe('object');
      const entity = res as KatexEntityCellValue;
      expect(entity.type).toBe('Entity');
    });

    it('returns #VALUE! error on invalid syntax without unhandled exceptions', async () => {
      const inv = CustomFunctionsMock.createInvocation('Sheet1!A5');
      const invalidFormulas = [
        '\\frac{1}',
        '\\begin{pmatrix} 1 & 2',
        '\\invalidMacro{foo}',
        '',
        '   ',
        '<script>alert(1)</script>'
      ];

      for (const formula of invalidFormulas) {
        await expect(mathKatex(formula, 0, undefined, 16, true, 'cell', inv)).rejects.toThrow();
      }
    });
  });

  // =========================================================================
  // 11. Invalid Syntax, Fuzzing & Unhandled Exception Defense
  // =========================================================================
  describe('11. Invalid Syntax, Fuzzing & Error Handling', () => {
    const invalidFormulas = [
      { name: 'unclosed brace', latex: '\\frac{1}{2' },
      { name: 'unclosed environment', latex: '\\begin{pmatrix} 1 & 2' },
      { name: 'mismatched environment', latex: '\\begin{matrix} 1 & 2 \\end{aligned}' },
      { name: 'dangling backslash', latex: 'x + y \\' },
      { name: 'unknown command', latex: '\\thisCommandDoesNotExist{foo}' },
      { name: 'incomplete subscript', latex: 'x_' },
      { name: 'incomplete superscript', latex: 'x^' },
      { name: 'double superscript conflict', latex: 'x^2^3' },
      { name: 'dangling fraction', latex: '\\frac' },
      { name: 'unmatched left/right', latex: '\\left( x + y' },
      { name: 'reversed delimiters', latex: '\\right) x + y \\left(' },
      { name: 'incomplete sqrt index', latex: '\\sqrt[' }
    ];

    invalidFormulas.forEach(({ name, latex }) => {
      it(`gracefully handles invalid syntax (${name}): "${latex}" without unhandled crash`, () => {
        const val = validateLatex(latex);
        expect(val.isValid).toBe(false);
        expect(val.error).toBeDefined();
        expect(typeof val.error).toBe('string');
        expect(val.error?.length).toBeGreaterThan(0);

        const safeHtml = compileLatexToHtml(latex, { throwOnError: false });
        expect(safeHtml).toContain('katex-error');

        const paramOutcome = parseMathKatexParams(latex);
        expect(paramOutcome.isValid).toBe(false);
        expect(paramOutcome.errorCode).toBe('#VALUE!');
      });
    });

    it('rejects XSS and script injection attacks', () => {
      const attacks = [
        '<script>alert("pwned")</script>',
        '<img src=x onerror="alert(1)">',
        '"><svg onload=alert(1)>',
        '\\text{<script>alert(1)</script>}',
        'javascript:void(0)',
        'onload=document.write(1)'
      ];

      attacks.forEach((attack) => {
        const val = validateLatex(attack);
        if (!val.isValid) {
          expect(val.error).toBeDefined();
        } else {
          const html = compileLatexToHtml(attack, { throwOnError: false });
          expect(html).not.toContain('<script>');
          expect(html).not.toContain('onerror=');
          expect(html).not.toContain('onload=');
        }
      });
    });
  });

  // =========================================================================
  // 12. LRU Cache Stress & Eviction Correctness
  // =========================================================================
  describe('12. LRU Cache High-Throughput Stress & Eviction Correctness', () => {
    it('executes 10,000 rapid LRU cache operations with accurate eviction count', () => {
      const cache = new LRUCache<string, number>({ maxSize: 250 });
      const N = 10000;

      const t0 = performance.now();
      for (let i = 0; i < N; i++) {
        const key = `key_${i % 1000}`;
        if (i % 3 === 0) {
          cache.get(key);
        } else {
          cache.set(key, i);
        }
      }
      const t1 = performance.now();

      expect(t1 - t0).toBeLessThan(1000);
      expect(cache.size()).toBe(250);

      const stats = cache.getStats();
      expect(stats.size).toBe(250);
      expect(stats.maxSize).toBe(250);
      expect(stats.evictions).toBeGreaterThan(0);
      expect(stats.totalRequests).toBe(stats.hits + stats.misses);
    });

    it('handles TTL expiration properly', () => {
      vi.useFakeTimers();
      const cache = new LRUCache<string, string>({ maxSize: 10, defaultTtlMs: 1000 });

      cache.set('temp', 'value1');
      expect(cache.get('temp')).toBe('value1');
      expect(cache.has('temp')).toBe(true);

      vi.advanceTimersByTime(1500);

      expect(cache.has('temp')).toBe(false);
      expect(cache.get('temp')).toBeUndefined();
      expect(cache.size()).toBe(0);

      vi.useRealTimers();
    });
  });
});
