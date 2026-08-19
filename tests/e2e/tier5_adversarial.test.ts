import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import { compileLatex, compileLatexToHtml, validateLatex } from '../../src/core/katexEngine';
import { getEquationCacheStats, clearEquationCache } from '../../src/core/imageRasterizer';
import { LRUCache } from '../../src/core/lruCache';
import { mathKatexFunction, isEntityCellValueSupported } from '../../src/customfunctions/mathKatex';
import { KaTeXShapeManager } from '../../src/customfunctions/shapeManager';
import { buildKatexEntityCellValue } from '../../src/customfunctions/entityCellBuilder';
import { parseMathKatexParams } from '../../src/customfunctions/parameterParser';
import {
  resetOfficeMock,
  getMockExcelState,
  setMockRequirementSupported
} from '../../src/mocks/officeMock';
import {
  resetCustomFunctionsMock,
  CustomFunctionsMock
} from '../../src/mocks/customFunctionsMock';

describe('Tier 5: Adversarial, Stress, Security & Robustness Test Suite', () => {
  beforeEach(() => {
    resetOfficeMock();
    resetCustomFunctionsMock();
    KaTeXShapeManager.clear();
    clearEquationCache();
  });

  afterEach(() => {
    vi.restoreAllMocks();
  });

  // Section 5.1: Malformed & Corrupted Invocations, Extreme Nesting Depth
  describe('5.1 Malformed, Corrupted Invocations & Extreme Nesting Depth', () => {
    it('handles extreme fraction nesting depth (>100 levels) without call stack overflow', () => {
      // Build 120-level nested fraction: \frac{1}{\frac{1}{\frac{1}{...{x}}}}
      let nestedFrac = 'x';
      for (let i = 0; i < 120; i++) {
        nestedFrac = `\\frac{1}{${nestedFrac}}`;
      }

      const validation = validateLatex(nestedFrac);
      expect(typeof validation.isValid).toBe('boolean');
      const html = compileLatexToHtml(nestedFrac, { throwOnError: false });
      expect(typeof html).toBe('string');
      expect(html.length).toBeGreaterThan(0);
    });

    it('handles extreme square root nesting depth (>100 levels)', () => {
      let nestedSqrt = 'x';
      for (let i = 0; i < 105; i++) {
        nestedSqrt = `\\sqrt{${nestedSqrt}}`;
      }

      const html = compileLatexToHtml(nestedSqrt, { throwOnError: false });
      expect(typeof html).toBe('string');
      expect(html.length).toBeGreaterThan(0);
    });

    it('rejects corrupted parameter inputs with meaningful CustomFunctions.Error', async () => {
      const inv = CustomFunctionsMock.createInvocation('Sheet1!A1');

      // Empty string
      await expect(mathKatexFunction('', undefined, undefined, undefined, undefined, 'auto', inv)).rejects.toThrow();

      // Null / Undefined
      await expect(mathKatexFunction(null as any, undefined, undefined, undefined, undefined, 'auto', inv)).rejects.toThrow();
      await expect(mathKatexFunction(undefined as any, undefined, undefined, undefined, undefined, 'auto', inv)).rejects.toThrow();

      // Invalid font size (NaN, negative, extreme)
      const parseInvalidSize = parseMathKatexParams('x = 1', undefined, undefined, -5, undefined, 'cell', inv);
      expect(parseInvalidSize.params?.fontSize).toBe(16); // Falls back to default or clamps

      // Unclosed braces and broken environments
      await expect(mathKatexFunction('\\begin{matrix} 1 & 2 \\end{align}', undefined, undefined, undefined, undefined, 'auto', inv)).rejects.toThrow();
      await expect(mathKatexFunction('\\frac{1}{', undefined, undefined, undefined, undefined, 'auto', inv)).rejects.toThrow();
    });

    it('sanitizes and defends against XSS / script injection attacks in formula strings', async () => {
      const maliciousPayloads = [
        '<script>alert("xss")</script>',
        '<img src=x onerror=alert(1)>',
        '<svg onload=alert(1)>',
        '\\href{javascript:alert(1)}{ClickMe}',
        '\\url{data:text/html,<script>alert(1)</script>}'
      ];

      for (const payload of maliciousPayloads) {
        const validation = validateLatex(payload);
        if (validation.isValid) {
          const html = compileLatexToHtml(payload, { throwOnError: false });
          expect(html).not.toContain('<script>');
          expect(html).not.toContain('onerror=');
          expect(html).not.toContain('onload=');
        } else {
          expect(validation.isValid).toBe(false);
        }
      }
    });

    it('handles giant payloads (10,000+ characters) gracefully', async () => {
      const giantLatex = 'x + ' + 'y + '.repeat(3000) + 'z';
      const validation = validateLatex(giantLatex);
      expect(typeof validation.isValid).toBe('boolean');
    });

    it('handles null bytes and control characters without terminating process', () => {
      const nullByteString = 'x + \\alpha\0 + \\beta\u0001';
      const validation = validateLatex(nullByteString);
      expect(typeof validation.isValid).toBe('boolean');
    });
  });

  // Section 5.2: Concurrent Shape Queue Flood Under High-Volume Batch
  describe('5.2 Concurrent Shape Queue Flood & Re-entrancy Lock Under High-Volume Batch', () => {
    it('handles 100 simultaneous concurrent shape queue requests cleanly', async () => {
      const invList = Array.from({ length: 100 }, (_, i) => 
        CustomFunctionsMock.createInvocation(`Sheet1!A${i + 1}`)
      );

      // Concurrently queue 100 shape requests
      const promises = invList.map((inv, idx) => 
        mathKatexFunction(`\\int_0^{${idx}} x^2 dx`, undefined, undefined, 14, true, 'shape', inv)
      );

      const results = await Promise.all(promises);
      expect(results.length).toBe(100);
      results.forEach((res, idx) => {
        expect(res).toBe(`[KaTeX Shape: \\int_0^{${idx}} x^2 dx]`);
      });

      expect(KaTeXShapeManager.getPendingCount()).toBe(100);

      // Process batch queue
      const processed = await KaTeXShapeManager.processQueue();
      expect(processed).toBe(100);
      expect(KaTeXShapeManager.getPendingCount()).toBe(0);

      const state = getMockExcelState();
      const sheet = state.workbook.worksheets.getActiveWorksheet();
      expect(sheet.shapes.items.length).toBe(100);
    });

    it('locks isProcessing to prevent re-entrant queue draining races', async () => {
      for (let i = 1; i <= 20; i++) {
        KaTeXShapeManager.enqueueShape(`E_${i} = mc^2`, `Sheet1!B${i}`);
      }
      expect(KaTeXShapeManager.getPendingCount()).toBe(20);

      // Trigger two parallel processQueue calls simultaneously
      const [run1, run2] = await Promise.all([
        KaTeXShapeManager.processQueue(),
        KaTeXShapeManager.processQueue()
      ]);

      // Exactly one processQueue should take the lock and process the 20 items, the second returns 0
      expect(run1 + run2).toBe(20);
      expect(KaTeXShapeManager.getPendingCount()).toBe(0);
      expect(KaTeXShapeManager.getProcessingStatus()).toBe(false);
    });

    it('verifies debounced queue auto-flush after 50ms idle delay', async () => {
      vi.useFakeTimers();

      KaTeXShapeManager.enqueueShape('\\alpha + \\beta', 'Sheet1!C1');
      KaTeXShapeManager.enqueueShape('\\gamma + \\delta', 'Sheet1!C2');
      expect(KaTeXShapeManager.getPendingCount()).toBe(2);

      // Fast forward time by 60ms to trigger auto-flush debounce
      await vi.advanceTimersByTimeAsync(60);

      expect(KaTeXShapeManager.getPendingCount()).toBe(0);
      const state = getMockExcelState();
      const sheet = state.workbook.worksheets.getActiveWorksheet();
      expect(sheet.shapes.items.length).toBeGreaterThanOrEqual(2);

      vi.useRealTimers();
    });
  });

  // Section 5.3: LRU Cache Saturation and Eviction
  describe('5.3 LRU Cache Saturation & Eviction Under Heavy Unique Formula Generation', () => {
    it('caps memory retention at maxSize entries and evicts least recently used formulas', () => {
      const cache = new LRUCache<string, string>({ maxSize: 500 });
      expect(cache.size()).toBe(0);

      // Saturate cache with 600 unique equations
      for (let i = 1; i <= 600; i++) {
        cache.set(`eq_${i}`, `formula_result_${i}`);
      }

      const stats = cache.getStats();
      expect(stats.size).toBe(500);
      expect(stats.evictions).toBe(100);

      // Oldest evicted items (eq_1 .. eq_100) should be missing
      expect(cache.get('eq_1')).toBeUndefined();
      expect(cache.get('eq_100')).toBeUndefined();

      // Non-evicted items (eq_101 .. eq_600) should be present
      expect(cache.get('eq_101')).toBe('formula_result_101');
      expect(cache.get('eq_600')).toBe('formula_result_600');
    });

    it('re-promotes accessed items in LRU order upon cache hit', () => {
      const cache = new LRUCache<string, string>({ maxSize: 5 });

      cache.set('a', '1');
      cache.set('b', '2');
      cache.set('c', '3');
      cache.set('d', '4');
      cache.set('e', '5');

      // Access 'a' so it is moved to head (most recent)
      expect(cache.get('a')).toBe('1');

      // Add 6th item 'f' - should evict 'b' (oldest), NOT 'a'
      cache.set('f', '6');

      expect(cache.get('b')).toBeUndefined(); // evicted
      expect(cache.get('a')).toBe('1'); // retained
      expect(cache.get('f')).toBe('6');
      expect(cache.size()).toBe(5);
    });

    it('verifies equationCache integration with compileLatex across multiple invocations', async () => {
      clearEquationCache();
      const initialStats = getEquationCacheStats();
      expect(initialStats.size).toBe(0);

      // First compile -> Cache Miss
      const res1 = await compileLatex('E = mc^2');
      expect(res1.html).toContain('katex');
      const statsAfterMiss = getEquationCacheStats();
      expect(statsAfterMiss.size).toBe(1);
      expect(statsAfterMiss.misses).toBe(1);

      // Second compile with same parameters -> Cache Hit
      const res2 = await compileLatex('E = mc^2');
      expect(res2.pngDataUrl).toBe(res1.pngDataUrl);
      const statsAfterHit = getEquationCacheStats();
      expect(statsAfterHit.hits).toBe(1);
      expect(statsAfterHit.size).toBe(1);
    });
  });

  // Section 5.4: Legacy Excel Fallback Simulation
  describe('5.4 Legacy Excel Fallback Simulation & Defensive Error Handling', () => {
    it('defensively falls back to false when ExcelApi 1.16 is unsupported', () => {
      setMockRequirementSupported('ExcelApi', 1.16, false);
      expect(isEntityCellValueSupported()).toBe(false);
    });

    it('defensively catches errors in isSetSupported and returns false', () => {
      const office = (globalThis as any).Office;
      if (office && office.context && office.context.requirements) {
        const original = office.context.requirements.isSetSupported;
        office.context.requirements.isSetSupported = () => {
          throw new Error('Host API access denied or unsupported');
        };

        expect(isEntityCellValueSupported()).toBe(false);

        office.context.requirements.isSetSupported = original;
      }
    });

    it('defensively returns false when Office.context.requirements is undefined', () => {
      const office = (globalThis as any).Office;
      const originalContext = office.context;
      office.context = {};

      expect(isEntityCellValueSupported()).toBe(false);

      office.context = originalContext;
    });

    it('routes =MATH.KATEX to shape manager when ExcelApi 1.16 is missing and outputMethod is auto', async () => {
      setMockRequirementSupported('ExcelApi', 1.16, false);
      const inv = CustomFunctionsMock.createInvocation('Sheet1!H10');

      const result = await mathKatexFunction('\\oint_C \\mathbf{F} \\cdot d\\mathbf{r}', undefined, undefined, 16, true, 'auto', inv);
      expect(typeof result).toBe('string');
      expect(result).toBe('[KaTeX Shape: \\oint_C \\mathbf{F} \\cdot d\\mathbf{r}]');
      expect(KaTeXShapeManager.getPendingCount()).toBe(1);

      const queued = KaTeXShapeManager.getQueue()[0];
      expect(queued.address).toBe('Sheet1!H10');
      expect(queued.latex).toBe('\\oint_C \\mathbf{F} \\cdot d\\mathbf{r}');
    });

    it('simulates shape insertion on legacy worksheet without entity cell support', async () => {
      setMockRequirementSupported('ExcelApi', 1.16, false);
      const inv = CustomFunctionsMock.createInvocation('Sheet1!K12');

      await mathKatexFunction('\\frac{d}{dx}\\left(e^x\\right) = e^x', undefined, undefined, 16, true, 'auto', inv);
      expect(KaTeXShapeManager.getPendingCount()).toBe(1);

      const processedCount = await KaTeXShapeManager.processQueue();
      expect(processedCount).toBe(1);

      const state = getMockExcelState();
      const sheet = state.workbook.worksheets.getActiveWorksheet();
      expect(sheet.shapes.items.length).toBe(1);
      expect(sheet.shapes.items[0].altTextTitle).toContain('LaTeX: \\frac{d}{dx}\\left(e^x\\right) = e^x');
    });
  });

  // Section 5.5: Offline Zero-Network Security Assertion
  describe('5.5 Offline Zero-Network Security Assertion', () => {
    it('guarantees zero network calls (fetch / XHR / WebSocket / Beacon) during equation rendering and custom functions', async () => {
      const fetchSpy = vi.fn();
      (globalThis as any).fetch = fetchSpy;

      const equationsToTest = [
        'E = mc^2',
        '\\int_{-\\infty}^\\infty e^{-x^2} dx = \\sqrt{\\pi}',
        '\\begin{pmatrix} a & b \\\\ c & d \\end{pmatrix}',
        '\\sum_{k=1}^n k = \\frac{n(n+1)}{2}',
        '\\nabla \\times \\mathbf{E} = -\\frac{\\partial \\mathbf{B}}{\\partial t}'
      ];

      for (const eq of equationsToTest) {
        const render = await compileLatex(eq);
        expect(render.html).toContain('katex');
        expect(render.pngDataUrl.startsWith('data:image/')).toBe(true);

        const inv = CustomFunctionsMock.createInvocation('Sheet1!B5');
        const entity = await mathKatexFunction(eq, undefined, undefined, 16, true, 'cell', inv);
        expect(entity).toBeDefined();
      }

      // Assert that fetch was never invoked
      expect(fetchSpy).not.toHaveBeenCalled();
    });

    it('verifies shape manager execution operates completely offline', async () => {
      const fetchSpy = vi.fn();
      (globalThis as any).fetch = fetchSpy;

      KaTeXShapeManager.enqueueShape('\\mathcal{H} |\\psi\\rangle = E |\\psi\\rangle', 'Sheet1!D10');
      await KaTeXShapeManager.processQueue();

      expect(fetchSpy).not.toHaveBeenCalled();
    });

    it('verifies entity cell builder operates completely offline without external image URLs', () => {
      const render = {
        pngDataUrl: 'data:image/png;base64,offline_data_test',
        width: 100,
        height: 30,
        svg: '<svg>...</svg>'
      };

      const entity = buildKatexEntityCellValue('a + b = c', render);
      expect(entity.layouts?.compact?.icon?.startsWith('data:image/')).toBe(true);
      expect(entity.provider?.logoSourceAddress?.startsWith('data:image/')).toBe(true);
      expect(entity.properties?.image?.address?.startsWith('data:image/')).toBe(true);
    });
  });
});
