import { describe, it, expect, beforeEach, vi } from 'vitest';
import {
  rasterizeLatex,
  generateStandaloneSvg,
  svgToPngDataUrl,
  clearEquationCache,
  getEquationCacheStats
} from '../../src/core/imageRasterizer';

describe('imageRasterizer.ts - Image and SVG Generator', () => {
  beforeEach(() => {
    clearEquationCache();
    vi.clearAllMocks();
  });

  it('should generate valid standalone SVG markup with explicit dimensions', () => {
    const svg = generateStandaloneSvg('<span class="katex">x</span>', 120, 60, 'transparent', '#000000', 21.3);
    expect(svg).toContain('<svg');
    expect(svg).toContain('xmlns="http://www.w3.org/2000/svg"');
    expect(svg).toContain('width="120"');
    expect(svg).toContain('height="60"');
    expect(svg).toContain('viewBox="0 0 120 60"');
    expect(svg).toContain('<foreignObject');
  });

  it('should apply background styling to SVG when specified', () => {
    const svg = generateStandaloneSvg('<span>test</span>', 100, 50, '#ffffff', '#111111', 18);
    expect(svg).toContain('background-color: #ffffff;');
    expect(svg).toContain('color: #111111;');
  });

  it('should not apply background styling to SVG when transparent or 0', () => {
    const svg1 = generateStandaloneSvg('<span>test</span>', 100, 50, 'transparent', '#000000', 16);
    expect(svg1).not.toContain('background-color:');

    const svg2 = generateStandaloneSvg('<span>test</span>', 100, 50, '0', '#000000', 16);
    expect(svg2).not.toContain('background-color:');
  });

  it('should rasterize LaTeX and return cached result on second call', async () => {
    const formula = '\\frac{a+b}{c-d}';
    const res1 = await rasterizeLatex(formula, { fontSize: 16 });
    expect(res1.svg).toContain('<svg');
    expect(res1.width).toBeGreaterThan(0);
    expect(res1.height).toBeGreaterThan(0);
    expect(res1.pngDataUrl.startsWith('data:image/png')).toBe(true);

    const stats1 = getEquationCacheStats();
    expect(stats1.size).toBe(1);

    // Second call should hit the cache
    const res2 = await rasterizeLatex(formula, { fontSize: 16 });
    expect(res2).toEqual(res1);

    const stats2 = getEquationCacheStats();
    expect(stats2.hits).toBe(1);
  });

  it('should generate data URL via svgToPngDataUrl', async () => {
    const svg = '<svg xmlns="http://www.w3.org/2000/svg" width="100" height="50"></svg>';
    const dataUrl = await svgToPngDataUrl(svg, 100, 50, 2);
    expect(dataUrl.startsWith('data:image/png')).toBe(true);
  });

  it('should position probe off-screen with visibility:visible so DOM walk computes visible styles', async () => {
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

  it('should invoke fillText for text nodes during DOM walk text rasterization', async () => {
    const dummyCanvas = document.createElement('canvas');
    const ctx = dummyCanvas.getContext('2d')!;

    await rasterizeLatex('a + b = c', { fontSize: 16 });

    expect(ctx.fillText).toHaveBeenCalled();
    const fillCalls = (ctx.fillText as any).mock.calls;
    const renderedTexts = fillCalls.map((call: any[]) => call[0]);
    expect(renderedTexts.some((t: string) => t.includes('a') || t.includes('b') || t.includes('+') || t.includes('='))).toBe(true);
  });

  it('should handle transparent background by clearing canvas and not filling opaque background', async () => {
    const dummyCanvas = document.createElement('canvas');
    const ctx = dummyCanvas.getContext('2d')!;

    await rasterizeLatex('x + y', { background: 'transparent' });
    expect(ctx.clearRect).toHaveBeenCalled();
    expect(ctx.fillRect).not.toHaveBeenCalled();

    vi.clearAllMocks();
    await rasterizeLatex('y + z', { background: '0' });
    expect(ctx.clearRect).toHaveBeenCalled();
    expect(ctx.fillRect).not.toHaveBeenCalled();
  });

  it('should handle solid background by filling canvas with specified color', async () => {
    const dummyCanvas = document.createElement('canvas');
    const ctx = dummyCanvas.getContext('2d')!;

    await rasterizeLatex('x + y', { background: '#ffffff' });
    expect(ctx.fillRect).toHaveBeenCalled();

    vi.clearAllMocks();
    await rasterizeLatex('k + 1', { background: '#ff0000' });
    expect(ctx.fillRect).toHaveBeenCalled();
    expect(ctx.fillStyle).toBe('#ff0000');
  });

  it('should render fraction lines and math elements with canvas line strokes', async () => {
    const dummyCanvas = document.createElement('canvas');
    const ctx = dummyCanvas.getContext('2d')!;

    await rasterizeLatex('\\frac{1}{2}', { fontSize: 16 });

    expect(ctx.beginPath).toHaveBeenCalled();
    expect(ctx.moveTo).toHaveBeenCalled();
    expect(ctx.lineTo).toHaveBeenCalled();
    expect(ctx.stroke).toHaveBeenCalled();
  });

  it('should preserve high-DPI scaling factor', async () => {
    const dummyCanvas = document.createElement('canvas');
    const ctx = dummyCanvas.getContext('2d')!;

    await rasterizeLatex('\\alpha + \\beta', { scale: 4 });
    expect(ctx.scale).toHaveBeenCalledWith(4, 4);
  });

  it('should handle KaTeX syntax errors gracefully and render error message', async () => {
    const res = await rasterizeLatex('\\invalidMacroThatDoesNotExist{123}', { fontSize: 16 });
    expect(res).toBeDefined();
    expect(res.pngDataUrl.startsWith('data:image/png')).toBe(true);
    expect(res.width).toBeGreaterThan(0);
    expect(res.height).toBeGreaterThan(0);
  });
});

