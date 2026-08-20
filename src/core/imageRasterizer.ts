/**
 * ExcelKaTeX Image Rasterizer — DOM-Walk approach (WebView2-safe)
 *
 * WHY PREVIOUS APPROACHES FAILED:
 *   Edge WebView2 (Excel's browser) blocks ALL foreignObject content in SVG
 *   when loaded via <img src="blob:...">. This includes HTML, MathML, CSS.
 *   The img.onload fires but the canvas is blank (just background color).
 *
 * THE WORKING APPROACH (this file):
 *   1. Render KaTeX HTML in the LIVE PAGE DOM where KaTeX CSS + fonts are active.
 *   2. Walk every text node → draw on canvas with getComputedStyle font.
 *   3. Walk every <svg> sub-element (sqrt signs, etc.) → these are pure SVG paths,
 *      no foreignObject → can be loaded as <img> and drawn to canvas without taint.
 *   4. Walk .frac-line elements → draw as canvas line strokes.
 *   5. Hard 8-second timeout so buttons NEVER hang forever.
 */

import katex from 'katex';
import { LRUCache, createRenderCacheKey } from './lruCache';
import { macroRegistry, MacroDictionary } from './macros';

export interface RenderOptions {
  background?: string | number;
  color?: string;
  fontSize?: number;
  displayMode?: boolean;
  scale?: number;
  macros?: MacroDictionary;
}

export interface RenderResult {
  html: string;
  svg: string;
  pngDataUrl: string;
  width: number;
  height: number;
  aspectRatio: number;
}

// ─── LRU Cache ───────────────────────────────────────────────────────────────
const equationCache = new LRUCache<string, RenderResult>({ maxSize: 500 });
export function getEquationCacheStats() { return equationCache.getStats(); }
export function clearEquationCache()    { equationCache.clear(); }

// ─── Utilities ────────────────────────────────────────────────────────────────
function ptToPx(pt: number): number {
  return Math.round(pt * 1.33333 * 10) / 10;
}

/**
 * Core renderer: places KaTeX HTML in the live page DOM (where KaTeX CSS and
 * fonts are fully active), then walks the DOM tree to redraw every element
 * onto an HTML5 canvas using the Canvas 2D API.
 *
 * This is the only approach that reliably works in Edge WebView2 inside Excel.
 */
async function renderKatexToCanvas(
  katexHtml: string,
  background: string | number,
  color: string,
  fontSizePx: number,
  scale: number
): Promise<{ dataUrl: string; width: number; height: number }> {
  const fallback = { dataUrl: '', width: 200, height: 60 };
  if (typeof document === 'undefined') return fallback;

  const isTransparent = !background || background === 'transparent' || background === 'none' || background === '0' || background === 0;
  let bg = 'transparent';
  if (!isTransparent) {
    if (background === 1 || background === '1' || background === 'white') {
      bg = '#ffffff';
    } else if (background === 2 || background === '2' || background === 'black') {
      bg = '#000000';
    } else {
      bg = String(background);
    }
  }

  // ── Step 1: Render KaTeX HTML in the page DOM ─────────────────────────────
  const probe = document.createElement('div');
  probe.style.cssText = [
    'position:absolute', 'top:0px', 'left:-99999px',
    'visibility:visible', 'opacity:1', 'pointer-events:none', 'z-index:-9999',
    `font-size:${fontSizePx}px`, `color:${color}`, `background:${bg}`,
    'white-space:nowrap', 'display:inline-block',
    'padding:4px 6px', 'box-sizing:border-box',
  ].join(';');
  probe.innerHTML = katexHtml;
  document.body.appendChild(probe);

  // ── Step 2: Wait for KaTeX fonts to be ready ──────────────────────────────
  try {
    await Promise.race([
      document.fonts.ready,
      new Promise<void>(r => setTimeout(r, 3000)),
    ]);
  } catch { /* ignore */ }

  // ── Step 3: Measure actual rendered size ──────────────────────────────────
  const containerRect = probe.getBoundingClientRect();
  let measuredWidth = containerRect.width;
  let measuredHeight = containerRect.height;

  // Headless JSDOM fallback (when getBoundingClientRect returns 0x0)
  if (measuredWidth === 0 || measuredHeight === 0) {
    const charCount = probe.textContent?.length || 10;
    const matrixMatchCount = (katexHtml.match(/\\\\|<tr|class="matrix-row"|class="col-align/g) || []).length;
    const rowCount = Math.max(1, matrixMatchCount);
    measuredWidth = Math.max(80, Math.ceil(fontSizePx * Math.max(2, charCount * 0.6)));
    measuredHeight = Math.max(36, Math.ceil(fontSizePx * (1.2 + (rowCount - 1) * 0.9)));
  }

  const width  = Math.max(80, Math.ceil(measuredWidth)  + 4);
  const height = Math.max(36, Math.ceil(measuredHeight) + 4);

  // ── Step 4: Set up canvas ─────────────────────────────────────────────────
  const canvas = document.createElement('canvas');
  canvas.width  = Math.ceil(width  * scale);
  canvas.height = Math.ceil(height * scale);
  const ctx = canvas.getContext('2d');
  if (!ctx) { document.body.removeChild(probe); return fallback; }

  ctx.scale(scale, scale);

  // Fill or clear background
  if (!isTransparent) {
    ctx.fillStyle = bg;
    ctx.fillRect(0, 0, width, height);
  } else {
    ctx.clearRect(0, 0, width, height);
  }

  // Offset from container's top-left corner
  const ox = containerRect.left;
  const oy = containerRect.top;

  // ── Step 5: Draw math lines (CSS border-bottom / border-top, no text node) ─
  const mathLines = probe.querySelectorAll('.frac-line, .overline-line, .underline-line, .hline');
  for (const el of mathLines) {
    const isOverline = el.classList.contains('overline-line');
    const style = window.getComputedStyle(el);
    const lineW = parseFloat(style.borderBottomWidth) || parseFloat(style.borderTopWidth) || 0.6;
    
    let rect = el.getBoundingClientRect();
    let x = rect.left - ox;
    let y = isOverline ? (rect.top - oy) : (rect.top - oy + rect.height);
    let w = rect.width;

    // Strictly clamp fraction lines to the exact bounding box of their numerator & denominator contents
    if (el.classList.contains('frac-line')) {
      const mfrac = el.closest('.mfrac');
      if (mfrac) {
        const items = mfrac.querySelectorAll('.mord, .mop, .mbin, .mrel, .mopen, .mclose, .mpunct, .minner, svg, .sqrt, .katex-logo');
        let minLeft = Infinity;
        let maxRight = -Infinity;
        for (let i = 0; i < items.length; i++) {
          const item = items[i];
          if (item !== el && !item.classList.contains('frac-line')) {
            const r = item.getBoundingClientRect();
            if (r.width > 0 && r.height > 0) {
              if (r.left < minLeft) minLeft = r.left;
              if (r.right > maxRight) maxRight = r.right;
            }
          }
        }
        if (maxRight > minLeft && isFinite(minLeft)) {
          x = (minLeft - ox) - 1;
          w = (maxRight - minLeft) + 2;
        } else {
          const mfracRect = mfrac.getBoundingClientRect();
          if (mfracRect.width > 0) {
            x = mfracRect.left - ox;
            w = mfracRect.width;
          }
        }
      }
    }

    if (w <= 0) {
      w = 30;
      x = 10;
      y = isOverline ? 10 : 25;
    }

    ctx.save();
    ctx.strokeStyle = style.borderBottomColor || style.borderTopColor || style.color || color;
    ctx.lineWidth   = lineW;
    ctx.beginPath();
    ctx.moveTo(x, y);
    ctx.lineTo(x + w, y);
    ctx.stroke();
    ctx.restore();
  }

  // ── Step 6: Walk text nodes and draw them ─────────────────────────────────
  const walker = document.createTreeWalker(probe, NodeFilter.SHOW_TEXT);
  let tnode = walker.nextNode();
  while (tnode) {
    const text = tnode.textContent || '';
    if (text) {
      const parent = (tnode as Text).parentElement;
      if (parent) {
        const style = window.getComputedStyle(parent);
        if (style.display !== 'none' && style.visibility !== 'hidden') {
          try {
            const fontStr = `${style.fontStyle} ${style.fontWeight} ${style.fontSize} ${style.fontFamily}`;
            ctx.save();
            ctx.font         = fontStr;
            ctx.fillStyle    = style.color || color;
            ctx.textBaseline = 'top';

            let drew = false;
            if (typeof document.createRange === 'function') {
              const range = document.createRange();
              range.selectNodeContents(tnode);
              const rects = range.getClientRects();
              if (rects && rects.length > 0 && rects[0].width > 0 && rects[0].height > 0) {
                ctx.fillText(text, rects[0].left - ox, rects[0].top - oy);
                drew = true;
              }
            }

            if (!drew) {
              const pRect = parent.getBoundingClientRect();
              if (pRect.width > 0 && pRect.height > 0) {
                ctx.fillText(text, pRect.left - ox, pRect.top - oy);
              } else {
                ctx.fillText(text, 0, 0);
              }
            }

            ctx.restore();
          } catch { /* skip this node */ }
        }
      }
    }
    tnode = walker.nextNode();
  }

  // ── Step 7: Draw inline SVG elements (sqrt sign, etc.) ───────────────────
  // ── Step 7: Draw inline SVG elements (sqrt sign, delimiters) via Path2D ───
  const svgEls = probe.querySelectorAll('svg');
  for (const svgEl of svgEls) {
    const rect = svgEl.getBoundingClientRect();
    const x = rect.left - ox;
    const y = rect.top  - oy;
    const w = rect.width;
    const h = rect.height;
    if (w <= 0 || h <= 0) continue;

    const pathEl = svgEl.querySelector('path');
    if (!pathEl) continue;

    const pathData = pathEl.getAttribute('d');
    if (!pathData) continue;

    try {
      if (typeof Path2D !== 'undefined') {
        // KaTeX virtual viewBox coordinate system (1000 units = 1em)
        const scaleRatio = fontSizePx / 1000;
        ctx.save();
        ctx.translate(x, y);
        ctx.scale(scaleRatio, scaleRatio);
        ctx.fillStyle = color;
        ctx.fill(new Path2D(pathData));
        ctx.restore();
      }
    } catch { /* ignore */ }
  }
  document.body.removeChild(probe);

  try {
    const dataUrl = canvas.toDataURL('image/png');
    return { dataUrl, width, height };
  } catch {
    return fallback;
  }
}

// ─── Public: generateStandaloneSvg (for taskpane live preview) ───────────────
/**
 * Returns a foreignObject SVG for the taskpane Live Preview only.
 * This is displayed inside the taskpane DOM where KaTeX CSS is active —
 * NOT passed to Excel shapes.addImage().
 */
export function generateStandaloneSvg(
  katexHtml: string,
  width: number,
  height: number,
  background: string | number,
  color: string,
  fontSizePx: number
): string {
  const isTransparent = !background || background === 'transparent' || background === 'none' || background === '0' || background === 0;
  const bgStyle = !isTransparent ? `background-color: ${background};` : '';
  return (
    `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}">` +
    `<foreignObject width="100%" height="100%">` +
    `<html xmlns="http://www.w3.org/1999/xhtml">` +
    `<head><style>html,body{margin:0;padding:0;height:100%;display:flex;` +
    `align-items:center;justify-content:center;` +
    `font-size:${fontSizePx}px;color: ${color};${bgStyle}}</style></head>` +
    `<body>${katexHtml}</body></html></foreignObject></svg>`
  );
}

/**
 * Converts a standalone SVG string to a PNG data URL (Base64).
 */
export async function svgToPngDataUrl(
  svg: string,
  width: number,
  height: number,
  scale: number = 2
): Promise<string> {
  if (typeof document === 'undefined') {
    return 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==';
  }
  const canvas = document.createElement('canvas');
  canvas.width = Math.max(1, Math.ceil(width * scale));
  canvas.height = Math.max(1, Math.ceil(height * scale));
  const ctx = canvas.getContext('2d');
  if (!ctx) return '';

  const cleanSvg = svg.includes('xmlns="http://www.w3.org/2000/svg"')
    ? svg
    : svg.replace('<svg', '<svg xmlns="http://www.w3.org/2000/svg"');

  const blob = new Blob([cleanSvg], { type: 'image/svg+xml;charset=utf-8' });
  const url = URL.createObjectURL(blob);

  return new Promise((resolve) => {
    const img = new Image();
    img.onload = () => {
      ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
      URL.revokeObjectURL(url);
      try {
        resolve(canvas.toDataURL('image/png'));
      } catch {
        resolve('');
      }
    };
    img.onerror = () => {
      URL.revokeObjectURL(url);
      try {
        resolve(canvas.toDataURL('image/png'));
      } catch {
        resolve('');
      }
    };
    img.src = url;
  });
}

// ─── Main Entry Point ─────────────────────────────────────────────────────────
export async function rasterizeLatex(
  latex: string,
  options: RenderOptions = {}
): Promise<RenderResult> {
  const cacheKey = createRenderCacheKey(latex, options);
  const cached   = equationCache.get(cacheKey);
  if (cached) return cached;

  const background  = options.background  || 'transparent';
  const color       = options.color       || '#000000';
  const fontSizePt  = options.fontSize    ?? 16;
  const fontSizePx  = ptToPx(fontSizePt);
  const displayMode = options.displayMode ?? true;
  const scale       = options.scale       ?? 2;

  const mergedMacros = { ...macroRegistry.getAll(), ...(options.macros || {}) };

  // 1. KaTeX → HTML (used for both preview and DOM rendering)
  let html = '';
  try {
    html = katex.renderToString(latex, {
      displayMode,
      throwOnError: false,
      errorColor: '#d13438',
      macros: mergedMacros,
      output: 'html',
      strict: false,
      trust: true,
    });
  } catch (err: any) {
    const msg = (err?.message || 'Syntax Error').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    html = `<span style="color:#d13438;font-family:monospace;">${msg}</span>`;
  }

  // 2. Render HTML → PNG via DOM-walk (8-second hard timeout)
  let pngDataUrl = '';
  let width  = 200;
  let height = 60;

  if (typeof document !== 'undefined') {
    try {
      const result = await Promise.race([
        renderKatexToCanvas(html, background, color, fontSizePx, scale),
        new Promise<{ dataUrl: string; width: number; height: number }>(
          r => setTimeout(() => r({ dataUrl: '', width: 200, height: 60 }), 8000)
        ),
      ]);
      pngDataUrl = result.dataUrl;
      width      = result.width;
      height     = result.height;
    } catch { /* use defaults */ }
  }

  // 3. Build display SVG for taskpane live preview
  const svg = generateStandaloneSvg(html, width, height, background, color, fontSizePx);

  const result: RenderResult = {
    html,
    svg,
    pngDataUrl,
    width,
    height,
    aspectRatio: width / Math.max(1, height),
  };

  equationCache.set(cacheKey, result);
  return result;
}

/**
 * Serializes equation options into JSON string for embedding in Alt Text and metadata.
 */
export function serializeEquationMetadata(latex: string, options?: RenderOptions): string {
  return JSON.stringify({
    latex: latex || '',
    bg: options?.background ?? 0,
    color: options?.color ?? '#000000',
    fontSize: options?.fontSize ?? 16,
    displayMode: options?.displayMode ?? true,
    v: 1
  });
}

/**
 * Parses equation metadata from Alt Text, JSON string, or standard LaTeX string.
 */
export function parseEquationMetadata(text?: string | null): {
  latex: string;
  background?: string | number;
  color?: string;
  fontSize?: number;
  displayMode?: boolean;
} | null {
  if (!text || typeof text !== 'string') return null;
  const trimmed = text.trim();
  if (!trimmed) return null;

  // Try parsing embedded JSON metadata
  if (trimmed.startsWith('{') && trimmed.endsWith('}')) {
    try {
      const parsed = JSON.parse(trimmed);
      if (parsed && typeof parsed.latex === 'string') {
        return {
          latex: parsed.latex,
          background: parsed.bg ?? parsed.background ?? 0,
          color: parsed.color ?? '#000000',
          fontSize: parsed.fontSize ?? parsed.size ?? 16,
          displayMode: parsed.displayMode ?? true
        };
      }
    } catch { /* fallback to standard parsing */ }
  }

  // Check for "LaTeX: <formula>" prefix
  if (trimmed.startsWith('LaTeX:')) {
    return {
      latex: trimmed.substring(6).trim(),
      background: 0
    };
  }

  // Standard raw LaTeX
  return {
    latex: trimmed,
    background: 0
  };
}
