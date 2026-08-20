import { compileLatex, compileLatexToNativeMathML, RenderOptions } from '../../core/katexEngine';

export interface ClipboardService {
  copyLatex(latex: string): Promise<boolean>;
  copyMathML(latex: string): Promise<boolean>;
  copySvg(svgString: string): Promise<boolean>;
  copyPng(pngDataUrl: string): Promise<boolean>;
  copyEquation(latex: string, options?: RenderOptions): Promise<{ latex: boolean; mathml: boolean; svg: boolean; png: boolean }>;
}

/**
 * Converts a base64 or data URL string to a Blob.
 */
export function dataUrlToBlob(dataUrl: string): Blob {
  const parts = dataUrl.split(',');
  const mimeMatch = parts[0].match(/:(.*?);/);
  const mime = mimeMatch ? mimeMatch[1] : 'image/png';
  const binary = atob(parts[1]);
  const array = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) {
    array[i] = binary.charCodeAt(i);
  }
  return new Blob([array], { type: mime });
}

/**
 * Fallback text copy using temporary textarea and document.execCommand('copy').
 */
function fallbackCopyText(text: string): boolean {
  if (typeof document === 'undefined') return false;
  try {
    const textarea = document.createElement('textarea');
    textarea.value = text;
    textarea.style.position = 'fixed';
    textarea.style.left = '-9999px';
    textarea.style.top = '-9999px';
    textarea.setAttribute('readonly', '');
    document.body.appendChild(textarea);
    textarea.select();
    const successful = document.execCommand('copy');
    document.body.removeChild(textarea);
    return successful;
  } catch {
    return false;
  }
}

/**
 * Copies plain text (e.g. LaTeX code) to the system clipboard.
 */
export async function copyLatex(latex: string): Promise<boolean> {
  if (!latex) return false;
  try {
    if (typeof navigator !== 'undefined' && navigator.clipboard && navigator.clipboard.writeText) {
      await navigator.clipboard.writeText(latex);
      return true;
    }
    return fallbackCopyText(latex);
  } catch {
    return fallbackCopyText(latex);
  }
}

/**
 * Copies native MathML XML directly to the system clipboard for Microsoft Word equation pasting.
 */
export async function copyMathML(latex: string): Promise<boolean> {
  if (!latex) return false;
  try {
    const mathml = compileLatexToNativeMathML(latex);
    if (typeof navigator !== 'undefined' && navigator.clipboard && navigator.clipboard.writeText) {
      await navigator.clipboard.writeText(mathml);
      return true;
    }
    return fallbackCopyText(mathml);
  } catch {
    return false;
  }
}

/**
 * Copies SVG markup text to the system clipboard.
 */
export async function copySvg(svgString: string): Promise<boolean> {
  if (!svgString) return false;
  return copyLatex(svgString);
}

/**
 * Copies PNG image data to the system clipboard.
 */
export async function copyPng(pngDataUrl: string): Promise<boolean> {
  if (!pngDataUrl) return false;
  try {
    if (
      typeof navigator !== 'undefined' &&
      navigator.clipboard &&
      typeof (window as any).ClipboardItem !== 'undefined'
    ) {
      const blob = dataUrlToBlob(pngDataUrl);
      const item = new (window as any).ClipboardItem({ [blob.type]: blob });
      await navigator.clipboard.write([item]);
      return true;
    }
    return copyLatex(pngDataUrl);
  } catch {
    return copyLatex(pngDataUrl);
  }
}

/**
 * Compiles equation and copies rendered formats.
 */
export async function copyEquation(
  latex: string,
  options: RenderOptions = {}
): Promise<{ latex: boolean; mathml: boolean; svg: boolean; png: boolean }> {
  const latexSuccess = await copyLatex(latex);
  const mathmlSuccess = await copyMathML(latex);
  let svgSuccess = false;
  let pngSuccess = false;

  try {
    const render = await compileLatex(latex, options);
    svgSuccess = await copySvg(render.svg);
    pngSuccess = await copyPng(render.pngDataUrl);
  } catch (err) {
    console.warn('[ExcelKaTeX] Failed to compile equation for clipboard copy:', err);
  }

  return {
    latex: latexSuccess,
    mathml: mathmlSuccess,
    svg: svgSuccess,
    png: pngSuccess
  };
}

export const clipboardService: ClipboardService = {
  copyLatex,
  copyMathML,
  copySvg,
  copyPng,
  copyEquation
};
