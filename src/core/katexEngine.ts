/**
 * Core KaTeX Local Math Engine for ExcelKaTeX.
 * Provides safe LaTeX parsing, AST analysis, complexity scoring, MathML & HTML compilation,
 * and high-fidelity rendering delegation.
 */

import katex from 'katex';
import { macroRegistry, MacroDictionary } from './macros';
import {
  rasterizeLatex,
  generateStandaloneSvg,
  getEquationCacheStats,
  clearEquationCache,
  RenderOptions,
  RenderResult
} from './imageRasterizer';

export {
  rasterizeLatex,
  generateStandaloneSvg,
  getEquationCacheStats,
  clearEquationCache
};
export type { RenderOptions, RenderResult };

export interface KatexCompileOptions {
  displayMode?: boolean;
  throwOnError?: boolean;
  errorColor?: string;
  macros?: MacroDictionary;
  output?: 'html' | 'mathml' | 'htmlAndMathml';
  fleqn?: boolean;
  strict?: boolean | string | ((errorCode: string, errorMsg: string, token: any) => string | boolean);
  trust?: boolean | ((context: any) => boolean);
}

export type ComplexityLevel = 'Simple' | 'Moderate' | 'Complex' | 'Advanced';

export interface LatexComplexity {
  tokens: number;
  depth: number;
  score: ComplexityLevel;
}

export interface LatexValidationResult {
  isValid: boolean;
  error?: string;
  complexity?: LatexComplexity;
}

/**
 * Calculates structural complexity of a LaTeX math string.
 */
export function calculateComplexity(latex: string): LatexComplexity {
  if (!latex || !latex.trim()) {
    return { tokens: 0, depth: 0, score: 'Simple' };
  }

  const trimmed = latex.trim();
  let maxDepth = 0;
  let currentDepth = 0;

  // Track brace and environment depth
  for (let i = 0; i < trimmed.length; i++) {
    const char = trimmed[i];
    if (char === '{' || char === '[' || char === '(') {
      currentDepth++;
      if (currentDepth > maxDepth) maxDepth = currentDepth;
    } else if (char === '}' || char === ']' || char === ')') {
      if (currentDepth > 0) currentDepth--;
    }
  }

  // Count environments
  const envMatches = trimmed.match(/\\begin\{[^}]+\}/g);
  const envCount = envMatches ? envMatches.length : 0;
  maxDepth += envCount * 2;

  const fractions = (trimmed.match(/\\(frac|dfrac|tfrac|cfrac)/g) || []).length;
  const roots = (trimmed.match(/\\sqrt/g) || []).length;
  const integrals = (trimmed.match(/\\(int|iint|iiint|oint|intinf)/g) || []).length;
  const bigOps = (trimmed.match(/\\(sum|prod|suminf|lim|limto|bigcup|bigcap)/g) || []).length;
  const matrices = (trimmed.match(/\\(matrix|pmatrix|bmatrix|vmatrix|cases|align|gather)/g) || []).length;

  // Approximate token count (commands, numbers, variables, symbols)
  const tokens = trimmed
    .split(/(\\[a-zA-Z]+|[{}\[\]()^_=+*/\-]|[\s,;]+)/)
    .filter(t => t && t.trim().length > 0).length;

  let score: ComplexityLevel = 'Simple';
  if (
    tokens > 25 ||
    maxDepth >= 4 ||
    envCount >= 2 ||
    (integrals > 0 && (fractions > 0 || roots > 0)) ||
    (fractions >= 2 && maxDepth >= 2) ||
    (bigOps > 0 && fractions > 0 && maxDepth >= 2)
  ) {
    score = 'Advanced';
  } else if (
    tokens > 15 ||
    maxDepth >= 2 ||
    envCount >= 1 ||
    matrices > 0 ||
    integrals > 0 ||
    (fractions > 0 && maxDepth >= 1)
  ) {
    score = 'Complex';
  } else if (tokens > 8 || maxDepth >= 1 || fractions > 0 || roots > 0 || bigOps > 0) {
    score = 'Moderate';
  }

  return {
    tokens,
    depth: maxDepth,
    score
  };
}

/**
 * Validates LaTeX syntax using KaTeX without throwing.
 */
export function validateLatex(
  latex: string,
  customMacros?: MacroDictionary
): LatexValidationResult {
  if (!latex || typeof latex !== 'string' || !latex.trim()) {
    return {
      isValid: false,
      error: 'LaTeX equation string is empty.'
    };
  }

  const trimmed = latex.trim();

  // Guard against HTML injection and script tags
  if (/<[a-zA-Z\/][^>]*>|javascript:|alert\(|onload=|onerror=/i.test(trimmed)) {
    return {
      isValid: false,
      error: 'KaTeX parse error: HTML tags and script injections are forbidden.'
    };
  }

  const mergedMacros = {
    ...macroRegistry.getAll(),
    ...(customMacros || {})
  };

  try {
    katex.renderToString(trimmed, {
      throwOnError: true,
      displayMode: true,
      macros: mergedMacros,
      strict: false,
      trust: true
    });

    const complexity = calculateComplexity(trimmed);
    return {
      isValid: true,
      complexity
    };
  } catch (err: any) {
    let errMsg = err?.message || 'KaTeX parse error: Invalid LaTeX syntax.';
    if (!errMsg.includes('KaTeX parse error')) {
      errMsg = `KaTeX parse error: ${errMsg}`;
    }
    return {
      isValid: false,
      error: errMsg
    };
  }
}

/**
 * Compiles LaTeX string directly to HTML string.
 */
export function compileLatexToHtml(
  latex: string,
  options: KatexCompileOptions = {}
): string {
  if (!latex || !latex.trim()) {
    return '';
  }

  const mergedMacros = {
    ...macroRegistry.getAll(),
    ...(options.macros || {})
  };

  const defaultOptions: katex.KatexOptions = {
    displayMode: options.displayMode ?? true,
    throwOnError: true,
    errorColor: options.errorColor ?? '#d13438',
    macros: mergedMacros,
    output: options.output ?? 'htmlAndMathml',
    strict: (options.strict as any) ?? false,
    trust: options.trust ?? false
  };

  try {
    return katex.renderToString(latex, defaultOptions);
  } catch (err: any) {
    if (options.throwOnError) {
      throw err;
    }
    const safeError = (err?.message || 'KaTeX parse error')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;');
    return `<span class="katex-error" style="color:${options.errorColor ?? '#d13438'};font-weight:bold;">KaTeX parse error: ${safeError}</span>`;
  }
}

/**
 * Compiles LaTeX string to MathML string.
 */
export function compileLatexToMathML(
  latex: string,
  options: KatexCompileOptions = {}
): string {
  return compileLatexToHtml(latex, {
    ...options,
    output: 'mathml'
  });
}

/**
 * Primary equation compiler returning full RenderResult with HTML, Standalone SVG, and High-DPI PNG.
 */
export async function compileLatex(
  latex: string,
  options: RenderOptions = {}
): Promise<RenderResult> {
  return rasterizeLatex(latex, options);
}
