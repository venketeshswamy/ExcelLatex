/**
 * Parameter Parser & Validator for ExcelKaTeX Custom Functions.
 * Validates and sanitizes parameters for =MATH.KATEX(latexString, background, color, fontSize, displayMode, outputMethod).
 */

import { validateLatex } from '../core/katexEngine';

export type OutputMethod = 'auto' | 'cell' | 'shape' | 'image' | 'entity';

export interface MathKatexInputParams {
  latexString: string;
  background?: any;
  color?: any;
  fontSize?: any;
  displayMode?: any;
  outputMethod?: OutputMethod | any;
  invocation?: any;
}

export interface ValidatedMathKatexParams {
  latexString: string;
  background: string;
  color: string;
  fontSize: number;
  displayMode: boolean;
  outputMethod: OutputMethod;
  cellAddress: string;
}

export interface ValidationOutcome {
  isValid: boolean;
  params?: ValidatedMathKatexParams;
  error?: string;
  errorCode?: string;
}

export const DEFAULT_BACKGROUND = 'transparent';
export const DEFAULT_COLOR = '#000000';
export const DEFAULT_FONT_SIZE = 16;
export const DEFAULT_DISPLAY_MODE = true;
export const DEFAULT_OUTPUT_METHOD: OutputMethod = 'auto';
export const DEFAULT_CELL_ADDRESS = 'Sheet1!A1';

/**
 * Checks if a given value is an Excel Custom Functions Invocation object.
 */
export function isInvocationObject(arg: any): boolean {
  if (!arg || typeof arg !== 'object' || Array.isArray(arg)) {
    return false;
  }
  return (
    'address' in arg ||
    'parameterAddresses' in arg ||
    'functionName' in arg ||
    'onCanceled' in arg ||
    Object.prototype.toString.call(arg) === '[object Object]'
  );
}

/**
 * Determines whether a color is dark (used for adaptive text color).
 */
export function isDarkColor(color: string): boolean {
  if (!color || typeof color !== 'string') return false;
  const lower = color.trim().toLowerCase();
  if (lower === 'transparent' || lower === 'currentcolor') return false;
  if (lower === '#000000' || lower === '#000' || lower === 'black') return true;

  // 6-digit or 8-digit hex (#rrggbb or #rrggbbaa)
  const hex6Match = /^#([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})/i.exec(lower);
  if (hex6Match) {
    const r = parseInt(hex6Match[1], 16);
    const g = parseInt(hex6Match[2], 16);
    const b = parseInt(hex6Match[3], 16);
    const luminance = 0.299 * r + 0.587 * g + 0.114 * b;
    return luminance < 128;
  }

  // 3-digit hex (#rgb)
  const hex3Match = /^#([0-9a-f])([0-9a-f])([0-9a-f])/i.exec(lower);
  if (hex3Match) {
    const r = parseInt(hex3Match[1] + hex3Match[1], 16);
    const g = parseInt(hex3Match[2] + hex3Match[2], 16);
    const b = parseInt(hex3Match[3] + hex3Match[3], 16);
    const luminance = 0.299 * r + 0.587 * g + 0.114 * b;
    return luminance < 128;
  }

  // rgb / rgba
  const rgbMatch = /^rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)/i.exec(lower);
  if (rgbMatch) {
    const r = parseInt(rgbMatch[1], 10);
    const g = parseInt(rgbMatch[2], 10);
    const b = parseInt(rgbMatch[3], 10);
    const luminance = 0.299 * r + 0.587 * g + 0.114 * b;
    return luminance < 128;
  }

  if (['darkblue', 'navy', 'darkgreen', 'darkred', 'purple', 'midnightblue', 'maroon'].includes(lower)) {
    return true;
  }

  return false;
}

/**
 * Returns default text color based on background (adaptive color).
 * Black/dark backgrounds -> '#ffffff', light/transparent -> '#000000'.
 */
export function getDefaultTextColor(resolvedBackground: string): string {
  return isDarkColor(resolvedBackground) ? '#ffffff' : DEFAULT_COLOR;
}

/**
 * Standardizes background parameter (0=Transparent, 1=White, 2=Black, or CSS color).
 */
export function sanitizeBackground(background?: any): string {
  if (background === undefined || background === null) {
    return DEFAULT_BACKGROUND;
  }

  if (typeof background === 'number') {
    if (background === 0) return 'transparent';
    if (background === 1) return '#ffffff';
    if (background === 2) return '#000000';
    return DEFAULT_BACKGROUND;
  }

  if (typeof background === 'string') {
    const trimmed = background.trim();
    if (!trimmed) return DEFAULT_BACKGROUND;
    const lower = trimmed.toLowerCase();

    if (lower === '0' || lower === 'transparent') return 'transparent';
    if (lower === '1' || lower === 'white') return '#ffffff';
    if (lower === '2' || lower === 'black') return '#000000';

    if (
      /^(#[0-9a-fA-F]{3,8}|rgba?\([^)]+\)|hsla?\([^)]+\)|[a-zA-Z]+|currentColor)$/.test(
        trimmed
      )
    ) {
      return trimmed;
    }
    return DEFAULT_BACKGROUND;
  }

  return DEFAULT_BACKGROUND;
}

/**
 * Sanitizes foreground text color with adaptive defaults based on resolved background.
 */
export function sanitizeTextColor(color?: any, resolvedBackground = DEFAULT_BACKGROUND): string {
  const defaultText = getDefaultTextColor(resolvedBackground);
  if (color === undefined || color === null) {
    return defaultText;
  }

  if (typeof color === 'string') {
    const trimmed = color.trim();
    if (!trimmed) {
      return defaultText;
    }
    if (
      /^(#[0-9a-fA-F]{3,8}|rgba?\([^)]+\)|hsla?\([^)]+\)|[a-zA-Z]+|transparent|currentColor)$/.test(
        trimmed
      )
    ) {
      return trimmed;
    }
  }

  return defaultText;
}

/**
 * Normalizes and sanitizes general color strings (kept for compatibility).
 */
export function sanitizeColor(color: string | undefined | any, defaultColor: string): string {
  if (!color || typeof color !== 'string' || !color.trim()) {
    return defaultColor;
  }
  const trimmed = color.trim();
  if (
    /^(#[0-9a-fA-F]{3,8}|rgba?\([^)]+\)|hsla?\([^)]+\)|[a-zA-Z]+|transparent|currentColor)$/.test(
      trimmed
    )
  ) {
    return trimmed;
  }
  return defaultColor;
}

/**
 * Normalizes font size with bounds clamping.
 */
export function sanitizeFontSize(fontSize: any): number {
  if (fontSize === undefined || fontSize === null || fontSize === '') {
    return DEFAULT_FONT_SIZE;
  }
  const num = typeof fontSize === 'number' ? fontSize : Number(fontSize);
  if (isNaN(num) || num <= 0) {
    return DEFAULT_FONT_SIZE;
  }
  // Clamp between 4 and 1000 pt
  return Math.min(1000, Math.max(4, num));
}

/**
 * Normalizes displayMode parameter.
 */
export function sanitizeDisplayMode(displayMode: any): boolean {
  if (displayMode === undefined || displayMode === null || displayMode === '') {
    return DEFAULT_DISPLAY_MODE;
  }
  if (typeof displayMode === 'boolean') {
    return displayMode;
  }
  if (typeof displayMode === 'number') {
    return displayMode !== 0;
  }
  if (typeof displayMode === 'string') {
    const lower = displayMode.trim().toLowerCase();
    if (lower === 'true' || lower === '1') return true;
    if (lower === 'false' || lower === '0') return false;
    if (!lower) return DEFAULT_DISPLAY_MODE;
  }
  return Boolean(displayMode);
}

/**
 * Normalizes output method parameter.
 */
export function sanitizeOutputMethod(outputMethod: any): OutputMethod {
  if (!outputMethod || typeof outputMethod !== 'string') {
    return DEFAULT_OUTPUT_METHOD;
  }
  const lower = outputMethod.trim().toLowerCase();
  if (lower === 'cell' || lower === 'shape' || lower === 'auto' || lower === 'entity') {
    return lower as OutputMethod;
  }
  if (lower === 'image' || lower === 'img' || lower === 'webimage') {
    return 'image';
  }
  return DEFAULT_OUTPUT_METHOD;
}

/**
 * Extracts normalized cell address from custom function invocation object.
 */
export function extractCellAddress(invocation?: any): string {
  if (!invocation) {
    return DEFAULT_CELL_ADDRESS;
  }
  if (typeof invocation === 'string' && invocation.trim()) {
    return invocation.trim();
  }
  if (typeof invocation === 'object' && invocation.address && typeof invocation.address === 'string') {
    return invocation.address.trim();
  }
  return DEFAULT_CELL_ADDRESS;
}

/**
 * Strips wrapping LaTeX math delimiters ($$, \[\], \(\), $) and trims whitespace.
 */
export function stripLatexDelimiters(latex: string): string {
  if (!latex || typeof latex !== 'string') return '';
  let trimmed = latex.trim();

  // $$ ... $$
  if (trimmed.startsWith('$$') && trimmed.endsWith('$$') && trimmed.length >= 4) {
    trimmed = trimmed.substring(2, trimmed.length - 2).trim();
  }
  // \[ ... \]
  else if (trimmed.startsWith('\\[') && trimmed.endsWith('\\]') && trimmed.length >= 4) {
    trimmed = trimmed.substring(2, trimmed.length - 2).trim();
  }
  // \( ... \)
  else if (trimmed.startsWith('\\(') && trimmed.endsWith('\\)') && trimmed.length >= 4) {
    trimmed = trimmed.substring(2, trimmed.length - 2).trim();
  }
  // $ ... $ (single delimiter)
  else if (trimmed.startsWith('$') && trimmed.endsWith('$') && trimmed.length >= 2 && !trimmed.startsWith('$$')) {
    trimmed = trimmed.substring(1, trimmed.length - 1).trim();
  }

  return trimmed;
}

/**
 * Parses and validates all parameters for =MATH.KATEX custom function.
 */
export function parseMathKatexParams(
  latexString: string,
  background?: any,
  color?: any,
  fontSize?: any,
  displayMode?: any,
  outputMethod?: any,
  invocation?: any
): ValidationOutcome {
  // Resilient Invocation Extraction across all parameter slots (1-6)
  let actualInvocation: any = invocation;
  let rawBackground = background;
  let rawColor = color;
  let rawFontSize = fontSize;
  let rawDisplayMode = displayMode;
  let rawOutputMethod = outputMethod;

  if (isInvocationObject(rawBackground)) {
    actualInvocation = rawBackground;
    rawBackground = undefined;
  } else if (isInvocationObject(rawColor)) {
    actualInvocation = rawColor;
    rawColor = undefined;
  } else if (isInvocationObject(rawFontSize)) {
    actualInvocation = rawFontSize;
    rawFontSize = undefined;
  } else if (isInvocationObject(rawDisplayMode)) {
    actualInvocation = rawDisplayMode;
    rawDisplayMode = undefined;
  } else if (isInvocationObject(rawOutputMethod)) {
    actualInvocation = rawOutputMethod;
    rawOutputMethod = undefined;
  }

  // 1. Validate LaTeX input presence and type
  if (latexString === undefined || latexString === null || typeof latexString !== 'string') {
    return {
      isValid: false,
      error: 'LaTeX equation string is required and must be a string.',
      errorCode: '#VALUE!'
    };
  }

  const strippedLatex = stripLatexDelimiters(latexString);
  if (!strippedLatex) {
    return {
      isValid: false,
      error: 'LaTeX equation string cannot be empty.',
      errorCode: '#VALUE!'
    };
  }

  // 2. Validate LaTeX syntax using KaTeX engine
  const syntaxValidation = validateLatex(strippedLatex);
  if (!syntaxValidation.isValid) {
    return {
      isValid: false,
      error: syntaxValidation.error || 'Invalid LaTeX syntax.',
      errorCode: '#VALUE!'
    };
  }

  // 3. Sanitize formatting options with adaptive color defaults
  const sanitizedBg = sanitizeBackground(rawBackground);
  const sanitizedFg = sanitizeTextColor(rawColor, sanitizedBg);
  const sanitizedSize = sanitizeFontSize(rawFontSize);
  const sanitizedDisplay = sanitizeDisplayMode(rawDisplayMode);
  const sanitizedMethod = sanitizeOutputMethod(rawOutputMethod);
  const sanitizedAddress = extractCellAddress(actualInvocation);

  return {
    isValid: true,
    params: {
      latexString: strippedLatex,
      background: sanitizedBg,
      color: sanitizedFg,
      fontSize: sanitizedSize,
      displayMode: sanitizedDisplay,
      outputMethod: sanitizedMethod,
      cellAddress: sanitizedAddress
    }
  };
}
