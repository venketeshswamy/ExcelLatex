/**
 * ExcelKaTeX Main Entry Point.
 * Exports core math compiler, image rasterizer, cache, macros, and initializes runtime.
 */

export * from './core/macros';
export * from './core/lruCache';
export * from './core/katexEngine';
export * from './core/imageRasterizer';
export * from './customfunctions';

// Ensure Office.js custom functions runtime is bound if available
if (typeof CustomFunctions !== 'undefined') {
  console.log('[ExcelKaTeX] CustomFunctions runtime detected.');
}

