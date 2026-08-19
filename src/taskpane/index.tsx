import React from 'react';
import { createRoot } from 'react-dom/client';
import 'mathlive';
import { MathfieldElement } from 'mathlive';
import 'katex/dist/katex.min.css';
import { App } from './App';
import { mathKatex } from '../customfunctions/mathKatex';
import { installKatexShapeAutoRenderer } from './services/excelService';

// Associate Custom Functions in SharedRuntime (taskpane window)
if (
  typeof (globalThis as any).CustomFunctions !== 'undefined' &&
  typeof (globalThis as any).CustomFunctions.associate === 'function'
) {
  try {
    (globalThis as any).CustomFunctions.associate('KATEX', mathKatex);
    (globalThis as any).CustomFunctions.associate('MATH.KATEX', mathKatex);
  } catch (err) {
    console.error('[ExcelKaTeX] CustomFunctions associate error:', err);
  }
}

// Zero-CDN Offline MathLive fonts & sounds configuration
if (typeof window !== 'undefined') {
  if (typeof MathfieldElement !== 'undefined') {
    MathfieldElement.fontsDirectory = './assets/mathlive-fonts/';
    MathfieldElement.soundsDirectory = null;
  }
}

function mountTaskpane() {
  const container = document.getElementById('root');
  if (container) {
    const root = createRoot(container);
    root.render(
      <React.StrictMode>
        <App />
      </React.StrictMode>
    );
  }
}

// Initialize on Office.js readiness or immediately in browser dev mode
if (typeof Office !== 'undefined') {
  Office.onReady(() => {
    mountTaskpane();
    // Install the onChanged listener that auto-converts "📐 <latex>" cells to images
    installKatexShapeAutoRenderer().catch((err) => {
      console.warn('[ExcelKaTeX] Auto-renderer install failed:', err);
    });
  });
} else if (typeof document !== 'undefined') {
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', mountTaskpane);
  } else {
    mountTaskpane();
  }
}

export { App };
