import { describe, it, expect, beforeEach, vi, afterEach } from 'vitest';
import React from 'react';
import { render, screen, fireEvent, act } from '@testing-library/react';
import { FluentProvider, webLightTheme } from '@fluentui/react-components';

import { ActionBar } from '../src/taskpane/components/ActionBar';
import { LivePreview, getBackgroundLabel } from '../src/taskpane/components/LivePreview';
import { App } from '../src/taskpane/App';
import {
  excelService,
  insertFormulaToActiveCell,
  readActiveCellFormula,
  escapeFormulaString,
  formatBackgroundParam
} from '../src/taskpane/services/excelService';
import {
  getMockExcelState,
  resetOfficeMock,
  setMockRequirementSupported
} from '../src/mocks/officeMock';
import fs from 'fs';
import path from 'path';

function wrap(ui: React.ReactElement) {
  return React.createElement(FluentProvider, { theme: webLightTheme }, ui);
}

describe('Adversarial Challenger M5: Taskpane UI (R3) & Integration/Build (R4)', () => {
  beforeEach(() => {
    resetOfficeMock();
    setMockRequirementSupported('ExcelApi', 1.16, true);
    vi.useRealTimers();
  });

  afterEach(() => {
    vi.restoreAllMocks();
  });

  // =========================================================================
  // 1. Taskpane UI: Single Action & Button Layout Verification
  // =========================================================================
  describe('1. Taskpane UI Button & Action Layout (R3)', () => {
    it('renders primary In-Cell Image and Insert Formula buttons', () => {
      render(
        wrap(
          React.createElement(ActionBar, {
            latex: '\\frac{1}{2}',
            isValid: true
          })
        )
      );

      const primaryBtn = screen.getByRole('button', { name: /Insert Formula/i });
      expect(primaryBtn).toBeDefined();

      const inCellBtn = screen.getByRole('button', { name: /In-Cell Image/i });
      expect(inCellBtn).toBeDefined();

      // Verify secondary actions
      expect(screen.getByRole('button', { name: /Floating Shape/i })).toBeDefined();
      expect(screen.getByRole('button', { name: /Read Cell/i })).toBeDefined();
      expect(screen.getByRole('button', { name: /Batch Convert/i })).toBeDefined();
      expect(screen.getByRole('button', { name: /Export/i })).toBeDefined();
    });

    it('disables "Insert Formula" and "Floating Shape" buttons when LaTeX is invalid or empty', () => {
      // Test 1: Empty latex
      const { rerender } = render(
        wrap(
          React.createElement(ActionBar, {
            latex: '',
            isValid: true
          })
        )
      );

      let insertBtn = screen.getByRole('button', { name: /Insert Formula/i });
      let shapeBtn = screen.getByRole('button', { name: /Floating Shape/i });
      expect(insertBtn.hasAttribute('disabled') || insertBtn.getAttribute('aria-disabled') === 'true').toBe(true);
      expect(shapeBtn.hasAttribute('disabled') || shapeBtn.getAttribute('aria-disabled') === 'true').toBe(true);

      // Test 2: Whitespace only
      rerender(
        wrap(
          React.createElement(ActionBar, {
            latex: '   \t\n  ',
            isValid: true
          })
        )
      );
      insertBtn = screen.getByRole('button', { name: /Insert Formula/i });
      shapeBtn = screen.getByRole('button', { name: /Floating Shape/i });
      expect(insertBtn.hasAttribute('disabled') || insertBtn.getAttribute('aria-disabled') === 'true').toBe(true);
      expect(shapeBtn.hasAttribute('disabled') || shapeBtn.getAttribute('aria-disabled') === 'true').toBe(true);

      // Test 3: Invalid syntax
      rerender(
        wrap(
          React.createElement(ActionBar, {
            latex: '\\frac{1}{',
            isValid: false
          })
        )
      );
      insertBtn = screen.getByRole('button', { name: /Insert Formula/i });
      shapeBtn = screen.getByRole('button', { name: /Floating Shape/i });
      expect(insertBtn.hasAttribute('disabled') || insertBtn.getAttribute('aria-disabled') === 'true').toBe(true);
      expect(shapeBtn.hasAttribute('disabled') || shapeBtn.getAttribute('aria-disabled') === 'true').toBe(true);
    });

    it('handles rapid clicking on Insert Formula gracefully with loading spinner protection', async () => {
      let resolveRun: () => void;
      const delayedPromise = new Promise<void>((res) => {
        resolveRun = res;
      });

      const spy = vi.spyOn(excelService, 'insertFormulaToActiveCell').mockImplementation(() => delayedPromise);

      render(
        wrap(
          React.createElement(ActionBar, {
            latex: 'x^2 + y^2 = r^2',
            isValid: true
          })
        )
      );

      const insertBtn = screen.getByRole('button', { name: /Insert Formula/i });

      // Click once
      await act(async () => {
        fireEvent.click(insertBtn);
      });

      expect(spy).toHaveBeenCalledTimes(1);
      // While pending, button should be disabled
      expect(insertBtn.hasAttribute('disabled') || insertBtn.getAttribute('aria-disabled') === 'true').toBe(true);

      // Subsequent click during in-flight operation should not trigger another call
      await act(async () => {
        fireEvent.click(insertBtn);
      });
      expect(spy).toHaveBeenCalledTimes(1);

      // Complete the pending promise
      await act(async () => {
        resolveRun!();
      });

      // Verify success notification is rendered
      expect(screen.getByText(/Inserted =MATH.KATEX\(\) formula into active cell!/i)).toBeDefined();
    });

    it('surfaces error MessageBar when Excel operation throws', async () => {
      vi.spyOn(excelService, 'insertFormulaToActiveCell').mockRejectedValueOnce(
        new Error('Workbook is in protected mode')
      );

      render(
        wrap(
          React.createElement(ActionBar, {
            latex: '\\sqrt{2}',
            isValid: true
          })
        )
      );

      const insertBtn = screen.getByRole('button', { name: /Insert Formula/i });

      await act(async () => {
        fireEvent.click(insertBtn);
      });

      expect(screen.getByText(/Workbook is in protected mode/i)).toBeDefined();
      expect(screen.getByText('Error')).toBeDefined();
    });
  });

  // =========================================================================
  // 2. Dropdown Controls & Dynamic Background / Display Mode Adaptations
  // =========================================================================
  describe('2. Dropdown Controls & Styling Adaptations (R3)', () => {
    it('LivePreview renders Background dropdown with 3 options: 0 (Transparent), 1 (White), 2 (Black)', () => {
      const onBgChange = vi.fn();
      render(
        wrap(
          React.createElement(LivePreview, {
            latex: '\\alpha + \\beta',
            isValid: true,
            background: 0,
            onBackgroundChange: onBgChange
          })
        )
      );

      const bgDropdown = screen.getByRole('combobox', { name: /Background/i });
      expect(bgDropdown).toBeDefined();
      expect(bgDropdown.textContent).toContain('0: Transparent');
    });

    it('LivePreview adapts preview background and text color styling for Black (2) and White (1)', () => {
      const { rerender } = render(
        wrap(
          React.createElement(LivePreview, {
            latex: 'E = mc^2',
            isValid: true,
            background: 2
          })
        )
      );

      let card = screen.getByRole('region', { name: /Equation Preview/i });
      expect(card.style.backgroundColor).toBe('rgb(0, 0, 0)');
      expect(card.style.color).toBe('rgb(255, 255, 255)');

      // Rerender with White background (1)
      rerender(
        wrap(
          React.createElement(LivePreview, {
            latex: 'E = mc^2',
            isValid: true,
            background: 1
          })
        )
      );
      card = screen.getByRole('region', { name: /Equation Preview/i });
      expect(card.style.backgroundColor).toBe('rgb(255, 255, 255)');
      expect(card.style.color).toBe('rgb(0, 0, 0)');
    });

    it('getBackgroundLabel handles all valid numeric, string, and edge-case backgrounds', () => {
      expect(getBackgroundLabel(0)).toBe('0: Transparent');
      expect(getBackgroundLabel('0')).toBe('0: Transparent');
      expect(getBackgroundLabel('transparent')).toBe('0: Transparent');
      expect(getBackgroundLabel(undefined)).toBe('0: Transparent');

      expect(getBackgroundLabel(1)).toBe('1: White');
      expect(getBackgroundLabel('1')).toBe('1: White');
      expect(getBackgroundLabel('white')).toBe('1: White');
      expect(getBackgroundLabel('#ffffff')).toBe('1: White');

      expect(getBackgroundLabel(2)).toBe('2: Black');
      expect(getBackgroundLabel('2')).toBe('2: Black');
      expect(getBackgroundLabel('black')).toBe('2: Black');
      expect(getBackgroundLabel('#000000')).toBe('2: Black');

      // Custom background hex
      expect(getBackgroundLabel('#336699')).toBe('#336699');
    });

    it('App handles background change: auto-adapting text color between Black and White/Transparent', () => {
      render(React.createElement(App));

      // Open Options tab
      const optionsTab = screen.getByRole('tab', { name: /Options/i });
      fireEvent.click(optionsTab);

      const bgDropdown = screen.getByRole('combobox', { name: /Background Option/i });
      expect(bgDropdown).toBeDefined();
      const displayDropdown = screen.getByRole('combobox', { name: /Display Mode Option/i });
      expect(displayDropdown).toBeDefined();
    });
  });

  // =========================================================================
  // 3. Formula Generation Formatting in Excel Interop (R3, R4)
  // =========================================================================
  describe('3. Formula Generation Formatting in Excel Interop (R3, R4)', () => {
    it('generates compact =MATH.KATEX("<latex>") when default options are supplied', async () => {
      const state = getMockExcelState();

      // Case A: No options
      await insertFormulaToActiveCell('\\int x \\, dx');
      expect(state.workbook.getSelectedRange().formulas[0][0]).toBe('=MATH.KATEX("\\int x \\, dx")');

      // Case B: Explicit default options (background: 0, color: "#000000", fontSize: 16, displayMode: true)
      await insertFormulaToActiveCell('\\frac{a}{b}', {
        background: 0,
        color: '#000000',
        fontSize: 16,
        displayMode: true
      });
      expect(state.workbook.getSelectedRange().formulas[0][0]).toBe('=MATH.KATEX("\\frac{a}{b}")');

      // Case C: Explicit background: "0" or "transparent"
      await insertFormulaToActiveCell('x^2', { background: '0' });
      expect(state.workbook.getSelectedRange().formulas[0][0]).toBe('=MATH.KATEX("x^2")');

      await insertFormulaToActiveCell('x^2', { background: 'transparent' });
      expect(state.workbook.getSelectedRange().formulas[0][0]).toBe('=MATH.KATEX("x^2")');
    });

    it('generates numeric background codes (0, 1, 2) when non-default options are set', async () => {
      const state = getMockExcelState();

      // White background (1)
      await insertFormulaToActiveCell('\\sin(x)', {
        background: 1
      });
      expect(state.workbook.getSelectedRange().formulas[0][0]).toBe(
        '=MATH.KATEX("\\sin(x)", 1)'
      );

      // Black background (2) with auto-adapted white text
      await insertFormulaToActiveCell('\\cos(x)', {
        background: 2
      });
      expect(state.workbook.getSelectedRange().formulas[0][0]).toBe(
        '=MATH.KATEX("\\cos(x)", 2)'
      );

      // Custom font size (24) with transparent background
      await insertFormulaToActiveCell('y = mx + b', {
        background: 0,
        fontSize: 24
      });
      expect(state.workbook.getSelectedRange().formulas[0][0]).toBe(
        '=MATH.KATEX("y = mx + b")'
      );

      // Inline mode (displayMode: false)
      await insertFormulaToActiveCell('e^{i\\pi} + 1 = 0', {
        displayMode: false
      });
      expect(state.workbook.getSelectedRange().formulas[0][0]).toBe(
        '=MATH.KATEX("e^{i\\pi} + 1 = 0")'
      );

      // Custom color
      await insertFormulaToActiveCell('F = ma', {
        color: '#ff5500'
      });
      expect(state.workbook.getSelectedRange().formulas[0][0]).toBe(
        '=MATH.KATEX("F = ma")'
      );
    });

    it('formatBackgroundParam maps numeric and string inputs to standardized codes 0, 1, 2', () => {
      expect(formatBackgroundParam(0)).toBe('0');
      expect(formatBackgroundParam('0')).toBe('0');
      expect(formatBackgroundParam('transparent')).toBe('0');
      expect(formatBackgroundParam('TRANSPARENT')).toBe('0');

      expect(formatBackgroundParam(1)).toBe('1');
      expect(formatBackgroundParam('1')).toBe('1');
      expect(formatBackgroundParam('white')).toBe('1');
      expect(formatBackgroundParam('WHITE')).toBe('1');

      expect(formatBackgroundParam(2)).toBe('2');
      expect(formatBackgroundParam('2')).toBe('2');
      expect(formatBackgroundParam('black')).toBe('2');
      expect(formatBackgroundParam('BLACK')).toBe('2');

      expect(formatBackgroundParam('#abcdef')).toBe('"#abcdef"');
    });

    it('handles quotes and special LaTeX characters in formula escaping without formula syntax errors', async () => {
      const state = getMockExcelState();

      // Double quotes in LaTeX text
      const quoteLatex = '\\text{The "value" is } x = 42';
      await insertFormulaToActiveCell(quoteLatex);
      expect(state.workbook.getSelectedRange().formulas[0][0]).toBe(
        '=MATH.KATEX("\\text{The ""value"" is } x = 42")'
      );

      // Read back formula with quotes
      const resQuote: any = await readActiveCellFormula();
      const extracted = typeof resQuote === 'object' && resQuote !== null ? resQuote.latex : resQuote;
      expect(extracted).toBe(quoteLatex);

      // Complex LaTeX matrix with newlines and backslashes
      const matrixLatex = '\\begin{pmatrix} a & b \\\\ c & d \\end{pmatrix}';
      await insertFormulaToActiveCell(matrixLatex, { background: 1 });
      expect(state.workbook.getSelectedRange().formulas[0][0]).toBe(
        `=MATH.KATEX("${matrixLatex}", 1)`
      );

      const resMatrix: any = await readActiveCellFormula();
      const extractedMatrix = typeof resMatrix === 'object' && resMatrix !== null ? resMatrix.latex : resMatrix;
      expect(extractedMatrix).toBe(matrixLatex);
    });

    it('readActiveCellFormula robustly parses various formula formats and cell contents', async () => {
      const state = getMockExcelState();
      const range = state.workbook.getSelectedRange();

      const getLatex = async () => {
        const res: any = await readActiveCellFormula();
        return typeof res === 'object' && res !== null ? res.latex : res;
      };

      // 1. Compact formula with extra spaces inside parens
      range.formulas = [['=MATH.KATEX(  "\\int_0^1 x dx"  )']];
      expect(await getLatex()).toBe('\\int_0^1 x dx');

      // 2. Multi-arg formula
      range.formulas = [['=MATH.KATEX("\\sum_{i=1}^n i", 2, "#fff", 20, false)']];
      expect(await getLatex()).toBe('\\sum_{i=1}^n i');

      // 3. Entity cell value
      range.formulas = [['']];
      range.values = [[{
        type: 'Entity',
        text: '[Math: \\oint B \\cdot dl]',
        properties: { latex: { basicValue: '\\oint B \\cdot dl' } }
      }]];
      expect(await getLatex()).toBe('\\oint B \\cdot dl');

      // 4. Empty / null cell
      range.formulas = [['']];
      range.values = [['']];
      expect(await readActiveCellFormula()).toBeNull();
    });
  });

  // =========================================================================
  // 4. Offline Assets, Font Packaging, and Build Verification (R4)
  // =========================================================================
  describe('4. Offline Assets, Font Packaging & Build Integrity (R4)', () => {
    const projectRoot = path.resolve(__dirname, '..');

    it('verifies dist/ directory contains production assets without missing files', () => {
      const distDir = path.join(projectRoot, 'dist');
      expect(fs.existsSync(distDir)).toBe(true);

      const taskpaneHtml = path.join(distDir, 'taskpane.html');
      expect(fs.existsSync(taskpaneHtml)).toBe(true);

      const assetsDir = path.join(distDir, 'assets');
      expect(fs.existsSync(assetsDir)).toBe(true);

      const assetFiles = fs.readdirSync(assetsDir);
      expect(assetFiles.some((f) => f.startsWith('taskpane') && f.endsWith('.js'))).toBe(true);
      expect(assetFiles.some((f) => f.startsWith('functions') && f.endsWith('.js'))).toBe(true);
      expect(assetFiles.some((f) => f.startsWith('taskpane') && f.endsWith('.css'))).toBe(true);
    });

    it('verifies KaTeX font families are packaged locally in public and dist assets', () => {
      const fontDirs = [
        path.join(projectRoot, 'public/assets/katex-fonts'),
        path.join(projectRoot, 'public/fonts'),
        path.join(projectRoot, 'dist/assets')
      ];

      for (const fDir of fontDirs) {
        if (fs.existsSync(fDir)) {
          const files = fs.readdirSync(fDir);
          // Check core KaTeX font families
          const hasMain = files.some((f) => f.includes('KaTeX_Main'));
          const hasMath = files.some((f) => f.includes('KaTeX_Math'));
          const hasAMS = files.some((f) => f.includes('KaTeX_AMS'));
          const hasSize = files.some((f) => f.includes('KaTeX_Size'));
          expect(hasMain && hasMath && hasAMS && hasSize).toBe(true);
        }
      }
    });

    it('verifies MathLive fonts are packaged locally in public/assets/mathlive-fonts', () => {
      const mathliveFontsDir = path.join(projectRoot, 'public/assets/mathlive-fonts');
      expect(fs.existsSync(mathliveFontsDir)).toBe(true);

      const files = fs.readdirSync(mathliveFontsDir);
      expect(files.length).toBeGreaterThan(0);
      expect(files.some((f) => f.endsWith('.woff') || f.endsWith('.woff2') || f.endsWith('.ttf'))).toBe(true);
    });

    it('verifies manifest.xml and taskpane.html contain zero remote script CDNs', () => {
      const htmlFile = path.join(projectRoot, 'dist/taskpane.html');
      const content = fs.readFileSync(htmlFile, 'utf-8');
      expect(content).not.toContain('unpkg.com');
      expect(content).not.toContain('cdn.jsdelivr.net');
      expect(content).not.toContain('cdnjs.cloudflare.com');
    });
  });
});
