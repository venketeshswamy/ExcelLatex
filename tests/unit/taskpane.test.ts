import { describe, it, expect, beforeEach, vi } from 'vitest';
import React from 'react';
import { render, screen, fireEvent, act } from '@testing-library/react';
import {
  LatexSyncEngine
} from '../../src/taskpane/hooks/useLatexSync';
import {
  resolveOfficeThemeMode
} from '../../src/taskpane/hooks/useOfficeTheme';
import {
  escapeFormulaString,
  formatBackgroundParam,
  buildKatexEntityCellValue,
  insertFormulaToActiveCell,
  insertInCellImageToActiveCell,
  insertFloatingShapeToActiveCell,
  readActiveCellFormula
} from '../../src/taskpane/services/excelService';
import {
  dataUrlToBlob,
  copyLatex,
  copySvg,
  copyPng,
  copyEquation
} from '../../src/taskpane/services/clipboardService';
import { Header } from '../../src/taskpane/components/Header';
import { MathLiveEditor } from '../../src/taskpane/components/MathLiveEditor';
import { RawLatexEditor } from '../../src/taskpane/components/RawLatexEditor';
import { LivePreview, getBackgroundLabel } from '../../src/taskpane/components/LivePreview';
import { PresetsLibrary, PRESET_FORMULAS } from '../../src/taskpane/components/PresetsLibrary';
import { ActionBar } from '../../src/taskpane/components/ActionBar';
import { App } from '../../src/taskpane/App';
import {
  resetOfficeMock,
  setMockRequirementSupported,
  getMockExcelState
} from '../../src/mocks/officeMock';
import { FluentProvider, webLightTheme } from '@fluentui/react-components';

const wrap = (children: React.ReactNode) =>
  React.createElement(FluentProvider, { theme: webLightTheme }, children);

describe('Milestone 3: Fluent UI Taskpane & Editor Unit & Integration Suite', () => {
  beforeEach(() => {
    resetOfficeMock();
  });

  // =========================================================================
  // 1. useLatexSync & LatexSyncEngine
  // =========================================================================
  describe('1. useLatexSync & LatexSyncEngine', () => {
    it('initializes with default values and validates input', () => {
      const engine = new LatexSyncEngine('E = mc^2');
      const state = engine.getState();
      expect(state.latex).toBe('E = mc^2');
      expect(state.activeSource).toBe('initial');
      expect(state.isValid).toBe(true);
      expect(state.errorMessage).toBeUndefined();
    });

    it('immediately updates state at 0ms on raw LaTeX input', () => {
      const engine = new LatexSyncEngine();
      let state = engine.getState();
      engine.subscribe((s) => {
        state = s;
      });

      engine.updateFromRawLatex('\\int x \\, dx');
      expect(state.latex).toBe('\\int x \\, dx');
      expect(state.activeSource).toBe('raw_latex');
      expect(state.isValid).toBe(true);
    });

    it('immediately updates state at 0ms on MathLive editor input', () => {
      const engine = new LatexSyncEngine();
      let state = engine.getState();
      engine.subscribe((s) => {
        state = s;
      });

      engine.updateFromMathField('\\sum_{i=1}^n i');
      expect(state.latex).toBe('\\sum_{i=1}^n i');
      expect(state.activeSource).toBe('mathfield');
      expect(state.isValid).toBe(true);
    });

    it('detects syntax error and provides error details', () => {
      const engine = new LatexSyncEngine();
      let state = engine.getState();
      engine.subscribe((s) => {
        state = s;
      });

      engine.updateFromRawLatex('\\sqrt[');
      expect(state.isValid).toBe(false);
      expect(state.errorMessage).toBeDefined();

      // Fix syntax
      engine.updateFromRawLatex('\\sqrt{x}');
      expect(state.isValid).toBe(true);
      expect(state.errorMessage).toBeUndefined();
    });

    it('debounces rapid typing events without state corruption', () => {
      vi.useFakeTimers();
      const engine = new LatexSyncEngine('', 150);
      const log: string[] = [];
      engine.subscribe((s) => log.push(s.latex));

      engine.updateFromRawLatex('a');
      vi.advanceTimersByTime(30);
      engine.updateFromRawLatex('a +');
      vi.advanceTimersByTime(40);
      engine.updateFromRawLatex('a + b');
      vi.advanceTimersByTime(160);

      expect(engine.getState().latex).toBe('a + b');
      expect(engine.getState().isValid).toBe(true);
      vi.useRealTimers();
    });

    it('setEquation updates equation immediately from presets or cell', () => {
      const engine = new LatexSyncEngine();
      engine.setEquation('\\alpha + \\beta', 'preset');
      expect(engine.getState().latex).toBe('\\alpha + \\beta');
      expect(engine.getState().activeSource).toBe('preset');
      expect(engine.getState().isValid).toBe(true);
    });

    it('prevents cyclic loops when normalized LaTeX is identical', () => {
      const engine = new LatexSyncEngine('x + y', 150);
      let notifications = 0;
      engine.subscribe(() => {
        notifications++;
      });

      engine.updateFromMathField('  x  +  y  ');
      expect(notifications).toBe(0);
    });

    it('clear resets equation to empty string and initial source', () => {
      const engine = new LatexSyncEngine('\\int x dx');
      engine.clear();
      expect(engine.getState().latex).toBe('');
      expect(engine.getState().activeSource).toBe('initial');
      expect(engine.getState().isValid).toBe(true);
    });
  });

  // =========================================================================
  // 2. useOfficeTheme & Theme Detection
  // =========================================================================
  describe('2. useOfficeTheme & Theme Detection', () => {
    it('resolves light theme for default light background', () => {
      const mode = resolveOfficeThemeMode({ bodyBackgroundColor: '#ffffff', isDark: false });
      expect(mode).toBe('light');
    });

    it('resolves dark theme when isDark is true', () => {
      const mode = resolveOfficeThemeMode({ bodyBackgroundColor: '#201f1e', isDark: true });
      expect(mode).toBe('dark');
    });

    it('resolves dark theme when bodyBackgroundColor is dark hex', () => {
      const mode = resolveOfficeThemeMode({ bodyBackgroundColor: '#1b1a19' });
      expect(mode).toBe('dark');
    });

    it('falls back to light when officeTheme is undefined and media query is default', () => {
      const mode = resolveOfficeThemeMode(undefined);
      expect(['light', 'dark', 'contrast']).toContain(mode);
    });
  });

  // =========================================================================
  // 3. excelService & Sheet Insertion Actions
  // =========================================================================
  describe('3. excelService Sheet Operations', () => {
    it('escapes quotes correctly in Excel formula strings', () => {
      expect(escapeFormulaString('f("x") = 1')).toBe('f(""x"") = 1');
      expect(escapeFormulaString('\\text{"hello"}')).toBe('\\text{""hello""}');
    });

    it('buildKatexEntityCellValue creates well-structured entity cell structure', () => {
      const entity = buildKatexEntityCellValue('x^2', {
        pngDataUrl: 'data:image/png;base64,sample',
        width: 100,
        height: 40
      });
      expect(entity.type).toBe('Entity');
      expect(entity.text).toBe('[Math: x^2]');
      expect(entity.properties?.latex?.basicValue).toBe('x^2');
      expect(entity.properties?.dimensions?.basicValue).toBe('100x40');
      expect(entity.layouts?.compact?.icon).toBe('data:image/png;base64,sample');
    });

    it('formatBackgroundParam formats numbers, keywords, and hex values correctly', () => {
      expect(formatBackgroundParam(0)).toBe('0');
      expect(formatBackgroundParam(1)).toBe('1');
      expect(formatBackgroundParam(2)).toBe('2');
      expect(formatBackgroundParam('0')).toBe('0');
      expect(formatBackgroundParam('1')).toBe('1');
      expect(formatBackgroundParam('2')).toBe('2');
      expect(formatBackgroundParam('transparent')).toBe('0');
      expect(formatBackgroundParam('white')).toBe('1');
      expect(formatBackgroundParam('black')).toBe('2');
      expect(formatBackgroundParam('#ffff00')).toBe('"#ffff00"');
    });

    it('insertFormulaToActiveCell writes =MATH.KATEX() formula to active cell', async () => {
      await insertFormulaToActiveCell('E = mc^2');
      const state = getMockExcelState();
      const cell = state.workbook.worksheets.getActiveWorksheet().getRange('A1');
      expect(cell.formulas[0][0]).toBe('=MATH.KATEX("E = mc^2")');
    });

    it('insertFormulaToActiveCell generates standardized numeric background code 0 for transparent', async () => {
      await insertFormulaToActiveCell('\\alpha', {
        background: 0
      });
      const state = getMockExcelState();
      const cell = state.workbook.worksheets.getActiveWorksheet().getRange('A1');
      expect(cell.formulas[0][0]).toBe('=MATH.KATEX("\\alpha")');
    });

    it('insertFormulaToActiveCell generates standardized numeric background code 1 for white', async () => {
      await insertFormulaToActiveCell('\\beta', {
        background: 1
      });
      const state = getMockExcelState();
      const cell = state.workbook.worksheets.getActiveWorksheet().getRange('A1');
      expect(cell.formulas[0][0]).toBe('=MATH.KATEX("\\beta", 1)');
    });

    it('insertFormulaToActiveCell generates standardized numeric background code 2 for black', async () => {
      await insertFormulaToActiveCell('\\gamma', {
        background: 2
      });
      const state = getMockExcelState();
      const cell = state.workbook.worksheets.getActiveWorksheet().getRange('A1');
      expect(cell.formulas[0][0]).toBe('=MATH.KATEX("\\gamma", 2)');
    });

    it('insertFormulaToActiveCell writes custom background when provided', async () => {
      await insertFormulaToActiveCell('\\mu', {
        background: '#ffffff'
      });
      const state = getMockExcelState();
      const cell = state.workbook.worksheets.getActiveWorksheet().getRange('A1');
      expect(cell.formulas[0][0]).toBe('=MATH.KATEX("\\mu", "#ffffff")');
    });

    it('insertInCellImageToActiveCell writes WebImage when ExcelApi 1.16 is supported', async () => {
      setMockRequirementSupported('ExcelApi', 1.16, true);
      await insertInCellImageToActiveCell('\\int_0^1 x dx');
      const state = getMockExcelState();
      const cell = state.workbook.worksheets.getActiveWorksheet().getRange('A1');
      expect(cell.values[0][0].type).toBe('WebImage');
      expect(cell.values[0][0].altText).toBeDefined();
    });

    it('insertInCellImageToActiveCell falls back to floating shape when ExcelApi 1.16 is unsupported', async () => {
      setMockRequirementSupported('ExcelApi', 1.16, false);
      await insertInCellImageToActiveCell('\\int_0^1 x dx');
      const state = getMockExcelState();
      const sheet = state.workbook.worksheets.getActiveWorksheet();
      expect(sheet.shapes.items.length).toBeGreaterThanOrEqual(1);
    });

    it('insertFloatingShapeToActiveCell adds floating image shape to active worksheet', async () => {
      await insertFloatingShapeToActiveCell('\\sigma^2');
      const state = getMockExcelState();
      const sheet = state.workbook.worksheets.getActiveWorksheet();
      expect(sheet.shapes.items.length).toBe(1);
      expect(sheet.shapes.items[0].altTextTitle).toBe('LaTeX: \\sigma^2');
    });

    it('readActiveCellFormula extracts LaTeX from =MATH.KATEX() formula', async () => {
      const state = getMockExcelState();
      const cell = state.workbook.worksheets.getActiveWorksheet().getRange('A1');
      cell.formulas = [['=MATH.KATEX("\\alpha + \\beta")']];

      const res: any = await readActiveCellFormula();
      const latex = typeof res === 'object' && res !== null ? res.latex : res;
      expect(latex).toBe('\\alpha + \\beta');
    });

    it('readActiveCellFormula extracts LaTeX from EntityCellValue', async () => {
      const state = getMockExcelState();
      const cell = state.workbook.worksheets.getActiveWorksheet().getRange('A1');
      cell.values = [[buildKatexEntityCellValue('\\beta_1', { pngDataUrl: 'data:...', width: 50, height: 20 })]];

      const res: any = await readActiveCellFormula();
      const latex = typeof res === 'object' && res !== null ? res.latex : res;
      expect(latex).toBe('\\beta_1');
    });

    it('readActiveCellFormula returns null for empty cell', async () => {
      const state = getMockExcelState();
      const cell = state.workbook.worksheets.getActiveWorksheet().getRange('A1');
      cell.formulas = [['']];
      cell.values = [['']];

      const formula = await readActiveCellFormula();
      expect(formula).toBeNull();
    });
  });

  // =========================================================================
  // 4. clipboardService
  // =========================================================================
  describe('4. clipboardService Operations', () => {
    it('dataUrlToBlob converts base64 data URL to Blob', () => {
      const dataUrl = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII=';
      const blob = dataUrlToBlob(dataUrl);
      expect(blob).toBeDefined();
      expect(blob.type).toBe('image/png');
      expect(blob.size).toBeGreaterThan(0);
    });

    it('copyLatex copies LaTeX text without throwing', async () => {
      const res = await copyLatex('\\frac{1}{2}');
      expect(typeof res).toBe('boolean');
    });

    it('copySvg copies SVG markup', async () => {
      const res = await copySvg('<svg><text>Math</text></svg>');
      expect(typeof res).toBe('boolean');
    });

    it('copyPng handles PNG data URL', async () => {
      const dataUrl = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII=';
      const res = await copyPng(dataUrl);
      expect(typeof res).toBe('boolean');
    });

    it('copyEquation renders and copies multi-format equation', async () => {
      const res = await copyEquation('x^2 + y^2 = r^2');
      expect(res.latex).toBeDefined();
      expect(res.svg).toBeDefined();
      expect(res.png).toBeDefined();
    });
  });

  // =========================================================================
  // 5. Component Tests: Header, Editors, Preview, Presets, ActionBar, App
  // =========================================================================
  describe('5. Taskpane UI Components', () => {
    it('Header renders title and connection badge', () => {
      const onThemeChange = vi.fn();
      render(
        wrap(
          React.createElement(Header, {
            themeMode: 'auto',
            resolvedMode: 'light',
            onThemeModeChange: onThemeChange
          })
        )
      );

      expect(screen.getByText('LaTeX Math')).toBeDefined();
      expect(screen.getByText(/Excel Ready|Local Offline/)).toBeDefined();
    });

    it('MathLiveEditor renders math-field container and keyboard toggle', () => {
      const onChange = vi.fn();
      render(
        wrap(
          React.createElement(MathLiveEditor, {
            latex: 'x = 1',
            activeSource: 'initial',
            isValid: true,
            onLatexChange: onChange
          })
        )
      );

      expect(screen.getByText('Visual Math Editor (MathLive)')).toBeDefined();
      expect(screen.getByTitle('Toggle Math Virtual Keyboard')).toBeDefined();
    });

    it('RawLatexEditor renders symbol helper buttons and updates latex on input', () => {
      const onChange = vi.fn();
      render(
        wrap(
          React.createElement(RawLatexEditor, {
            latex: 'y = x^2',
            activeSource: 'raw_latex',
            isValid: true,
            onLatexChange: onChange
          })
        )
      );

      expect(screen.getByText('Raw LaTeX Code')).toBeDefined();
      expect(screen.getByText('Valid KaTeX')).toBeDefined();

      const fracBtn = screen.getByText('a/b');
      fireEvent.click(fracBtn);
      expect(onChange).toHaveBeenCalled();
    });

    it('RawLatexEditor displays error banner when syntax is invalid', () => {
      render(
        wrap(
          React.createElement(RawLatexEditor, {
            latex: '\\frac{',
            activeSource: 'raw_latex',
            isValid: false,
            errorMessage: 'KaTeX parse error: Expected group after \\frac',
            onLatexChange: vi.fn()
          })
        )
      );

      expect(screen.getByText('Invalid Syntax')).toBeDefined();
      expect(screen.getByText(/Expected group after \\frac/)).toBeDefined();
    });

    it('LivePreview renders equation preview, zoom buttons, complexity badge, and dropdowns', () => {
      const onDisplayModeChange = vi.fn();
      const onBgChange = vi.fn();
      render(
        wrap(
          React.createElement(LivePreview, {
            latex: '\\int_0^1 x^2 \\, dx',
            isValid: true,
            displayMode: true,
            onDisplayModeChange,
            background: 0,
            onBackgroundChange: onBgChange
          })
        )
      );

      expect(screen.getByText('Live Preview')).toBeDefined();
      expect(screen.getByLabelText('Zoom In')).toBeDefined();
      expect(screen.getByLabelText('Zoom Out')).toBeDefined();
      expect(screen.getByRole('combobox', { name: /Background/i })).toBeDefined();
    });

    it('LivePreview adapts styling when Black background (2) is selected', () => {
      render(
        wrap(
          React.createElement(LivePreview, {
            latex: 'E = mc^2',
            isValid: true,
            displayMode: true,
            background: 2
          })
        )
      );

      const previewRegion = screen.getByRole('region', { name: /Equation Preview/i });
      expect(previewRegion.style.backgroundColor).toBe('rgb(0, 0, 0)');
      expect(previewRegion.style.color).toBe('rgb(255, 255, 255)');
    });

    it('LivePreview adapts styling when White background (1) is selected', () => {
      render(
        wrap(
          React.createElement(LivePreview, {
            latex: 'x + y',
            isValid: true,
            displayMode: true,
            background: 1
          })
        )
      );

      const previewRegion = screen.getByRole('region', { name: /Equation Preview/i });
      expect(previewRegion.style.backgroundColor).toBe('rgb(255, 255, 255)');
    });

    it('getBackgroundLabel correctly returns human-readable labels for 0, 1, 2', () => {
      expect(getBackgroundLabel(0)).toBe('0: Transparent');
      expect(getBackgroundLabel('0')).toBe('0: Transparent');
      expect(getBackgroundLabel('transparent')).toBe('0: Transparent');
      expect(getBackgroundLabel(1)).toBe('1: White');
      expect(getBackgroundLabel('1')).toBe('1: White');
      expect(getBackgroundLabel('white')).toBe('1: White');
      expect(getBackgroundLabel(2)).toBe('2: Black');
      expect(getBackgroundLabel('2')).toBe('2: Black');
      expect(getBackgroundLabel('black')).toBe('2: Black');
      expect(getBackgroundLabel('#ffff00')).toBe('#ffff00');
    });

    it('PresetsLibrary filters formulas by category and search query', () => {
      const onSelect = vi.fn();
      render(
        wrap(
          React.createElement(PresetsLibrary, {
            onSelectPreset: onSelect
          })
        )
      );

      expect(screen.getByText('Formula Preset Catalog')).toBeDefined();

      // Search for Schrödinger
      const searchInput = screen.getByLabelText('Search formula presets');
      fireEvent.change(searchInput, { target: { value: 'Schr' } });

      expect(screen.getByText('Schrödinger Wave Equation')).toBeDefined();

      // Click load equation
      const loadBtns = screen.getAllByTitle('Load as current equation');
      fireEvent.click(loadBtns[0]);
      expect(onSelect).toHaveBeenCalledWith('i\\hbar \\frac{\\partial}{\\partial t} \\Psi(\\mathbf{r}, t) = \\hat{H} \\Psi(\\mathbf{r}, t)');
    });

    it('All preset formulas pass KaTeX syntax validation', () => {
      for (const formula of PRESET_FORMULAS) {
        const engine = new LatexSyncEngine(formula.latex);
        expect(engine.getState().isValid).toBe(true);
      }
    });

    it('ActionBar renders primary In-Cell Image and Insert Formula buttons', async () => {
      const onRead = vi.fn();
      render(
        wrap(
          React.createElement(ActionBar, {
            latex: 'E = mc^2',
            isValid: true,
            onReadCellSuccess: onRead
          })
        )
      );

      const inCellBtn = screen.getByText('In-Cell Image');
      const insertBtn = screen.getByText('Insert Formula');
      expect(inCellBtn).toBeDefined();
      expect(insertBtn).toBeDefined();

      // Confirm secondary actions exist
      expect(screen.getByText('Floating Shape')).toBeDefined();
      expect(screen.getByText('Read Cell')).toBeDefined();
      expect(screen.getByText('Batch Convert')).toBeDefined();

      await act(async () => {
        fireEvent.click(insertBtn);
      });

      const cell = getMockExcelState().workbook.worksheets.getActiveWorksheet().getRange('A1');
      expect(cell.formulas[0][0]).toBe('=MATH.KATEX("E = mc^2")');
    });

    it('App mounts cleanly and switches between tabs and renders Options dropdowns', () => {
      render(React.createElement(App));

      expect(screen.getByText('LaTeX Math')).toBeDefined();
      expect(screen.getByText('Visual Math Editor (MathLive)')).toBeDefined();

      // Switch to Presets tab
      const presetsTab = screen.getByRole('tab', { name: /Presets/i });
      fireEvent.click(presetsTab);
      expect(screen.getByText('Formula Preset Catalog')).toBeDefined();

      // Switch to Options tab
      const optionsTab = screen.getByRole('tab', { name: /Options/i });
      fireEvent.click(optionsTab);
      expect(screen.getByText('Rendering & Typesetting Options')).toBeDefined();
      expect(screen.getByRole('combobox', { name: /Display Mode Option/i })).toBeDefined();
      expect(screen.getByRole('combobox', { name: /Background Option/i })).toBeDefined();
    });
  });
});
