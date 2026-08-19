import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import { LatexSyncEngine, LatexSyncState } from '../../src/taskpane/hooks/useLatexSync';

describe('Unit Test: Bi-directional Synchronization State Engine (useLatexSync Contract)', () => {
  beforeEach(() => {
    vi.useFakeTimers();
  });

  afterEach(() => {
    vi.useRealTimers();
  });

  it('initializes with default state and source "initial"', () => {
    const engine = new LatexSyncEngine('E = mc^2');
    const state = engine.getState();
    expect(state.latex).toBe('E = mc^2');
    expect(state.activeSource).toBe('initial');
    expect(state.isValid).toBe(true);
    expect(state.errorMessage).toBeUndefined();
  });

  it('immediately updates state (0ms) when raw LaTeX is entered', () => {
    const engine = new LatexSyncEngine();
    let latestState: LatexSyncState | null = null;
    engine.subscribe(s => { latestState = s; });

    engine.updateFromRawLatex('\\int_0^1 x dx');
    expect(latestState).not.toBeNull();
    expect(latestState!.latex).toBe('\\int_0^1 x dx');
    expect(latestState!.activeSource).toBe('raw_latex');
    expect(latestState!.isValid).toBe(true);
  });

  it('detects invalid LaTeX syntax immediately and sets error flags', () => {
    const engine = new LatexSyncEngine();
    let latestState: LatexSyncState | null = null;
    engine.subscribe(s => { latestState = s; });

    engine.updateFromRawLatex('\\frac{1}'); // Missing denominator
    expect(latestState!.isValid).toBe(false);
    expect(latestState!.errorMessage).toContain('KaTeX parse error');

    // Fix syntax
    engine.updateFromRawLatex('\\frac{1}{2}');
    expect(latestState!.isValid).toBe(true);
    expect(latestState!.errorMessage).toBeUndefined();
  });

  it('debounces rapid typing events to prevent excessive recalculations', () => {
    const engine = new LatexSyncEngine('', 150);
    const notifications: string[] = [];
    engine.subscribe(s => { notifications.push(s.latex); });

    // Rapid keystrokes within 100ms
    engine.updateFromRawLatex('\\s');
    vi.advanceTimersByTime(20);
    engine.updateFromRawLatex('\\sq');
    vi.advanceTimersByTime(30);
    engine.updateFromRawLatex('\\sqrt');
    vi.advanceTimersByTime(40);
    engine.updateFromRawLatex('\\sqrt{x}');

    // Advance beyond debounce window (150ms)
    vi.advanceTimersByTime(200);

    // Final state should be \sqrt{x} and valid
    expect(engine.getState().latex).toBe('\\sqrt{x}');
    expect(engine.getState().isValid).toBe(true);
  });

  it('correctly sets equation from preset or cell without debounced latency', () => {
    const engine = new LatexSyncEngine('', 150);
    engine.setEquation('\\mathcal{L}_{\\text{BS}}', 'preset');

    const state = engine.getState();
    expect(state.latex).toBe('\\mathcal{L}_{\\text{BS}}');
    expect(state.activeSource).toBe('preset');
    expect(state.isValid).toBe(true);
  });

  it('prevents cyclic ping-pong feedback loops when normalized LaTeX is identical', () => {
    const engine = new LatexSyncEngine('x + y', 150);
    let notifyCount = 0;
    engine.subscribe(() => { notifyCount++; });

    engine.updateFromMathField('x + y');
    expect(notifyCount).toBe(0); // Ignored because identical content
  });
});
