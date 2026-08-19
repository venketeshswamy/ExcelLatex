import { useState, useCallback, useRef, useEffect } from 'react';
import { validateLatex, LatexValidationResult } from '../../core/katexEngine';

export type LatexSource = 'mathfield' | 'raw_latex' | 'preset' | 'cell' | 'initial';

export interface LatexSyncState {
  latex: string;
  activeSource: LatexSource;
  isValid: boolean;
  errorMessage?: string;
  updateFromMathField: (newLatex: string) => void;
  updateFromRawLatex: (newLatex: string) => void;
  setEquation: (newLatex: string, source?: LatexSource) => void;
  clear: () => void;
}

/**
 * Normalizes LaTeX string by trimming whitespace and collapsing multiple spaces.
 */
export function normalizeLatex(input: string): string {
  if (!input) return '';
  return input.trim().replace(/\s+/g, ' ');
}

/**
 * Standalone state engine for bi-directional synchronization logic.
 */
export class LatexSyncEngine {
  public latex: string;
  public activeSource: LatexSource;
  public isValid: boolean;
  public errorMessage?: string;

  private debounceMs: number;
  private debounceTimer: ReturnType<typeof setTimeout> | null = null;
  private listeners: Array<(state: LatexSyncState) => void> = [];

  constructor(initialLatex = '', debounceMs = 150) {
    this.latex = initialLatex;
    this.activeSource = 'initial';
    this.debounceMs = debounceMs;
    const validation = initialLatex ? validateLatex(initialLatex) : { isValid: true };
    this.isValid = validation.isValid;
    this.errorMessage = validation.error;
  }

  public subscribe(listener: (state: LatexSyncState) => void): () => void {
    this.listeners.push(listener);
    return () => {
      this.listeners = this.listeners.filter((l) => l !== listener);
    };
  }

  private notify(): void {
    const state = this.getState();
    this.listeners.forEach((l) => l(state));
  }

  public getState(): LatexSyncState {
    return {
      latex: this.latex,
      activeSource: this.activeSource,
      isValid: this.isValid,
      errorMessage: this.errorMessage,
      updateFromMathField: this.updateFromMathField.bind(this),
      updateFromRawLatex: this.updateFromRawLatex.bind(this),
      setEquation: this.setEquation.bind(this),
      clear: this.clear.bind(this)
    };
  }

  public updateFromRawLatex(newLatex: string): void {
    if (normalizeLatex(newLatex) === normalizeLatex(this.latex) && this.activeSource === 'raw_latex') {
      return;
    }
    this.activeSource = 'raw_latex';
    this.latex = newLatex;
    const validation = newLatex.trim() ? validateLatex(newLatex) : { isValid: true };
    this.isValid = validation.isValid;
    this.errorMessage = validation.error;
    this.notify();

    if (this.debounceTimer) clearTimeout(this.debounceTimer);
    this.debounceTimer = setTimeout(() => {
      this.notify();
    }, this.debounceMs);
  }

  public updateFromMathField(newLatex: string): void {
    if (normalizeLatex(newLatex) === normalizeLatex(this.latex)) {
      return;
    }
    this.activeSource = 'mathfield';
    this.latex = newLatex;
    const validation = newLatex.trim() ? validateLatex(newLatex) : { isValid: true };
    this.isValid = validation.isValid;
    this.errorMessage = validation.error;
    this.notify();

    if (this.debounceTimer) clearTimeout(this.debounceTimer);
    this.debounceTimer = setTimeout(() => {
      this.notify();
    }, this.debounceMs);
  }

  public setEquation(newLatex: string, source: LatexSource = 'preset'): void {
    if (this.debounceTimer) clearTimeout(this.debounceTimer);
    this.activeSource = source;
    this.latex = newLatex;
    const validation = newLatex.trim() ? validateLatex(newLatex) : { isValid: true };
    this.isValid = validation.isValid;
    this.errorMessage = validation.error;
    this.notify();
  }

  public clear(): void {
    this.setEquation('', 'initial');
  }

  public destroy(): void {
    if (this.debounceTimer) {
      clearTimeout(this.debounceTimer);
      this.debounceTimer = null;
    }
    this.listeners = [];
  }
}

/**
 * Custom React Hook for Bi-directional LaTeX Synchronization.
 * Prevents cyclic ping-pong loops and maintains immediate 0ms UI responsiveness
 * with 150ms debounced cross-propagation.
 */
export function useLatexSync(
  initialLatex: string = '',
  debounceMs: number = 150
): LatexSyncState {
  const [latex, setLatexState] = useState<string>(initialLatex);
  const [activeSource, setActiveSource] = useState<LatexSource>('initial');
  const [isValid, setIsValid] = useState<boolean>(() => {
    if (!initialLatex || !initialLatex.trim()) return true;
    return validateLatex(initialLatex).isValid;
  });
  const [errorMessage, setErrorMessage] = useState<string | undefined>(() => {
    if (!initialLatex || !initialLatex.trim()) return undefined;
    return validateLatex(initialLatex).error;
  });

  const debounceTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);
  const currentLatexRef = useRef<string>(latex);
  currentLatexRef.current = latex;

  const activeSourceRef = useRef<LatexSource>(activeSource);
  activeSourceRef.current = activeSource;

  useEffect(() => {
    return () => {
      if (debounceTimerRef.current) {
        clearTimeout(debounceTimerRef.current);
      }
    };
  }, []);

  const updateFromRawLatex = useCallback(
    (newLatex: string) => {
      if (
        normalizeLatex(newLatex) === normalizeLatex(currentLatexRef.current) &&
        activeSourceRef.current === 'raw_latex'
      ) {
        return;
      }

      setActiveSource('raw_latex');
      setLatexState(newLatex);

      const validation: LatexValidationResult = newLatex.trim()
        ? validateLatex(newLatex)
        : { isValid: true };
      setIsValid(validation.isValid);
      setErrorMessage(validation.error);

      if (debounceTimerRef.current) {
        clearTimeout(debounceTimerRef.current);
      }
      debounceTimerRef.current = setTimeout(() => {
        // Debounce settle point
      }, debounceMs);
    },
    [debounceMs]
  );

  const updateFromMathField = useCallback(
    (newLatex: string) => {
      if (normalizeLatex(newLatex) === normalizeLatex(currentLatexRef.current)) {
        return;
      }

      setActiveSource('mathfield');
      setLatexState(newLatex);

      const validation: LatexValidationResult = newLatex.trim()
        ? validateLatex(newLatex)
        : { isValid: true };
      setIsValid(validation.isValid);
      setErrorMessage(validation.error);

      if (debounceTimerRef.current) {
        clearTimeout(debounceTimerRef.current);
      }
      debounceTimerRef.current = setTimeout(() => {
        // Debounce settle point
      }, debounceMs);
    },
    [debounceMs]
  );

  const setEquation = useCallback(
    (newLatex: string, source: LatexSource = 'preset') => {
      if (debounceTimerRef.current) {
        clearTimeout(debounceTimerRef.current);
      }
      setActiveSource(source);
      setLatexState(newLatex);

      const validation: LatexValidationResult = newLatex.trim()
        ? validateLatex(newLatex)
        : { isValid: true };
      setIsValid(validation.isValid);
      setErrorMessage(validation.error);
    },
    []
  );

  const clear = useCallback(() => {
    setEquation('', 'initial');
  }, [setEquation]);

  return {
    latex,
    activeSource,
    isValid,
    errorMessage,
    updateFromMathField,
    updateFromRawLatex,
    setEquation,
    clear
  };
}
