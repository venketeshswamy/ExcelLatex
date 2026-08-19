import React, { useRef } from 'react';
import {
  makeStyles,
  tokens,
  shorthands,
  Textarea,
  Caption1,
  Badge,
  Button,
  Tooltip
} from '@fluentui/react-components';
import {
  DismissCircle16Filled,
  CheckmarkCircle16Filled,
  Delete20Regular
} from '@fluentui/react-icons';
import { LatexSource } from '../hooks/useLatexSync';

const useStyles = makeStyles({
  container: {
    display: 'flex',
    flexDirection: 'column',
    ...shorthands.gap('8px'),
    width: '100%'
  },
  toolbarRow: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between'
  },
  labelRow: {
    display: 'flex',
    alignItems: 'center',
    ...shorthands.gap('6px')
  },
  label: {
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground1
  },
  symbolsBar: {
    display: 'flex',
    flexWrap: 'wrap',
    ...shorthands.gap('4px'),
    ...shorthands.padding('4px', '0')
  },
  symbolBtn: {
    minWidth: '32px',
    height: '28px',
    fontFamily: 'KaTeX_Main, "Times New Roman", serif',
    fontSize: '13px',
    ...shorthands.padding('2px', '6px')
  },
  textarea: {
    fontFamily: 'Consolas, Monaco, "Courier New", monospace',
    fontSize: '13px',
    lineHeight: '1.4',
    width: '100%'
  },
  metaRow: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    fontSize: '11px',
    color: tokens.colorNeutralForeground3
  },
  errorBox: {
    display: 'flex',
    alignItems: 'flex-start',
    ...shorthands.gap('6px'),
    ...shorthands.padding('6px', '8px'),
    ...shorthands.borderRadius(tokens.borderRadiusSmall),
    backgroundColor: tokens.colorPaletteRedBackground1,
    ...shorthands.border('1px', 'solid', tokens.colorPaletteRedBorder1),
    color: tokens.colorPaletteRedForeground1,
    fontSize: '12px'
  }
});

const QUICK_SYMBOLS = [
  { label: 'a/b', insert: '\\frac{a}{b}', tooltip: 'Fraction' },
  { label: '√x', insert: '\\sqrt{x}', tooltip: 'Square Root' },
  { label: 'xⁿ', insert: 'x^{n}', tooltip: 'Superscript / Power' },
  { label: 'xₙ', insert: 'x_{n}', tooltip: 'Subscript' },
  { label: '∫', insert: '\\int_{a}^{b} f(x) \\, dx', tooltip: 'Definite Integral' },
  { label: '∑', insert: '\\sum_{i=1}^{n} x_i', tooltip: 'Summation' },
  { label: '∏', insert: '\\prod_{i=1}^{n} x_i', tooltip: 'Product' },
  { label: 'lim', insert: '\\lim_{x \\to 0}', tooltip: 'Limit' },
  { label: 'α', insert: '\\alpha', tooltip: 'Alpha' },
  { label: 'β', insert: '\\beta', tooltip: 'Beta' },
  { label: 'θ', insert: '\\theta', tooltip: 'Theta' },
  { label: 'π', insert: '\\pi', tooltip: 'Pi' },
  { label: '∞', insert: '\\infty', tooltip: 'Infinity' },
  { label: '∂', insert: '\\partial', tooltip: 'Partial Derivative' },
  { label: '±', insert: '\\pm', tooltip: 'Plus-Minus' },
  { label: '×', insert: '\\times', tooltip: 'Multiplication' },
  { label: '≤', insert: '\\le', tooltip: 'Less than or Equal' },
  { label: '≥', insert: '\\ge', tooltip: 'Greater than or Equal' },
  { label: '≠', insert: '\\neq', tooltip: 'Not Equal' },
  { label: '[M]', insert: '\\begin{pmatrix} a & b \\\\ c & d \\end{pmatrix}', tooltip: '2x2 Matrix' }
];

export interface RawLatexEditorProps {
  latex: string;
  activeSource: LatexSource;
  isValid: boolean;
  errorMessage?: string;
  onLatexChange: (newLatex: string) => void;
  onClear?: () => void;
}

export const RawLatexEditor: React.FC<RawLatexEditorProps> = ({
  latex,
  isValid,
  errorMessage,
  onLatexChange,
  onClear
}) => {
  const styles = useStyles();
  const textareaRef = useRef<HTMLTextAreaElement>(null);

  const handleChange = (e: React.ChangeEvent<HTMLTextAreaElement>) => {
    onLatexChange(e.target.value);
  };

  const handleInsertSymbol = (snippet: string) => {
    const el = textareaRef.current;
    if (!el) {
      onLatexChange((latex ? latex + ' ' : '') + snippet);
      return;
    }

    const start = el.selectionStart ?? latex.length;
    const end = el.selectionEnd ?? latex.length;
    const before = latex.substring(0, start);
    const after = latex.substring(end);
    const updated = before + snippet + after;

    onLatexChange(updated);

    setTimeout(() => {
      el.focus();
      const newPos = start + snippet.length;
      el.setSelectionRange(newPos, newPos);
    }, 0);
  };

  return (
    <div className={styles.container}>
      <div className={styles.toolbarRow}>
        <div className={styles.labelRow}>
          <Caption1 className={styles.label}>Raw LaTeX Code</Caption1>
          <Badge
            appearance="tint"
            color={isValid ? 'success' : 'danger'}
            icon={isValid ? <CheckmarkCircle16Filled /> : <DismissCircle16Filled />}
            size="small"
          >
            {isValid ? 'Valid KaTeX' : 'Invalid Syntax'}
          </Badge>
        </div>

        {onClear && (
          <Button
            size="small"
            appearance="subtle"
            icon={<Delete20Regular />}
            onClick={onClear}
            aria-label="Clear LaTeX"
            title="Clear LaTeX"
          />
        )}
      </div>

      <div className={styles.symbolsBar} role="toolbar" aria-label="Quick LaTeX Math Symbols">
        {QUICK_SYMBOLS.map((sym) => (
          <Tooltip key={sym.label} content={sym.tooltip} relationship="description">
            <Button
              size="small"
              appearance="subtle"
              className={styles.symbolBtn}
              onClick={() => handleInsertSymbol(sym.insert)}
            >
              {sym.label}
            </Button>
          </Tooltip>
        ))}
      </div>

      <Textarea
        ref={textareaRef}
        className={styles.textarea}
        value={latex}
        onChange={handleChange}
        placeholder="Enter LaTeX formula here (e.g. \int_0^1 x^2 \, dx)"
        rows={4}
        resize="vertical"
        aria-label="Raw LaTeX Code Editor"
      />

      <div className={styles.metaRow}>
        <span>{latex.length} characters</span>
        <span>{latex.trim() ? latex.trim().split(/\s+/).length : 0} tokens</span>
      </div>

      {!isValid && errorMessage && (
        <div className={styles.errorBox} role="alert">
          <DismissCircle16Filled style={{ flexShrink: 0, marginTop: '2px' }} />
          <div>{errorMessage}</div>
        </div>
      )}
    </div>
  );
};
