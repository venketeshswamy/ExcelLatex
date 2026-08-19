import React, { useEffect, useRef, useState } from 'react';
import {
  makeStyles,
  tokens,
  shorthands,
  Button,
  Caption1,
  Badge
} from '@fluentui/react-components';
import {
  Keyboard20Regular,
  Delete20Regular
} from '@fluentui/react-icons';
import { LatexSource, normalizeLatex } from '../hooks/useLatexSync';

// Configure MathLive global settings for 100% offline usage (zero CDN)
if (typeof window !== 'undefined') {
  try {
    // Import mathlive dynamically or ensure MathfieldElement configuration is set
    const win = window as any;
    if (win.MathfieldElement) {
      win.MathfieldElement.fontsDirectory = './assets/mathlive-fonts/';
      win.MathfieldElement.soundsDirectory = null;
    }
  } catch (e) {
    // Silently ignore if MathfieldElement is not yet loaded
  }
}

// Ensure custom element is known to JSX
declare global {
  namespace JSX {
    interface IntrinsicElements {
      'math-field': React.DetailedHTMLProps<React.HTMLAttributes<HTMLElement>, HTMLElement> & {
        ref?: any;
        value?: string;
        'virtual-keyboard-mode'?: string;
      };
    }
  }
}

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
  editorBox: {
    minHeight: '80px',
    backgroundColor: tokens.colorNeutralBackground1,
    ...shorthands.border('1px', 'solid', tokens.colorNeutralStroke1),
    ...shorthands.borderRadius(tokens.borderRadiusMedium),
    ...shorthands.padding('10px', '12px'),
    display: 'flex',
    alignItems: 'center',
    fontSize: '18px',
    boxShadow: tokens.shadow2,
    ':focus-within': {
      ...shorthands.border('1.5px', 'solid', tokens.colorCompoundBrandStroke)
    }
  },
  mathfield: {
    width: '100%',
    minHeight: '60px',
    outline: 'none',
    border: 'none',
    backgroundColor: 'transparent',
    fontSize: '20px',
    color: tokens.colorNeutralForeground1
  }
});

export interface MathLiveEditorProps {
  latex: string;
  activeSource: LatexSource;
  isValid: boolean;
  onLatexChange: (newLatex: string) => void;
  onClear?: () => void;
}

export const MathLiveEditor: React.FC<MathLiveEditorProps> = ({
  latex,
  activeSource,
  isValid,
  onLatexChange,
  onClear
}) => {
  const styles = useStyles();
  const mathfieldRef = useRef<any>(null);
  const [keyboardVisible, setKeyboardVisible] = useState(false);

  // Initialize MathLive element and configure properties
  useEffect(() => {
    const mf = mathfieldRef.current;
    if (!mf) return;

    // Set offline fonts directory
    if (typeof (window as any).MathfieldElement !== 'undefined') {
      (window as any).MathfieldElement.fontsDirectory = './assets/mathlive-fonts/';
      (window as any).MathfieldElement.soundsDirectory = null;
    }

    if (mf.mathVirtualKeyboardPolicy !== undefined) {
      mf.mathVirtualKeyboardPolicy = 'manual';
    }

    try {
      (mf as any).mathModeSpace = '\\ ';
    } catch { /* ignore */ }

    const handleKeyDown = (e: KeyboardEvent) => {
      if (e.key === ' ' && !e.ctrlKey && !e.metaKey && !e.altKey) {
        e.preventDefault();
        if (typeof mf.executeCommand === 'function') {
          mf.executeCommand(['insert', '\\ ']);
        } else if (typeof mf.insert === 'function') {
          mf.insert('\\ ');
        }
      }
    };

    const handleInput = (e: any) => {
      const target = e.target;
      if (target && typeof target.getValue === 'function') {
        const val = target.getValue('latex-expanded') || target.value || '';
        onLatexChange(val);
      } else if (target && typeof target.value === 'string') {
        onLatexChange(target.value);
      }
    };

    mf.addEventListener('keydown', handleKeyDown);
    mf.addEventListener('input', handleInput);

    return () => {
      mf.removeEventListener('keydown', handleKeyDown);
      mf.removeEventListener('input', handleInput);
    };
  }, [onLatexChange]);

  // Synchronize latex into mathfield only when updated from outside mathfield
  useEffect(() => {
    const mf = mathfieldRef.current;
    if (!mf) return;

    if (activeSource !== 'mathfield') {
      const currentVal = typeof mf.getValue === 'function' ? mf.getValue('latex-expanded') : mf.value;
      if (normalizeLatex(currentVal) !== normalizeLatex(latex)) {
        if (typeof mf.setValue === 'function') {
          mf.setValue(latex || '', { suppressChangeNotifications: true });
        } else {
          mf.value = latex || '';
        }
      }
    }
  }, [latex, activeSource]);

  const toggleVirtualKeyboard = () => {
    try {
      const win = window as any;
      if (win.mathVirtualKeyboard) {
        if (win.mathVirtualKeyboard.visible) {
          win.mathVirtualKeyboard.hide();
          setKeyboardVisible(false);
        } else {
          win.mathVirtualKeyboard.show();
          setKeyboardVisible(true);
        }
      } else {
        setKeyboardVisible(!keyboardVisible);
      }
    } catch {
      setKeyboardVisible(!keyboardVisible);
    }
  };

  return (
    <div className={styles.container}>
      <div className={styles.toolbarRow}>
        <div className={styles.labelRow}>
          <Caption1 className={styles.label}>Visual Math Editor (MathLive)</Caption1>
          <Badge appearance="tint" color={isValid ? 'informative' : 'danger'} size="small">
            {isValid ? 'Interactive' : 'Syntax Error'}
          </Badge>
        </div>

        <div style={{ display: 'flex', gap: '4px' }}>
          <Button
            size="small"
            appearance={keyboardVisible ? 'primary' : 'subtle'}
            icon={<Keyboard20Regular />}
            onClick={toggleVirtualKeyboard}
            aria-label="Toggle Math Virtual Keyboard"
            title="Toggle Math Virtual Keyboard"
          >
            Keyboard
          </Button>

          {onClear && (
            <Button
              size="small"
              appearance="subtle"
              icon={<Delete20Regular />}
              onClick={onClear}
              aria-label="Clear Equation"
              title="Clear Equation"
            />
          )}
        </div>
      </div>

      <div className={styles.editorBox}>
        <math-field
          ref={mathfieldRef}
          className={styles.mathfield}
          virtual-keyboard-mode="manual"
          tabIndex={0}
          aria-label="Interactive Mathematical Formula Editor"
        />
      </div>
    </div>
  );
};
