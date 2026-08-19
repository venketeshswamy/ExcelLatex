import React, { useState } from 'react';
import {
  makeStyles,
  shorthands,
  Button,
  Menu,
  MenuTrigger,
  MenuPopover,
  MenuList,
  MenuItem,
  Spinner,
  MessageBar,
  MessageBarBody,
  MessageBarTitle,
  Tooltip
} from '@fluentui/react-components';
import {
  MathFormula20Regular,
  Image20Regular,
  Shapes20Regular,
  ArrowDownload20Regular,
  Copy20Regular,
  DocumentCopy20Regular,
  Checkmark20Regular
} from '@fluentui/react-icons';
import { excelService } from '../services/excelService';
import { clipboardService } from '../services/clipboardService';
import { RenderOptions } from '../../core/katexEngine';

const useStyles = makeStyles({
  container: {
    display: 'flex',
    flexDirection: 'column',
    ...shorthands.gap('10px'),
    width: '100%'
  },
  primaryRow: {
    display: 'flex',
    width: '100%'
  },
  secondaryRow: {
    display: 'grid',
    gridTemplateColumns: '1fr 1fr',
    ...shorthands.gap('8px')
  },
  copyRow: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    ...shorthands.gap('8px')
  },
  fullWidthBtn: {
    width: '100%'
  }
});

export interface ActionBarProps {
  latex: string;
  isValid: boolean;
  options?: RenderOptions;
  onReadCellSuccess?: (extractedLatex: string) => void;
}

export const ActionBar: React.FC<ActionBarProps> = ({
  latex,
  isValid,
  options = {},
  onReadCellSuccess
}) => {
  const styles = useStyles();
  const [loadingAction, setLoadingAction] = useState<string | null>(null);
  const [notification, setNotification] = useState<{
    type: 'success' | 'error' | 'info';
    message: string;
  } | null>(null);

  const showNotification = (
    type: 'success' | 'error' | 'info',
    message: string
  ) => {
    setNotification({ type, message });
    setTimeout(() => {
      setNotification(null);
    }, 4000);
  };

  const handleInsertFormula = async () => {
    if (!latex.trim() || !isValid) {
      showNotification('error', 'Please enter a valid LaTeX formula first.');
      return;
    }
    setLoadingAction('insert-formula');
    try {
      await excelService.insertFormulaToActiveCell(latex, options);
      showNotification('success', 'Inserted =MATH.KATEX() formula into active cell!');
    } catch (err: any) {
      showNotification('error', `Failed to insert formula: ${err?.message || err}`);
    } finally {
      setLoadingAction(null);
    }
  };

  const handleInsertShape = async () => {
    if (!latex.trim() || !isValid) {
      showNotification('error', 'Please enter a valid LaTeX formula first.');
      return;
    }
    setLoadingAction('insert-shape');
    try {
      await excelService.insertFloatingShapeToActiveCell(latex, options);
      showNotification('success', 'Inserted floating high-DPI shape into sheet!');
    } catch (err: any) {
      showNotification('error', `Failed to insert shape: ${err?.message || err}`);
    } finally {
      setLoadingAction(null);
    }
  };

  const handleReadCell = async () => {
    setLoadingAction('read-cell');
    try {
      const extracted = await excelService.readActiveCellFormula();
      if (extracted) {
        if (onReadCellSuccess) {
          onReadCellSuccess(extracted);
        }
        showNotification('success', 'Read LaTeX formula from selected cell!');
      } else {
        showNotification('info', 'Active cell has no formula or text.');
      }
    } catch (err: any) {
      showNotification('error', `Failed to read cell: ${err?.message || err}`);
    } finally {
      setLoadingAction(null);
    }
  };

  const handleCopyLatex = async () => {
    if (!latex.trim()) return;
    const ok = await clipboardService.copyLatex(latex);
    if (ok) {
      showNotification('success', 'LaTeX code copied to clipboard!');
    } else {
      showNotification('error', 'Failed to copy LaTeX to clipboard.');
    }
  };

  const handleCopySvg = async () => {
    if (!latex.trim() || !isValid) return;
    try {
      const res = await clipboardService.copyEquation(latex, options);
      if (res.svg) {
        showNotification('success', 'Standalone SVG copied to clipboard!');
      } else {
        showNotification('error', 'Failed to copy SVG.');
      }
    } catch (err: any) {
      showNotification('error', `Error copying SVG: ${err?.message || err}`);
    }
  };

  const handleCopyPng = async () => {
    if (!latex.trim() || !isValid) return;
    try {
      const res = await clipboardService.copyEquation(latex, options);
      if (res.png) {
        showNotification('success', 'High-DPI PNG copied to clipboard!');
      } else {
        showNotification('error', 'Failed to copy PNG image.');
      }
    } catch (err: any) {
      showNotification('error', `Error copying PNG: ${err?.message || err}`);
    }
  };

  return (
    <div className={styles.container}>
      {notification && (
        <MessageBar
          intent={notification.type === 'error' ? 'error' : notification.type === 'info' ? 'info' : 'success'}
        >
          <MessageBarBody>
            <MessageBarTitle>
              {notification.type === 'error' ? 'Error' : 'Notice'}
            </MessageBarTitle>
            {notification.message}
          </MessageBarBody>
        </MessageBar>
      )}

      {/* Row 1: Primary Sheet Insert Action */}
      <div className={styles.primaryRow}>
        <Tooltip content="Insert =MATH.KATEX() dynamic custom function into active Excel cell" relationship="description">
          <Button
            appearance="primary"
            size="medium"
            icon={loadingAction === 'insert-formula' ? <Spinner size="tiny" /> : <MathFormula20Regular />}
            onClick={handleInsertFormula}
            disabled={!isValid || !latex.trim() || loadingAction !== null}
            className={styles.fullWidthBtn}
          >
            Insert Formula
          </Button>
        </Tooltip>
      </div>

      {/* Row 2: Secondary Shape & Cell Actions */}
      <div className={styles.secondaryRow}>
        <Tooltip content="Insert floating vector/raster shape anchored to active cell" relationship="description">
          <Button
            appearance="outline"
            icon={loadingAction === 'insert-shape' ? <Spinner size="tiny" /> : <Shapes20Regular />}
            onClick={handleInsertShape}
            disabled={!isValid || !latex.trim() || loadingAction !== null}
            className={styles.fullWidthBtn}
          >
            Floating Shape
          </Button>
        </Tooltip>

        <Tooltip content="Read LaTeX formula or entity from active cell" relationship="description">
          <Button
            appearance="outline"
            icon={loadingAction === 'read-cell' ? <Spinner size="tiny" /> : <ArrowDownload20Regular />}
            onClick={handleReadCell}
            disabled={loadingAction !== null}
            className={styles.fullWidthBtn}
          >
            Read Cell
          </Button>
        </Tooltip>
      </div>

      {/* Row 3: Clipboard Dropdown */}
      <div className={styles.copyRow}>
        <Button
          appearance="subtle"
          icon={<Copy20Regular />}
          onClick={handleCopyLatex}
          disabled={!latex.trim()}
          style={{ flex: 1 }}
        >
          Copy LaTeX
        </Button>

        <Menu>
          <MenuTrigger disableButtonEnhancement>
            <Button
              appearance="subtle"
              icon={<DocumentCopy20Regular />}
              disabled={!latex.trim() || !isValid}
            >
              Export...
            </Button>
          </MenuTrigger>
          <MenuPopover>
            <MenuList>
              <MenuItem icon={<Copy20Regular />} onClick={handleCopyLatex}>
                Copy LaTeX Code
              </MenuItem>
              <MenuItem icon={<Image20Regular />} onClick={handleCopySvg}>
                Copy SVG Vector Markup
              </MenuItem>
              <MenuItem icon={<Checkmark20Regular />} onClick={handleCopyPng}>
                Copy High-DPI PNG Image
              </MenuItem>
            </MenuList>
          </MenuPopover>
        </Menu>
      </div>
    </div>
  );
};
