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
  Checkmark20Regular,
  TableSimple20Regular,
  Delete20Regular
} from '@fluentui/react-icons';
import { excelService, ReadCellResult } from '../services/excelService';
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
    display: 'grid',
    gridTemplateColumns: '1fr 1fr',
    ...shorthands.gap('8px'),
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
  onReadCellSuccess?: (result: ReadCellResult | string) => void;
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

  const handleInsertInCellImage = async () => {
    if (!latex.trim() || !isValid) {
      showNotification('error', 'Please enter a valid LaTeX formula first.');
      return;
    }
    setLoadingAction('insert-cell-image');
    try {
      await excelService.insertInCellImageToActiveCell(latex, options);
      showNotification('success', 'Inserted math equation directly into active cell!');
    } catch (err: any) {
      showNotification('error', `Failed to insert in-cell image: ${err?.message || err}`);
    } finally {
      setLoadingAction(null);
    }
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
      showNotification('success', 'Placed high-resolution equation shape on sheet!');
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
        showNotification('success', 'Equation loaded from active cell into editor!');
      } else {
        showNotification('info', 'Selected cell or shape has no formula or equation.');
      }
    } catch (err: any) {
      showNotification('error', `Failed to read cell: ${err?.message || err}`);
    } finally {
      setLoadingAction(null);
    }
  };

  const handleBatchConvert = async () => {
    setLoadingAction('batch-convert');
    try {
      const res = await excelService.batchConvertSelectedRange('image', options);
      if (res.total > 0) {
        showNotification('success', `Batch converted ${res.converted} of ${res.total} equations!`);
      } else {
        showNotification('info', 'Select a range of cells containing LaTeX text first.');
      }
    } catch (err: any) {
      showNotification('error', `Batch conversion error: ${err?.message || err}`);
    } finally {
      setLoadingAction(null);
    }
  };

  const handleDeleteShapes = async () => {
    setLoadingAction('delete-shapes');
    try {
      const count = await excelService.deleteShapesInSelection();
      showNotification('info', `Removed ${count} equation shape(s) in selection.`);
    } catch (err: any) {
      showNotification('error', `Failed to remove shapes: ${err?.message || err}`);
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
      showNotification('error', 'Failed to copy LaTeX.');
    }
  };

  const handleCopyMathML = async () => {
    if (!latex.trim() || !isValid) return;
    const ok = await clipboardService.copyMathML(latex);
    if (ok) {
      showNotification('success', 'Native MathML copied! Paste directly into Microsoft Word.');
    } else {
      showNotification('error', 'Failed to copy MathML.');
    }
  };

  const handleCopyPng = async () => {
    if (!latex.trim() || !isValid) return;
    try {
      const res = await clipboardService.copyEquation(latex, options);
      if (res.png) {
        showNotification('success', '4x Retina PNG copied! Paste directly into PowerPoint.');
      } else {
        showNotification('error', 'Failed to copy PNG image.');
      }
    } catch (err: any) {
      showNotification('error', `Error copying PNG: ${err?.message || err}`);
    }
  };

  const handleCopySvg = async () => {
    if (!latex.trim() || !isValid) return;
    try {
      const res = await clipboardService.copyEquation(latex, options);
      if (res.svg) {
        showNotification('success', 'SVG Vector copied to clipboard!');
      } else {
        showNotification('error', 'Failed to copy SVG.');
      }
    } catch (err: any) {
      showNotification('error', `Error copying SVG: ${err?.message || err}`);
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

      {/* Row 1: Primary Insertion Actions */}
      <div className={styles.primaryRow}>
        <Tooltip content="Insert native in-cell picture into active cell" relationship="description">
          <Button
            appearance="primary"
            size="medium"
            icon={loadingAction === 'insert-cell-image' ? <Spinner size="tiny" /> : <Image20Regular />}
            onClick={handleInsertInCellImage}
            disabled={!isValid || !latex.trim() || loadingAction !== null}
            className={styles.fullWidthBtn}
          >
            In-Cell Image
          </Button>
        </Tooltip>

        <Tooltip content="Insert =MATH.KATEX() dynamic custom function" relationship="description">
          <Button
            appearance="outline"
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
        <Tooltip content="Insert draggable floating shape on worksheet" relationship="description">
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

        <Tooltip content="Read equation or metadata from selected cell/shape" relationship="description">
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

      {/* Row 3: Batch Tools & Export */}
      <div className={styles.copyRow}>
        <Tooltip content="Batch convert selected range of LaTeX text to equations" relationship="description">
          <Button
            appearance="subtle"
            icon={<TableSimple20Regular />}
            onClick={handleBatchConvert}
            disabled={loadingAction !== null}
          >
            Batch Convert
          </Button>
        </Tooltip>

        <Tooltip content="Remove floating math shapes in selection" relationship="description">
          <Button
            appearance="subtle"
            icon={<Delete20Regular />}
            onClick={handleDeleteShapes}
            disabled={loadingAction !== null}
          >
            Clear Shapes
          </Button>
        </Tooltip>

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
              <MenuItem icon={<Checkmark20Regular />} onClick={handleCopyMathML}>
                Copy as MathML (for Microsoft Word)
              </MenuItem>
              <MenuItem icon={<Image20Regular />} onClick={handleCopyPng}>
                Copy 4x PNG (for PowerPoint)
              </MenuItem>
              <MenuItem icon={<Copy20Regular />} onClick={handleCopyLatex}>
                Copy Raw LaTeX Code
              </MenuItem>
              <MenuItem icon={<Image20Regular />} onClick={handleCopySvg}>
                Copy Standalone SVG
              </MenuItem>
            </MenuList>
          </MenuPopover>
        </Menu>
      </div>
    </div>
  );
};
