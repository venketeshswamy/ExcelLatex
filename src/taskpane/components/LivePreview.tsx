import React, { useState, useMemo } from 'react';
import {
  makeStyles,
  tokens,
  shorthands,
  Caption1,
  Slider,
  Button,
  Dropdown,
  Option,
  Tooltip
} from '@fluentui/react-components';
import {
  ZoomIn20Regular,
  ZoomOut20Regular,
  ArrowReset20Regular,
  Warning20Regular,
  CheckmarkCircle16Filled
} from '@fluentui/react-icons';
import {
  compileLatexToHtml
} from '../../core/katexEngine';

const useStyles = makeStyles({
  container: {
    display: 'flex',
    flexDirection: 'column',
    ...shorthands.gap('8px'),
    width: '100%'
  },
  headerRow: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    flexWrap: 'wrap',
    ...shorthands.gap('6px')
  },
  labelGroup: {
    display: 'flex',
    alignItems: 'center',
    ...shorthands.gap('6px')
  },
  label: {
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground1
  },
  dropdownsGroup: {
    display: 'flex',
    alignItems: 'center',
    ...shorthands.gap('6px'),
    flexWrap: 'wrap'
  },
  dropdown: {
    minWidth: '130px'
  },
  controlsRow: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    fontSize: '12px',
    color: tokens.colorNeutralForeground2,
    ...shorthands.padding('2px', '0')
  },
  zoomControls: {
    display: 'flex',
    alignItems: 'center',
    ...shorthands.gap('4px')
  },
  previewCard: {
    minHeight: '100px',
    maxHeight: '260px',
    overflowY: 'auto',
    overflowX: 'auto',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    ...shorthands.padding('16px'),
    backgroundColor: tokens.colorNeutralBackground1,
    ...shorthands.border('1px', 'solid', tokens.colorNeutralStroke1),
    ...shorthands.borderRadius(tokens.borderRadiusMedium),
    boxShadow: tokens.shadow2
  },
  renderedWrapper: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    width: '100%',
    maxWidth: '100%',
    overflowX: 'auto',
    transitionProperty: 'transform',
    transitionDuration: '0.15s',
    transitionTimingFunction: 'ease-out'
  },
  metricsBar: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    fontSize: '11px',
    color: tokens.colorNeutralForeground3,
    ...shorthands.padding('2px', '4px')
  },
  errorBanner: {
    display: 'flex',
    alignItems: 'center',
    ...shorthands.gap('8px'),
    ...shorthands.padding('10px', '12px'),
    ...shorthands.borderRadius(tokens.borderRadiusMedium),
    backgroundColor: tokens.colorPaletteRedBackground1,
    ...shorthands.border('1px', 'solid', tokens.colorPaletteRedBorder1),
    color: tokens.colorPaletteRedForeground1,
    fontSize: '12px'
  }
});

export interface LivePreviewProps {
  latex: string;
  isValid: boolean;
  errorMessage?: string;
  displayMode?: boolean;
  onDisplayModeChange?: (displayMode: boolean) => void;
  background?: number | string;
  onBackgroundChange?: (background: number | string) => void;
  textColor?: string;
}

export function getBackgroundLabel(bg?: number | string): string {
  if (bg === 0 || bg === '0' || bg === 'transparent' || bg === undefined) return '0: Transparent';
  if (bg === 1 || bg === '1' || bg === 'white' || bg === '#ffffff') return '1: White';
  if (bg === 2 || bg === '2' || bg === 'black' || bg === '#000000') return '2: Black';
  return String(bg);
}

export const LivePreview: React.FC<LivePreviewProps> = ({
  latex,
  isValid,
  errorMessage,
  background = 0,
  onBackgroundChange,
  textColor = '#000000'
}) => {
  const styles = useStyles();
  const [zoom, setZoom] = useState<number>(100); // 50% to 200%

  const isBlackBg = background === 2 || background === '2' || background === 'black' || background === '#000000';
  const isWhiteBg = background === 1 || background === '1' || background === 'white' || background === '#ffffff';

  const effectiveTextColor = useMemo(() => {
    if (isBlackBg) {
      return textColor && textColor !== '#000000' ? textColor : '#ffffff';
    }
    return textColor || '#000000';
  }, [isBlackBg, textColor]);

  const compiledHtml = useMemo(() => {
    if (!latex || !latex.trim()) return '';
    return compileLatexToHtml(latex, { displayMode: true, throwOnError: false });
  }, [latex]);

  const handleZoomChange = (_ev: any, data: { value: number }) => {
    setZoom(data.value);
  };

  const handleResetZoom = () => {
    setZoom(100);
  };

  const previewCardStyle = useMemo<React.CSSProperties>(() => {
    if (isBlackBg) {
      return {
        backgroundColor: '#000000',
        color: effectiveTextColor,
        borderColor: '#333333'
      };
    }
    if (isWhiteBg) {
      return {
        backgroundColor: '#ffffff',
        color: effectiveTextColor
      };
    }
    return {
      color: effectiveTextColor
    };
  }, [isBlackBg, isWhiteBg, effectiveTextColor]);

  return (
    <div className={styles.container}>
      <div className={styles.headerRow}>
        <div className={styles.labelGroup}>
          <Caption1 className={styles.label}>Live Preview</Caption1>
        </div>

        <div className={styles.dropdownsGroup}>
          {onBackgroundChange && (
            <Dropdown
              aria-label="Background"
              value={getBackgroundLabel(background)}
              selectedOptions={[String(background ?? 0)]}
              onOptionSelect={(_e, data) => {
                if (data.optionValue !== undefined && onBackgroundChange) {
                  onBackgroundChange(Number(data.optionValue));
                }
              }}
              size="small"
              className={styles.dropdown}
            >
              <Option value="0" text="0: Transparent">0: Transparent</Option>
              <Option value="1" text="1: White">1: White</Option>
              <Option value="2" text="2: Black">2: Black</Option>
            </Dropdown>
          )}
        </div>
      </div>

      <div className={styles.controlsRow}>
        <div className={styles.zoomControls}>
          <Tooltip content="Zoom Out" relationship="label">
            <Button
              appearance="subtle"
              icon={<ZoomOut20Regular />}
              size="small"
              onClick={() => setZoom((z) => Math.max(50, z - 10))}
              aria-label="Zoom Out"
            />
          </Tooltip>

          <Slider
            min={50}
            max={200}
            step={5}
            value={zoom}
            onChange={handleZoomChange}
            style={{ width: '120px' }}
            aria-label="Zoom slider"
          />

          <Tooltip content="Zoom In" relationship="label">
            <Button
              appearance="subtle"
              icon={<ZoomIn20Regular />}
              size="small"
              onClick={() => setZoom((z) => Math.min(200, z + 10))}
              aria-label="Zoom In"
            />
          </Tooltip>

          <Tooltip content="Reset Zoom" relationship="label">
            <Button
              appearance="subtle"
              icon={<ArrowReset20Regular />}
              size="small"
              onClick={handleResetZoom}
              aria-label="Reset Zoom"
            />
          </Tooltip>

          <span>{zoom}%</span>
        </div>
      </div>

      {!isValid && errorMessage && (
        <div className={styles.errorBanner} role="alert">
          <Warning20Regular />
          <span>{errorMessage}</span>
        </div>
      )}

      <div className={styles.previewCard} style={previewCardStyle} role="region" aria-label="Equation Preview">
        {compiledHtml ? (
          <div
            className={styles.renderedWrapper}
            style={{
              transform: `scale(${zoom / 100})`,
              transformOrigin: 'center center'
            }}
            dangerouslySetInnerHTML={{ __html: compiledHtml }}
          />
        ) : (
          <span style={{ color: tokens.colorNeutralForeground4, fontStyle: 'italic' }}>
            No equation to preview
          </span>
        )}
      </div>

      <div className={styles.metricsBar}>
        <span style={{ display: 'flex', alignItems: 'center', gap: '4px' }}>
          {isValid ? (
            <>
              <CheckmarkCircle16Filled style={{ color: tokens.colorPaletteGreenForeground1 }} />
              <span>Syntax Clean</span>
            </>
          ) : (
            <span style={{ color: tokens.colorPaletteRedForeground1 }}>Syntax Error</span>
          )}
        </span>
        <span>Scale: {zoom}%</span>
      </div>
    </div>
  );
};
