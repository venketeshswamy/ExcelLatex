import React, { useState, useMemo } from 'react';
import {
  makeStyles,
  tokens,
  shorthands,
  Card,
  Caption1,
  Badge,
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
  compileLatexToHtml,
  calculateComplexity,
  LatexComplexity
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
  displayMode = true,
  onDisplayModeChange,
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
    return compileLatexToHtml(latex, { displayMode, throwOnError: false });
  }, [latex, displayMode]);

  const complexity: LatexComplexity = useMemo(() => {
    return calculateComplexity(latex);
  }, [latex]);

  const getComplexityColor = (score: string) => {
    switch (score) {
      case 'Simple':
        return 'success';
      case 'Moderate':
        return 'informative';
      case 'Complex':
        return 'warning';
      case 'Advanced':
        return 'danger';
      default:
        return 'informative';
    }
  };

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
          {onDisplayModeChange && (
            <Dropdown
              aria-label="Display Mode"
              value={displayMode ? 'Display Mode (Block)' : 'Inline Mode'}
              selectedOptions={[displayMode ? 'true' : 'false']}
              onOptionSelect={(_e, data) => {
                if (data.optionValue !== undefined && onDisplayModeChange) {
                  onDisplayModeChange(data.optionValue === 'true');
                }
              }}
              size="small"
              className={styles.dropdown}
            >
              <Option value="true" text="Display Mode (Block)">Display Mode (Block)</Option>
              <Option value="false" text="Inline Mode">Inline Mode</Option>
            </Dropdown>
          )}

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
          <Tooltip content="Zoom Out" relationship="description">
            <Button
              size="small"
              appearance="subtle"
              icon={<ZoomOut20Regular />}
              onClick={() => setZoom((z) => Math.max(50, z - 25))}
              disabled={zoom <= 50}
              aria-label="Zoom out preview"
            />
          </Tooltip>

          <Slider
            min={50}
            max={200}
            step={10}
            value={zoom}
            onChange={handleZoomChange}
            style={{ width: '100px' }}
            aria-label="Zoom percentage"
          />

          <Tooltip content="Zoom In" relationship="description">
            <Button
              size="small"
              appearance="subtle"
              icon={<ZoomIn20Regular />}
              onClick={() => setZoom((z) => Math.min(200, z + 25))}
              disabled={zoom >= 200}
              aria-label="Zoom in preview"
            />
          </Tooltip>

          <Tooltip content="Reset Zoom (100%)" relationship="description">
            <Button
              size="small"
              appearance="subtle"
              icon={<ArrowReset20Regular />}
              onClick={handleResetZoom}
              aria-label="Reset zoom"
            >
              {zoom}%
            </Button>
          </Tooltip>
        </div>
      </div>

      {!isValid && errorMessage ? (
        <div className={styles.errorBanner} role="alert">
          <Warning20Regular />
          <div>{errorMessage}</div>
        </div>
      ) : (
        <Card className={styles.previewCard} style={previewCardStyle} role="region" aria-label="Equation Preview">
          {latex.trim() ? (
            <div
              className={styles.renderedWrapper}
              style={{
                transform: `scale(${zoom / 100})`,
                transformOrigin: 'center center',
                color: effectiveTextColor
              }}
              dangerouslySetInnerHTML={{ __html: compiledHtml }}
            />
          ) : (
            <div style={{ color: tokens.colorNeutralForeground4, fontStyle: 'italic' }}>
              Equation preview will appear here...
            </div>
          )}
        </Card>
      )}

      {latex.trim() && (
        <div className={styles.metricsBar}>
          <span>
            {isValid ? (
              <span style={{ color: tokens.colorPaletteGreenForeground1 }}>
                <CheckmarkCircle16Filled style={{ verticalAlign: 'middle', marginRight: '4px' }} />
                Syntax Clean
              </span>
            ) : (
              'Syntax Error'
            )}
          </span>
          <span>Scale: {zoom}%</span>
        </div>
      )}
    </div>
  );
};
