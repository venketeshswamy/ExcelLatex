import React, { useState } from 'react';
import {
  FluentProvider,
  TabList,
  Tab,
  makeStyles,
  tokens,
  shorthands,
  Divider,
  Card,
  Subtitle2,
  Caption1,
  Input,
  Dropdown,
  Option,
  Button
} from '@fluentui/react-components';
import {
  Edit20Regular,
  Book20Regular,
  Settings20Regular,
  Delete20Regular
} from '@fluentui/react-icons';
import { Header } from './components/Header';
import { MathLiveEditor } from './components/MathLiveEditor';
import { RawLatexEditor } from './components/RawLatexEditor';
import { LivePreview, getBackgroundLabel } from './components/LivePreview';
import { PresetsLibrary } from './components/PresetsLibrary';
import { ActionBar } from './components/ActionBar';
import { useLatexSync } from './hooks/useLatexSync';
import { useOfficeTheme } from './hooks/useOfficeTheme';
import { getEquationCacheStats, clearEquationCache, RenderOptions } from '../core/katexEngine';

const useStyles = makeStyles({
  appRoot: {
    display: 'flex',
    flexDirection: 'column',
    height: '100vh',
    width: '100%',
    boxSizing: 'border-box',
    backgroundColor: tokens.colorNeutralBackground2,
    color: tokens.colorNeutralForeground1,
    overflow: 'hidden'
  },
  contentContainer: {
    display: 'flex',
    flexDirection: 'column',
    flex: 1,
    overflowY: 'auto',
    ...shorthands.padding('12px', '14px'),
    ...shorthands.gap('12px')
  },
  navigationBar: {
    ...shorthands.padding('0', '12px'),
    backgroundColor: tokens.colorNeutralBackground1,
    ...shorthands.borderBottom('1px', 'solid', tokens.colorNeutralStroke2)
  },
  settingsCard: {
    display: 'flex',
    flexDirection: 'column',
    ...shorthands.gap('12px'),
    ...shorthands.padding('16px'),
    backgroundColor: tokens.colorNeutralBackground1
  },
  settingRow: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    ...shorthands.gap('12px')
  }
});

export const App: React.FC = () => {
  const styles = useStyles();
  const {
    theme,
    themeMode,
    resolvedMode,
    setThemeMode
  } = useOfficeTheme('auto');

  const {
    latex,
    activeSource,
    isValid,
    errorMessage,
    updateFromMathField,
    updateFromRawLatex,
    setEquation,
    clear
  } = useLatexSync('\\int_{-\\infty}^{+\\infty} e^{-x^2} \\, dx = \\sqrt{\\pi}');

  const [selectedTab, setSelectedTab] = useState<'editor' | 'presets' | 'settings'>('editor');
  const [displayMode, setDisplayMode] = useState<boolean>(true);
  const [fontSize, setFontSize] = useState<number>(16);
  const [textColor, setTextColor] = useState<string>('#000000');
  const [bgColor, setBgColor] = useState<number | string>(0);

  const handleBgColorChange = (newBg: number | string) => {
    const numBg = typeof newBg === 'string' && ['0', '1', '2'].includes(newBg) ? Number(newBg) : newBg;
    setBgColor(numBg);
    if (numBg === 2 || numBg === '2' || numBg === 'black' || numBg === '#000000') {
      if (textColor === '#000000') {
        setTextColor('#ffffff');
      }
    } else if (numBg === 0 || numBg === '0' || numBg === 1 || numBg === '1' || numBg === 'transparent' || numBg === 'white') {
      if (textColor === '#ffffff') {
        setTextColor('#000000');
      }
    }
  };

  const renderOptions: RenderOptions = {
    displayMode,
    fontSize,
    color: textColor,
    background: bgColor as any
  };

  const handleSelectPreset = (presetLatex: string) => {
    setEquation(presetLatex, 'preset');
    setSelectedTab('editor');
  };

  const handleAppendPreset = (presetLatex: string) => {
    const updated = latex ? `${latex} \\quad ${presetLatex}` : presetLatex;
    setEquation(updated, 'preset');
    setSelectedTab('editor');
  };

  const handleReadCellSuccess = (extractedLatex: string) => {
    setEquation(extractedLatex, 'cell');
  };

  const cacheStats = getEquationCacheStats();

  return (
    <FluentProvider theme={theme}>
      <div className={styles.appRoot}>
        <Header
          themeMode={themeMode}
          resolvedMode={resolvedMode}
          onThemeModeChange={setThemeMode}
        />

        <div className={styles.navigationBar}>
          <TabList
            selectedValue={selectedTab}
            onTabSelect={(_e, data) => setSelectedTab(data.value as any)}
            size="medium"
          >
            <Tab value="editor" icon={<Edit20Regular />}>
              Editor
            </Tab>
            <Tab value="presets" icon={<Book20Regular />}>
              Presets
            </Tab>
            <Tab value="settings" icon={<Settings20Regular />}>
              Options
            </Tab>
          </TabList>
        </div>

        <main className={styles.contentContainer}>
          {selectedTab === 'editor' && (
            <>
              {/* Visual MathLive Editor */}
              <MathLiveEditor
                latex={latex}
                activeSource={activeSource}
                isValid={isValid}
                onLatexChange={updateFromMathField}
                onClear={clear}
              />

              <Divider style={{ margin: '4px 0' }} />

              {/* Raw LaTeX Code Editor */}
              <RawLatexEditor
                latex={latex}
                activeSource={activeSource}
                isValid={isValid}
                errorMessage={errorMessage}
                onLatexChange={updateFromRawLatex}
                onClear={clear}
              />

              <Divider style={{ margin: '4px 0' }} />

              {/* Live KaTeX Preview */}
              <LivePreview
                latex={latex}
                isValid={isValid}
                errorMessage={errorMessage}
                displayMode={displayMode}
                onDisplayModeChange={setDisplayMode}
                background={bgColor}
                onBackgroundChange={handleBgColorChange}
                textColor={textColor}
              />

              <Divider style={{ margin: '4px 0' }} />

              {/* Action Buttons for Sheet & Clipboard */}
              <ActionBar
                latex={latex}
                isValid={isValid}
                options={renderOptions}
                onReadCellSuccess={handleReadCellSuccess}
              />
            </>
          )}

          {selectedTab === 'presets' && (
            <PresetsLibrary
              onSelectPreset={handleSelectPreset}
              onAppendPreset={handleAppendPreset}
            />
          )}

          {selectedTab === 'settings' && (
            <Card className={styles.settingsCard}>
              <Subtitle2>Rendering &amp; Typesetting Options</Subtitle2>
              <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
                Configure default font sizing and appearance for Excel sheet insertion
              </Caption1>

              <div className={styles.settingRow}>
                <Caption1>Font Size (pt):</Caption1>
                <Input
                  type="number"
                  value={String(fontSize)}
                  onChange={(_e, d) => setFontSize(Number(d.value) || 16)}
                  style={{ width: '80px' }}
                />
              </div>

              <div className={styles.settingRow}>
                <Caption1>Display Mode:</Caption1>
                <Dropdown
                  aria-label="Display Mode Option"
                  value={displayMode ? 'Display Mode (Block)' : 'Inline Mode'}
                  selectedOptions={[displayMode ? 'true' : 'false']}
                  onOptionSelect={(_e, d) => {
                    if (d.optionValue !== undefined) {
                      setDisplayMode(d.optionValue === 'true');
                    }
                  }}
                  size="small"
                  style={{ minWidth: '170px' }}
                >
                  <Option value="true" text="Display Mode (Block)">Display Mode (Block)</Option>
                  <Option value="false" text="Inline Mode">Inline Mode</Option>
                </Dropdown>
              </div>

              <div className={styles.settingRow}>
                <Caption1>Background:</Caption1>
                <Dropdown
                  aria-label="Background Option"
                  value={getBackgroundLabel(bgColor)}
                  selectedOptions={[String(bgColor)]}
                  onOptionSelect={(_e, d) => {
                    if (d.optionValue !== undefined) {
                      handleBgColorChange(Number(d.optionValue));
                    }
                  }}
                  size="small"
                  style={{ minWidth: '170px' }}
                >
                  <Option value="0" text="0: Transparent">0: Transparent</Option>
                  <Option value="1" text="1: White">1: White</Option>
                  <Option value="2" text="2: Black">2: Black</Option>
                </Dropdown>
              </div>

              <div className={styles.settingRow}>
                <Caption1>Text Color:</Caption1>
                <Input
                  value={textColor}
                  onChange={(_e, d) => setTextColor(d.value)}
                  style={{ width: '120px' }}
                />
              </div>

              <Divider style={{ margin: '8px 0' }} />

              <Subtitle2>In-Memory LRU Cache</Subtitle2>
              <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
                Cached Equations: {cacheStats.size} / {cacheStats.maxSize} (Hits: {cacheStats.hits}, Misses: {cacheStats.misses})
              </Caption1>

              <Button
                appearance="outline"
                icon={<Delete20Regular />}
                onClick={() => clearEquationCache()}
                style={{ alignSelf: 'flex-start' }}
              >
                Clear Rendering Cache
              </Button>
            </Card>
          )}
        </main>
      </div>
    </FluentProvider>
  );
};
