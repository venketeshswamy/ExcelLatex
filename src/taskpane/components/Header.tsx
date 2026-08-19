import React from 'react';
import {
  Title3,
  Caption1,
  Badge,
  Button,
  Menu,
  MenuTrigger,
  MenuPopover,
  MenuList,
  MenuItemRadio,
  MenuProps,
  makeStyles,
  tokens,
  shorthands
} from '@fluentui/react-components';
import {
  WeatherMoon20Regular,
  WeatherSunny20Regular,
  DarkTheme20Regular,
  Desktop20Regular,
  CheckmarkCircle16Filled,
  Info16Regular
} from '@fluentui/react-icons';
import { AppThemeMode } from '../hooks/useOfficeTheme';

const useStyles = makeStyles({
  container: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    ...shorthands.padding('12px', '16px'),
    ...shorthands.borderBottom('1px', 'solid', tokens.colorNeutralStroke2),
    backgroundColor: tokens.colorNeutralBackground1
  },
  leftGroup: {
    display: 'flex',
    flexDirection: 'column',
    ...shorthands.gap('2px')
  },
  titleRow: {
    display: 'flex',
    alignItems: 'center',
    ...shorthands.gap('8px')
  },
  titleText: {
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground1
  },
  rightGroup: {
    display: 'flex',
    alignItems: 'center',
    ...shorthands.gap('8px')
  }
});

export interface HeaderProps {
  themeMode: AppThemeMode;
  resolvedMode: 'light' | 'dark' | 'contrast';
  onThemeModeChange: (mode: AppThemeMode) => void;
}

export const Header: React.FC<HeaderProps> = ({
  themeMode,
  resolvedMode,
  onThemeModeChange
}) => {
  const styles = useStyles();

  const isExcelConnected =
    typeof Office !== 'undefined' &&
    typeof Excel !== 'undefined' &&
    (Office?.context?.host === Office.HostType.Excel || (Office?.context?.host as any) === 'Excel');

  const handleCheckedValueChange: MenuProps['onCheckedValueChange'] = (
    _e,
    data
  ) => {
    if (data.name === 'theme' && data.checkedItems.length > 0) {
      onThemeModeChange(data.checkedItems[0] as AppThemeMode);
    }
  };

  const getThemeIcon = () => {
    switch (resolvedMode) {
      case 'dark':
        return <WeatherMoon20Regular />;
      case 'contrast':
        return <DarkTheme20Regular />;
      case 'light':
      default:
        return <WeatherSunny20Regular />;
    }
  };

  return (
    <header className={styles.container} role="banner">
      <div className={styles.leftGroup}>
        <div className={styles.titleRow}>
          <Title3 className={styles.titleText}>LaTeX Math</Title3>
          <Badge
            appearance="tint"
            color={isExcelConnected ? 'success' : 'informative'}
            icon={isExcelConnected ? <CheckmarkCircle16Filled /> : <Info16Regular />}
            size="small"
          >
            {isExcelConnected ? 'Excel Ready' : 'Local Offline'}
          </Badge>
        </div>
        <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
          Equation Editor
        </Caption1>
      </div>

      <div className={styles.rightGroup}>
        <Menu
          checkedValues={{ theme: [themeMode] }}
          onCheckedValueChange={handleCheckedValueChange}
        >
          <MenuTrigger disableButtonEnhancement>
            <Button
              appearance="subtle"
              icon={getThemeIcon()}
              aria-label="Select color theme"
              title={`Current Theme: ${themeMode} (${resolvedMode})`}
            />
          </MenuTrigger>
          <MenuPopover>
            <MenuList>
              <MenuItemRadio name="theme" value="auto" icon={<Desktop20Regular />}>
                System / Office Default
              </MenuItemRadio>
              <MenuItemRadio name="theme" value="light" icon={<WeatherSunny20Regular />}>
                Light Theme
              </MenuItemRadio>
              <MenuItemRadio name="theme" value="dark" icon={<WeatherMoon20Regular />}>
                Dark Theme
              </MenuItemRadio>
              <MenuItemRadio name="theme" value="contrast" icon={<DarkTheme20Regular />}>
                High Contrast
              </MenuItemRadio>
            </MenuList>
          </MenuPopover>
        </Menu>
      </div>
    </header>
  );
};
