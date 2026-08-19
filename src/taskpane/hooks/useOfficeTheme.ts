import { useState, useEffect, useCallback, useMemo } from 'react';
import {
  Theme,
  webLightTheme,
  webDarkTheme,
  teamsHighContrastTheme
} from '@fluentui/react-components';

export type AppThemeMode = 'light' | 'dark' | 'contrast' | 'auto';

export interface OfficeThemeHookResult {
  theme: Theme;
  themeMode: AppThemeMode;
  resolvedMode: 'light' | 'dark' | 'contrast';
  isDark: boolean;
  setThemeMode: (mode: AppThemeMode) => void;
  officeTheme: any;
}

/**
 * Checks if a color hex or rgb is dark based on perceived luminance.
 */
function isColorDark(colorStr?: string): boolean {
  if (!colorStr) return false;
  const hex = colorStr.replace('#', '').trim();
  if (hex.length === 6) {
    const r = parseInt(hex.substring(0, 2), 16);
    const g = parseInt(hex.substring(2, 4), 16);
    const b = parseInt(hex.substring(4, 6), 16);
    const luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255;
    return luminance < 0.5;
  }
  return false;
}

/**
 * Resolves current Office theme into 'light', 'dark', or 'contrast'.
 */
export function resolveOfficeThemeMode(officeTheme?: any): 'light' | 'dark' | 'contrast' {
  if (!officeTheme) {
    // Check OS media query
    if (typeof window !== 'undefined' && window.matchMedia) {
      if (window.matchMedia('(forced-colors: active)').matches) {
        return 'contrast';
      }
      if (window.matchMedia('(prefers-color-scheme: dark)').matches) {
        return 'dark';
      }
    }
    return 'light';
  }

  if (officeTheme.isDark === true) {
    return 'dark';
  }

  if (officeTheme.bodyBackgroundColor) {
    if (isColorDark(officeTheme.bodyBackgroundColor)) {
      return 'dark';
    }
  }

  return 'light';
}

/**
 * Custom hook to dynamically monitor Office.js theme changes and provide Fluent UI v9 theme tokens.
 */
export function useOfficeTheme(defaultMode: AppThemeMode = 'auto'): OfficeThemeHookResult {
  const [themeMode, setThemeMode] = useState<AppThemeMode>(defaultMode);
  const [officeTheme, setOfficeTheme] = useState<any>(() => {
    if (typeof Office !== 'undefined' && Office?.context?.officeTheme) {
      return Office.context.officeTheme;
    }
    return null;
  });

  // Listen to Office Theme updates
  useEffect(() => {
    const checkOfficeTheme = () => {
      if (typeof Office !== 'undefined' && Office?.context?.officeTheme) {
        setOfficeTheme({ ...Office.context.officeTheme });
      }
    };

    checkOfficeTheme();

    // Polling or event listener fallback for Office environment
    const interval = setInterval(checkOfficeTheme, 1000);

    return () => {
      clearInterval(interval);
    };
  }, []);

  const resolvedMode: 'light' | 'dark' | 'contrast' = useMemo(() => {
    if (themeMode === 'auto') {
      return resolveOfficeThemeMode(officeTheme);
    }
    return themeMode;
  }, [themeMode, officeTheme]);

  const theme: Theme = useMemo(() => {
    switch (resolvedMode) {
      case 'dark':
        return webDarkTheme;
      case 'contrast':
        return teamsHighContrastTheme;
      case 'light':
      default:
        return webLightTheme;
    }
  }, [resolvedMode]);

  const isDark = resolvedMode === 'dark';

  const handleSetThemeMode = useCallback((mode: AppThemeMode) => {
    setThemeMode(mode);
  }, []);

  return {
    theme,
    themeMode,
    resolvedMode,
    isDark,
    setThemeMode: handleSetThemeMode,
    officeTheme
  };
}
