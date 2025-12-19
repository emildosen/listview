import { createContext, useContext, useEffect, useState, useMemo } from 'react';
import type { ReactNode } from 'react';
import {
  FluentProvider,
  webLightTheme,
  type Theme,
  type BrandVariants,
  createDarkTheme,
} from '@fluentui/react-components';

// Custom brand colors (using the default blue brand)
const brandColors: BrandVariants = {
  10: '#061724',
  20: '#082338',
  30: '#0a2e4a',
  40: '#0c3b5e',
  50: '#0e4775',
  60: '#0f548c',
  70: '#115ea3',
  80: '#0f6cbd',
  90: '#2886de',
  100: '#479ef5',
  110: '#62abf5',
  120: '#77b7f7',
  130: '#96c6fa',
  140: '#b4d6fa',
  150: '#cfe4fa',
  160: '#ebf3fc',
};

// Custom dark theme with darker backgrounds
const customDarkTheme: Theme = {
  ...createDarkTheme(brandColors),
  colorNeutralBackground1: '#1a1a1a',  // Main page background - dark gray
  colorNeutralBackground2: '#121212',  // Sidebar background - almost black
  colorNeutralBackground3: '#252525',  // Hover states
  colorNeutralBackground4: '#2a2a2a',
  colorNeutralBackground5: '#303030',
  colorNeutralBackground6: '#383838',
};

type ThemeMode = 'light' | 'dark';

interface ThemeContextType {
  theme: ThemeMode;
  setTheme: (theme: ThemeMode) => void;
  toggleTheme: () => void;
  fluentTheme: Theme;
}

const ThemeContext = createContext<ThemeContextType | undefined>(undefined);

const STORAGE_KEY = 'listview-theme';

function getInitialTheme(): ThemeMode {
  const stored = localStorage.getItem(STORAGE_KEY);
  if (stored === 'light' || stored === 'dark') {
    return stored;
  }
  // Check system preference
  if (window.matchMedia('(prefers-color-scheme: dark)').matches) {
    return 'dark';
  }
  return 'light';
}

export function ThemeProvider({ children }: { children: ReactNode }) {
  const [theme, setThemeState] = useState<ThemeMode>(getInitialTheme);

  useEffect(() => {
    localStorage.setItem(STORAGE_KEY, theme);
    // Apply theme class to document for global CSS styling (scrollbars, etc.)
    document.documentElement.classList.toggle('dark-theme', theme === 'dark');
  }, [theme]);

  const fluentTheme = useMemo(
    () => (theme === 'dark' ? customDarkTheme : webLightTheme),
    [theme]
  );

  const setTheme = (newTheme: ThemeMode) => {
    setThemeState(newTheme);
  };

  const toggleTheme = () => {
    setThemeState((prev) => (prev === 'light' ? 'dark' : 'light'));
  };

  const value = useMemo(
    () => ({ theme, setTheme, toggleTheme, fluentTheme }),
    [theme, fluentTheme]
  );

  return (
    <ThemeContext.Provider value={value}>
      <FluentProvider theme={fluentTheme} style={{ height: '100%' }}>
        {children}
      </FluentProvider>
    </ThemeContext.Provider>
  );
}

// eslint-disable-next-line react-refresh/only-export-components
export function useTheme() {
  const context = useContext(ThemeContext);
  if (context === undefined) {
    throw new Error('useTheme must be used within a ThemeProvider');
  }
  return context;
}
