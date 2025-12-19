import { createContext, useContext, useState, useCallback, useEffect, useMemo, type ReactNode } from 'react';

// Navigation entry representing a single item in the modal history
export interface NavigationEntry {
  listId: string;
  siteId: string;
  siteUrl?: string;
  itemId: string;
  listName: string;
}

interface NavigationState {
  stack: NavigationEntry[];
  currentIndex: number;
}

interface ModalNavigationContextValue {
  // Current navigation state
  currentEntry: NavigationEntry | null;
  canGoBack: boolean;
  canGoForward: boolean;
  historyLength: number;

  // Navigation actions
  navigateToItem: (entry: NavigationEntry) => void;
  goBack: () => void;
  goForward: () => void;
  reset: () => void;

  // Loading state for navigation transitions
  isNavigating: boolean;
  setIsNavigating: (value: boolean) => void;
}

const ModalNavigationContext = createContext<ModalNavigationContextValue | null>(null);

interface ModalNavigationProviderProps {
  children: ReactNode;
  modalId: string;
  initialEntry: NavigationEntry;
}

const STORAGE_KEY_PREFIX = 'listview-modal-nav-';

export function ModalNavigationProvider({ children, modalId, initialEntry }: ModalNavigationProviderProps) {
  const storageKey = `${STORAGE_KEY_PREFIX}${modalId}`;

  // Initialize state from sessionStorage or with initial entry
  const [navState, setNavState] = useState<NavigationState>(() => {
    try {
      const stored = sessionStorage.getItem(storageKey);
      if (stored) {
        const parsed = JSON.parse(stored) as NavigationState;
        // Validate the stored state
        if (parsed.stack && Array.isArray(parsed.stack) && parsed.stack.length > 0) {
          return parsed;
        }
      }
    } catch {
      // Invalid stored state, start fresh
    }
    return {
      stack: [initialEntry],
      currentIndex: 0,
    };
  });

  const [isNavigating, setIsNavigating] = useState(false);

  // Persist to sessionStorage on change
  useEffect(() => {
    try {
      sessionStorage.setItem(storageKey, JSON.stringify(navState));
    } catch {
      // Storage full or unavailable
    }
  }, [navState, storageKey]);

  // Clean up sessionStorage when component unmounts
  useEffect(() => {
    return () => {
      try {
        sessionStorage.removeItem(storageKey);
      } catch {
        // Ignore cleanup errors
      }
    };
  }, [storageKey]);

  const currentEntry = useMemo(() => {
    return navState.stack[navState.currentIndex] ?? null;
  }, [navState]);

  const canGoBack = navState.currentIndex > 0;
  const canGoForward = navState.currentIndex < navState.stack.length - 1;

  const navigateToItem = useCallback((entry: NavigationEntry) => {
    setNavState(prev => {
      // Remove any forward history (like browser behavior)
      const newStack = prev.stack.slice(0, prev.currentIndex + 1);
      newStack.push(entry);
      return {
        stack: newStack,
        currentIndex: newStack.length - 1,
      };
    });
  }, []);

  const goBack = useCallback(() => {
    setNavState(prev => {
      if (prev.currentIndex <= 0) return prev;
      return {
        ...prev,
        currentIndex: prev.currentIndex - 1,
      };
    });
  }, []);

  const goForward = useCallback(() => {
    setNavState(prev => {
      if (prev.currentIndex >= prev.stack.length - 1) return prev;
      return {
        ...prev,
        currentIndex: prev.currentIndex + 1,
      };
    });
  }, []);

  const reset = useCallback(() => {
    setNavState({
      stack: [initialEntry],
      currentIndex: 0,
    });
  }, [initialEntry]);

  const value: ModalNavigationContextValue = {
    currentEntry,
    canGoBack,
    canGoForward,
    historyLength: navState.stack.length,
    navigateToItem,
    goBack,
    goForward,
    reset,
    isNavigating,
    setIsNavigating,
  };

  return (
    <ModalNavigationContext.Provider value={value}>
      {children}
    </ModalNavigationContext.Provider>
  );
}

export function useModalNavigation(): ModalNavigationContextValue {
  const context = useContext(ModalNavigationContext);
  if (!context) {
    throw new Error('useModalNavigation must be used within a ModalNavigationProvider');
  }
  return context;
}
