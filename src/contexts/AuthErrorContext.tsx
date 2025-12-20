/* eslint-disable react-refresh/only-export-components */
import { createContext, useContext, useState, useCallback, useEffect, type ReactNode } from 'react';
import {
  registerAuthErrorCallback,
  unregisterAuthErrorCallback,
} from '../services/sharepoint';

interface AuthErrorState {
  isSessionExpired: boolean;
  errorMessage: string | null;
}

interface AuthErrorContextValue extends AuthErrorState {
  setSessionExpired: (message?: string) => void;
  clearError: () => void;
}

const AuthErrorContext = createContext<AuthErrorContextValue | null>(null);

/**
 * AuthErrorProvider manages global authentication error state.
 * When a token expires or becomes invalid, components can call setSessionExpired()
 * to show the session expired modal.
 *
 * Also registers with the SharePoint service to catch token errors from API calls.
 */
export function AuthErrorProvider({ children }: { children: ReactNode }) {
  const [state, setState] = useState<AuthErrorState>({
    isSessionExpired: false,
    errorMessage: null,
  });

  const setSessionExpired = useCallback((message?: string) => {
    setState({
      isSessionExpired: true,
      errorMessage: message || 'Your session has expired. Please reload the app to continue.',
    });
  }, []);

  const clearError = useCallback(() => {
    setState({
      isSessionExpired: false,
      errorMessage: null,
    });
  }, []);

  // Register the callback with SharePoint service to catch token errors
  useEffect(() => {
    registerAuthErrorCallback(setSessionExpired);
    return () => {
      unregisterAuthErrorCallback();
    };
  }, [setSessionExpired]);

  return (
    <AuthErrorContext.Provider value={{ ...state, setSessionExpired, clearError }}>
      {children}
    </AuthErrorContext.Provider>
  );
}

export function useAuthError(): AuthErrorContextValue {
  const context = useContext(AuthErrorContext);
  if (!context) {
    throw new Error('useAuthError must be used within an AuthErrorProvider');
  }
  return context;
}

/**
 * Check if an error indicates an expired or invalid JWT token.
 * Matches common error messages from SharePoint/Graph API.
 */
export function isTokenExpiredError(error: unknown): boolean {
  if (!error) return false;

  // Check for error message patterns
  const errorStr = String(error);
  const errorMessage = error instanceof Error ? error.message : '';

  const expiredPatterns = [
    'Invalid JWT token',
    'token is expired',
    'Token has expired',
    'access_token is expired',
    'AADSTS700024', // Token expired
    'AADSTS50173', // Token expired
    'AADSTS500133', // Token not yet valid or expired
    'interaction_required',
    'login_required',
    'token_refresh_required',
  ];

  for (const pattern of expiredPatterns) {
    if (errorStr.includes(pattern) || errorMessage.includes(pattern)) {
      return true;
    }
  }

  // Check for status code 401 with token-related errors
  if (typeof error === 'object' && error !== null) {
    const err = error as { status?: number; statusCode?: number; error_description?: string };
    if ((err.status === 401 || err.statusCode === 401) && err.error_description) {
      return expiredPatterns.some((p) => err.error_description!.includes(p));
    }
  }

  return false;
}
