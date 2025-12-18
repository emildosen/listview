import { useEffect, useRef, useCallback, useState, type ReactNode } from 'react';
import { useIsAuthenticated, useMsal } from '@azure/msal-react';
import { InteractionRequiredAuthError } from '@azure/msal-browser';
import { graphScopes, getSharePointScopes } from './msalConfig';

// Refresh tokens 15 minutes before they expire
// Access tokens typically expire in 1 hour, so refresh every 45 minutes
const TOKEN_REFRESH_INTERVAL_MS = 45 * 60 * 1000; // 45 minutes

// Also refresh shortly after the app becomes visible again (after being in background)
const VISIBILITY_REFRESH_DELAY_MS = 5000; // 5 seconds after becoming visible

// LocalStorage key for SharePoint hostname (must match SettingsContext)
const HOSTNAME_STORAGE_KEY = 'listview-sharepoint-hostname';

interface TokenRefreshProviderProps {
  children: ReactNode;
}

/**
 * TokenRefreshProvider proactively refreshes authentication tokens before they expire.
 * This ensures the app continues to work even after being open for hours or days.
 *
 * Features:
 * - Refreshes tokens every 45 minutes (before 1-hour access token expiry)
 * - Refreshes when the browser tab becomes visible after being hidden
 * - Handles both Microsoft Graph and SharePoint tokens
 * - Reads SharePoint hostname from localStorage (set by SettingsContext)
 * - Gracefully handles token refresh failures
 */
export function TokenRefreshProvider({ children }: TokenRefreshProviderProps) {
  const isAuthenticated = useIsAuthenticated();
  const { instance, accounts } = useMsal();
  const intervalRef = useRef<ReturnType<typeof setInterval> | null>(null);
  const lastRefreshRef = useRef<number>(0);

  // Track SharePoint hostname from localStorage
  const [sharePointHostname, setSharePointHostname] = useState<string | null>(
    () => localStorage.getItem(HOSTNAME_STORAGE_KEY)
  );

  // Poll for hostname changes (in case it's set after initial load)
  useEffect(() => {
    const checkHostname = () => {
      const stored = localStorage.getItem(HOSTNAME_STORAGE_KEY);
      if (stored !== sharePointHostname) {
        setSharePointHostname(stored);
      }
    };

    // Check periodically in case SettingsContext sets the hostname
    const hostnameInterval = setInterval(checkHostname, 10000);

    // Also listen for storage events (cross-tab sync)
    const handleStorage = (e: StorageEvent) => {
      if (e.key === HOSTNAME_STORAGE_KEY) {
        setSharePointHostname(e.newValue);
      }
    };
    window.addEventListener('storage', handleStorage);

    return () => {
      clearInterval(hostnameInterval);
      window.removeEventListener('storage', handleStorage);
    };
  }, [sharePointHostname]);

  const refreshTokens = useCallback(async (force = false) => {
    const account = accounts[0];
    if (!account) {
      return;
    }

    // Don't refresh if we just did (unless forced)
    const now = Date.now();
    if (!force && now - lastRefreshRef.current < 60000) {
      return;
    }
    lastRefreshRef.current = now;

    // Get current hostname from localStorage (may have been updated)
    const currentHostname = localStorage.getItem(HOSTNAME_STORAGE_KEY);

    try {
      // Refresh Graph API token
      await instance.acquireTokenSilent({
        scopes: graphScopes,
        account,
        forceRefresh: force,
      });

      // Refresh SharePoint token if hostname is known
      if (currentHostname) {
        const spScopes = getSharePointScopes(currentHostname);
        await instance.acquireTokenSilent({
          scopes: spScopes,
          account,
          forceRefresh: force,
        });
      }

      console.debug('[TokenRefresh] Tokens refreshed successfully');
    } catch (error) {
      if (error instanceof InteractionRequiredAuthError) {
        // Token refresh failed and user interaction is required
        // This typically means the refresh token is expired or revoked
        console.warn('[TokenRefresh] Token refresh requires user interaction:', error.errorCode);

        // Attempt to acquire token via popup for better UX than redirect
        try {
          await instance.acquireTokenPopup({
            scopes: graphScopes,
            account,
          });
          console.info('[TokenRefresh] Token refreshed via popup');
        } catch (popupError) {
          console.error('[TokenRefresh] Failed to refresh token via popup:', popupError);
          // At this point, the user will need to sign in again when they try to use the app
          // The existing error handling in graphClient.ts will handle this
        }
      } else {
        console.error('[TokenRefresh] Token refresh failed:', error);
      }
    }
  }, [instance, accounts]);

  // Set up periodic token refresh
  useEffect(() => {
    if (!isAuthenticated || accounts.length === 0) {
      // Clear any existing interval if not authenticated
      if (intervalRef.current) {
        clearInterval(intervalRef.current);
        intervalRef.current = null;
      }
      return;
    }

    // Initial refresh on mount
    refreshTokens();

    // Set up periodic refresh
    intervalRef.current = setInterval(() => {
      refreshTokens();
    }, TOKEN_REFRESH_INTERVAL_MS);

    return () => {
      if (intervalRef.current) {
        clearInterval(intervalRef.current);
        intervalRef.current = null;
      }
    };
  }, [isAuthenticated, accounts.length, refreshTokens]);

  // Refresh tokens when the page becomes visible (user returns to the tab)
  useEffect(() => {
    if (!isAuthenticated) {
      return;
    }

    const handleVisibilityChange = () => {
      if (document.visibilityState === 'visible') {
        // Delay slightly to avoid refreshing during rapid tab switches
        setTimeout(() => {
          if (document.visibilityState === 'visible') {
            refreshTokens();
          }
        }, VISIBILITY_REFRESH_DELAY_MS);
      }
    };

    document.addEventListener('visibilitychange', handleVisibilityChange);

    return () => {
      document.removeEventListener('visibilitychange', handleVisibilityChange);
    };
  }, [isAuthenticated, refreshTokens]);

  return <>{children}</>;
}
