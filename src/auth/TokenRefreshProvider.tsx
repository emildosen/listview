import { useEffect, useRef, useCallback, useState, type ReactNode } from 'react';
import { useIsAuthenticated, useMsal } from '@azure/msal-react';
import { InteractionRequiredAuthError } from '@azure/msal-browser';
import { graphScopes, getSharePointScopes } from './msalConfig';
import { useAuthError } from '../contexts/AuthErrorContext';

// Refresh tokens 15 minutes before they expire
// Access tokens typically expire in 1 hour, so refresh every 45 minutes
const TOKEN_REFRESH_INTERVAL_MS = 45 * 60 * 1000; // 45 minutes

// Force refresh if the tab has been hidden for longer than this
const LONG_INACTIVITY_THRESHOLD_MS = 30 * 60 * 1000; // 30 minutes

// Delay before refreshing after visibility change (avoid rapid tab switches)
const VISIBILITY_REFRESH_DELAY_MS = 2000; // 2 seconds

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
 * - Forces refresh after long inactivity (> 30 minutes)
 * - Handles both Microsoft Graph and SharePoint tokens
 * - Shows session expired modal when token refresh fails
 */
export function TokenRefreshProvider({ children }: TokenRefreshProviderProps) {
  const isAuthenticated = useIsAuthenticated();
  const { instance, accounts } = useMsal();
  const { setSessionExpired } = useAuthError();
  const intervalRef = useRef<ReturnType<typeof setInterval> | null>(null);
  const lastRefreshRef = useRef<number>(Date.now());
  const lastVisibleRef = useRef<number>(Date.now());

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
      console.warn('[TokenRefresh] Silent token refresh failed:', error);

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

          // Also refresh SharePoint token if needed
          if (currentHostname) {
            const spScopes = getSharePointScopes(currentHostname);
            await instance.acquireTokenPopup({
              scopes: spScopes,
              account,
            });
          }

          console.info('[TokenRefresh] Token refreshed via popup');
        } catch (popupError) {
          console.error('[TokenRefresh] Failed to refresh token via popup:', popupError);
          // Show session expired modal
          setSessionExpired('Your SharePoint session has expired. Please reload the app to sign in again.');
        }
      } else {
        // Check if this is a token expiration error
        const errorStr = String(error);
        if (
          errorStr.includes('Invalid JWT') ||
          errorStr.includes('token is expired') ||
          errorStr.includes('Token has expired')
        ) {
          setSessionExpired('Your SharePoint session has expired. Please reload the app to sign in again.');
        } else {
          console.error('[TokenRefresh] Token refresh failed:', error);
        }
      }
    }
  }, [instance, accounts, setSessionExpired]);

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

  // Handle visibility changes - refresh when returning to tab
  useEffect(() => {
    if (!isAuthenticated) {
      return;
    }

    const handleVisibilityChange = () => {
      if (document.visibilityState === 'visible') {
        const now = Date.now();
        const timeSinceLastVisible = now - lastVisibleRef.current;
        const timeSinceLastRefresh = now - lastRefreshRef.current;

        console.debug('[TokenRefresh] Tab became visible', {
          timeSinceLastVisible: Math.round(timeSinceLastVisible / 1000) + 's',
          timeSinceLastRefresh: Math.round(timeSinceLastRefresh / 1000) + 's',
        });

        // Force refresh if we've been away for a long time
        const shouldForceRefresh = timeSinceLastVisible > LONG_INACTIVITY_THRESHOLD_MS;

        // Delay slightly to avoid refreshing during rapid tab switches
        setTimeout(() => {
          if (document.visibilityState === 'visible') {
            refreshTokens(shouldForceRefresh);
          }
        }, VISIBILITY_REFRESH_DELAY_MS);
      } else {
        // Tab is becoming hidden, record the time
        lastVisibleRef.current = Date.now();
      }
    };

    document.addEventListener('visibilitychange', handleVisibilityChange);

    return () => {
      document.removeEventListener('visibilitychange', handleVisibilityChange);
    };
  }, [isAuthenticated, refreshTokens]);

  return <>{children}</>;
}
