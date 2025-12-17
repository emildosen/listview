/* eslint-disable react-refresh/only-export-components */
import {
  createContext,
  useContext,
  useState,
  useCallback,
  useEffect,
  useRef,
  type ReactNode,
} from 'react';
import { useMsal } from '@azure/msal-react';
import type { SPFI } from '@pnp/sp';
import {
  DEFAULT_SETTINGS_SITE_PATH,
  createSPClient,
  buildSiteUrl,
  getSite,
  findSettingsList,
  createSettingsList,
  getSettings,
  setSetting,
  type SharePointSite,
  type SharePointList,
} from '../services/sharepoint';
import { getSharePointHostname } from '../auth/graphClient';

const LOCAL_STORAGE_KEY = 'listview-settings-site-override';
const HOSTNAME_STORAGE_KEY = 'listview-sharepoint-hostname';

type SetupStatus =
  | 'loading'
  | 'no-site-configured'   // Hostname detected, need to choose standard or custom site
  | 'site-not-found'       // Site URL set but site doesn't exist/not accessible
  | 'list-not-found'       // Site exists but settings list not found, need to create
  | 'creating-list'        // Currently creating the settings list
  | 'list-creation-failed' // Failed to create the list
  | 'ready'                // All good, settings loaded
  | 'error';               // General error

interface SettingsState {
  setupStatus: SetupStatus;
  error: string | null;
  hostname: string | null;
  sitePath: string | null;
  isCustomSite: boolean;
  site: SharePointSite | null;
  settingsList: SharePointList | null;
  settings: Record<string, string>;
}

interface SettingsContextValue extends SettingsState {
  spClient: SPFI | null;
  initialize: () => Promise<void>;
  configureSite: (sitePath: string, isCustom: boolean) => Promise<boolean>;
  createList: () => Promise<boolean>;
  clearSiteOverride: () => void;
  getSetting: (key: string) => string | undefined;
  updateSetting: (key: string, value: string) => Promise<void>;
}

const SettingsContext = createContext<SettingsContextValue | null>(null);

export function SettingsProvider({ children }: { children: ReactNode }) {
  const { instance, accounts } = useMsal();
  const spClientRef = useRef<SPFI | null>(null);
  const [state, setState] = useState<SettingsState>({
    setupStatus: 'loading',
    error: null,
    hostname: localStorage.getItem(HOSTNAME_STORAGE_KEY),
    sitePath: null,
    isCustomSite: false,
    site: null,
    settingsList: null,
    settings: {},
  });

  const getAccount = useCallback(() => {
    const account = accounts[0];
    if (!account) {
      throw new Error('No authenticated account');
    }
    return account;
  }, [accounts]);

  const initialize = useCallback(async () => {
    setState((prev) => ({ ...prev, setupStatus: 'loading', error: null }));

    try {
      const account = getAccount();

      // Auto-detect SharePoint hostname via Graph API
      // First check localStorage cache, then fetch from Graph if needed
      let hostname = localStorage.getItem(HOSTNAME_STORAGE_KEY);

      if (!hostname) {
        hostname = await getSharePointHostname(instance, account);
        localStorage.setItem(HOSTNAME_STORAGE_KEY, hostname);
      }

      // Check for custom site override in localStorage
      const override = localStorage.getItem(LOCAL_STORAGE_KEY);
      const sitePath = override || DEFAULT_SETTINGS_SITE_PATH;
      const isCustomSite = !!override;

      // Try to access the site
      const site = await getSite(instance, account, hostname, sitePath);

      if (!site) {
        setState((prev) => ({
          ...prev,
          setupStatus: isCustomSite ? 'site-not-found' : 'no-site-configured',
          hostname,
          sitePath,
          isCustomSite,
          site: null,
          settingsList: null,
        }));
        return;
      }

      // Save hostname if we derived it
      localStorage.setItem(HOSTNAME_STORAGE_KEY, hostname);

      // Create SP client for this site
      const siteUrl = buildSiteUrl(hostname, sitePath);
      const sp = await createSPClient(instance, account, siteUrl);
      spClientRef.current = sp;

      // Check for settings list
      const settingsList = await findSettingsList(sp, siteUrl);

      if (!settingsList) {
        setState({
          setupStatus: 'list-not-found',
          error: null,
          hostname,
          sitePath,
          isCustomSite,
          site,
          settingsList: null,
          settings: {},
        });
        return;
      }

      // Load existing settings
      const settings = await getSettings(sp);

      setState({
        setupStatus: 'ready',
        error: null,
        hostname,
        sitePath,
        isCustomSite,
        site,
        settingsList,
        settings,
      });
    } catch (error) {
      console.error('Failed to initialize settings:', error);
      setState((prev) => ({
        ...prev,
        setupStatus: 'error',
        error: error instanceof Error ? error.message : 'Unknown error',
      }));
    }
  }, [instance, getAccount]);

  const configureSite = useCallback(
    async (sitePath: string, isCustom: boolean): Promise<boolean> => {
      try {
        const account = getAccount();
        const hostname = state.hostname;

        if (!hostname) {
          return false;
        }

        const site = await getSite(instance, account, hostname, sitePath);

        if (!site) {
          setState((prev) => ({
            ...prev,
            setupStatus: 'site-not-found',
            hostname,
            sitePath,
            isCustomSite: isCustom,
            site: null,
          }));
          return false;
        }

        // Save override if custom
        if (isCustom) {
          localStorage.setItem(LOCAL_STORAGE_KEY, sitePath);
        } else {
          localStorage.removeItem(LOCAL_STORAGE_KEY);
        }

        // Create SP client for this site
        const siteUrl = buildSiteUrl(hostname, sitePath);
        const sp = await createSPClient(instance, account, siteUrl);
        spClientRef.current = sp;

        // Check for existing settings list
        const settingsList = await findSettingsList(sp, siteUrl);

        if (!settingsList) {
          setState({
            setupStatus: 'list-not-found',
            error: null,
            hostname,
            sitePath,
            isCustomSite: isCustom,
            site,
            settingsList: null,
            settings: {},
          });
          return true; // Site found, but list needs to be created
        }

        const settings = await getSettings(sp);

        setState({
          setupStatus: 'ready',
          error: null,
          hostname,
          sitePath,
          isCustomSite: isCustom,
          site,
          settingsList,
          settings,
        });

        return true;
      } catch (error) {
        console.error('Failed to configure site:', error);
        setState((prev) => ({
          ...prev,
          setupStatus: 'error',
          error: error instanceof Error ? error.message : 'Unknown error',
        }));
        return false;
      }
    },
    [instance, getAccount, state.hostname]
  );

  const createList = useCallback(async (): Promise<boolean> => {
    if (!state.site || !state.hostname || !state.sitePath) {
      return false;
    }

    const account = accounts[0];
    if (!account) {
      return false;
    }

    setState((prev) => ({ ...prev, setupStatus: 'creating-list', error: null }));

    try {
      const siteUrl = buildSiteUrl(state.hostname, state.sitePath);

      // Ensure we have a SP client
      let sp = spClientRef.current;
      if (!sp) {
        sp = await createSPClient(instance, account, siteUrl);
        spClientRef.current = sp;
      }

      const settingsList = await createSettingsList(sp, siteUrl);
      const settings = await getSettings(sp);

      setState((prev) => ({
        ...prev,
        setupStatus: 'ready',
        settingsList,
        settings,
      }));

      return true;
    } catch (error) {
      console.error('Failed to create settings list:', error);
      setState((prev) => ({
        ...prev,
        setupStatus: 'list-creation-failed',
        error: error instanceof Error ? error.message : 'Failed to create settings list',
      }));
      return false;
    }
  }, [instance, accounts, state.site, state.hostname, state.sitePath]);

  const clearSiteOverride = useCallback(() => {
    localStorage.removeItem(LOCAL_STORAGE_KEY);
    spClientRef.current = null;
    setState((prev) => ({
      ...prev,
      setupStatus: 'loading',
      isCustomSite: false,
      sitePath: DEFAULT_SETTINGS_SITE_PATH,
    }));
  }, []);

  const getSettingValue = useCallback(
    (key: string): string | undefined => {
      return state.settings[key];
    },
    [state.settings]
  );

  const updateSetting = useCallback(
    async (key: string, value: string): Promise<void> => {
      const sp = spClientRef.current;
      if (!sp) {
        throw new Error('SharePoint client not initialized');
      }

      await setSetting(sp, key, value);

      setState((prev) => ({
        ...prev,
        settings: {
          ...prev.settings,
          [key]: value,
        },
      }));
    },
    []
  );

  // Auto-initialize when accounts change
  useEffect(() => {
    if (accounts.length > 0) {
      initialize();
    }
  }, [accounts.length, initialize]);

  const contextValue: SettingsContextValue = {
    ...state,
    spClient: spClientRef.current,
    initialize,
    configureSite,
    createList,
    clearSiteOverride,
    getSetting: getSettingValue,
    updateSetting,
  };

  return (
    <SettingsContext.Provider value={contextValue}>
      {children}
    </SettingsContext.Provider>
  );
}

export function useSettings(): SettingsContextValue {
  const context = useContext(SettingsContext);
  if (!context) {
    throw new Error('useSettings must be used within a SettingsProvider');
  }
  return context;
}
