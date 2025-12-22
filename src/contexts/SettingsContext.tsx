/* eslint-disable react-refresh/only-export-components */
import {
  createContext,
  useContext,
  useState,
  useCallback,
  useEffect,
  useRef,
  useMemo,
  type ReactNode,
} from 'react';
import { useMsal } from '@azure/msal-react';
import type { SPFI } from '@pnp/sp';
import {
  createSPClient,
  buildSiteUrl,
  getSite,
  findSettingsList,
  createSettingsList,
  getSettings,
  setSetting,
  findPagesList,
  createPagesList,
  getPages,
  createPage,
  updatePage,
  deletePage,
  type SharePointSite,
  type SharePointList,
} from '../services/sharepoint';
import { getSharePointHostname } from '../auth/graphClient';
import type { PageDefinition, ListDetailConfig } from '../types/page';

const LIST_DETAIL_CONFIGS_KEY = 'ListDetailConfigs';

const LOCAL_STORAGE_KEY = 'listview-primary-site';
const HOSTNAME_STORAGE_KEY = 'listview-sharepoint-hostname';

type SetupStatus =
  | 'loading'
  | 'no-site-configured'   // Hostname detected, need to choose standard or custom site
  | 'site-not-found'       // Site URL set but site doesn't exist/not accessible
  | 'creating-site'        // Currently creating the SharePoint site
  | 'site-creation-failed' // Failed to create the site
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
  pagesList: SharePointList | null;
  pages: PageDefinition[];
}

interface SettingsContextValue extends SettingsState {
  spClient: SPFI | null;
  initialize: () => Promise<void>;
  configureSite: (sitePath: string) => Promise<boolean>;
  createList: () => Promise<boolean>;
  changePrimarySite: () => void;
  getSetting: (key: string) => string | undefined;
  updateSetting: (key: string, value: string) => Promise<void>;
  // Pages operations
  createPagesList: () => Promise<boolean>;
  loadPages: () => Promise<void>;
  savePage: (page: PageDefinition) => Promise<PageDefinition>;
  removePage: (id: string) => Promise<void>;
  // List detail config operations (per-list popup settings)
  listDetailConfigs: Record<string, ListDetailConfig>;
  getListDetailConfig: (listId: string) => ListDetailConfig | undefined;
  saveListDetailConfig: (config: ListDetailConfig) => Promise<void>;
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
    pagesList: null,
    pages: [],
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

      // Check for primary site in localStorage - user must explicitly choose
      const sitePath = localStorage.getItem(LOCAL_STORAGE_KEY);

      if (!sitePath) {
        // No primary site configured - show site picker
        setState((prev) => ({
          ...prev,
          setupStatus: 'no-site-configured',
          hostname,
          sitePath: null,
          isCustomSite: false,
          site: null,
          settingsList: null,
        }));
        return;
      }

      // Try to access the configured site
      const site = await getSite(instance, account, hostname, sitePath);

      if (!site) {
        setState((prev) => ({
          ...prev,
          setupStatus: 'site-not-found',
          hostname,
          sitePath,
          isCustomSite: false,
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
          isCustomSite: false,
          site,
          settingsList: null,
          settings: {},
          pagesList: null,
          pages: [],
        });
        return;
      }

      // Load existing settings
      const settings = await getSettings(sp);

      // Check for pages list and load pages
      const pagesList = await findPagesList(sp, siteUrl);
      let pages: PageDefinition[] = [];
      if (pagesList) {
        pages = await getPages(sp);
      }

      setState({
        setupStatus: 'ready',
        error: null,
        hostname,
        sitePath,
        isCustomSite: false,
        site,
        settingsList,
        settings,
        pagesList,
        pages,
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
    async (sitePath: string): Promise<boolean> => {
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
            isCustomSite: false,
            site: null,
          }));
          return false;
        }

        // Save primary site to localStorage
        localStorage.setItem(LOCAL_STORAGE_KEY, sitePath);

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
            isCustomSite: false,
            site,
            settingsList: null,
            settings: {},
            pagesList: null,
            pages: [],
          });
          return true; // Site found, but list needs to be created
        }

        const settings = await getSettings(sp);

        // Check for pages list and load pages
        const pagesList = await findPagesList(sp, siteUrl);
        let pages: PageDefinition[] = [];
        if (pagesList) {
          pages = await getPages(sp);
        }

        setState({
          setupStatus: 'ready',
          error: null,
          hostname,
          sitePath,
          isCustomSite: false,
          site,
          settingsList,
          settings,
          pagesList,
          pages,
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

      // Create all system lists
      const settingsList = await createSettingsList(sp, siteUrl);
      const settings = await getSettings(sp);
      const pagesList = await createPagesList(sp, siteUrl);

      setState((prev) => ({
        ...prev,
        setupStatus: 'ready',
        settingsList,
        settings,
        pagesList,
        pages: [],
      }));

      return true;
    } catch (error) {
      console.error('Failed to create system lists:', error);
      setState((prev) => ({
        ...prev,
        setupStatus: 'list-creation-failed',
        error: error instanceof Error ? error.message : 'Failed to create system lists',
      }));
      return false;
    }
  }, [instance, accounts, state.site, state.hostname, state.sitePath]);

  const changePrimarySite = useCallback(() => {
    localStorage.removeItem(LOCAL_STORAGE_KEY);
    spClientRef.current = null;
    setState((prev) => ({
      ...prev,
      setupStatus: 'no-site-configured',
      isCustomSite: false,
      sitePath: null,
      site: null,
      settingsList: null,
      settings: {},
      pagesList: null,
      pages: [],
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

  // Pages operations
  const createPagesListCallback = useCallback(async (): Promise<boolean> => {
    if (!state.site || !state.hostname || !state.sitePath) {
      return false;
    }

    const sp = spClientRef.current;
    if (!sp) {
      return false;
    }

    try {
      const siteUrl = buildSiteUrl(state.hostname, state.sitePath);
      const pagesList = await createPagesList(sp, siteUrl);

      setState((prev) => ({
        ...prev,
        pagesList,
        pages: [],
      }));

      return true;
    } catch (error) {
      console.error('Failed to create pages list:', error);
      return false;
    }
  }, [state.site, state.hostname, state.sitePath]);

  const loadPagesCallback = useCallback(async (): Promise<void> => {
    const sp = spClientRef.current;
    if (!sp || !state.pagesList) {
      return;
    }

    try {
      const pages = await getPages(sp);
      setState((prev) => ({ ...prev, pages }));
    } catch (error) {
      console.error('Failed to load pages:', error);
    }
  }, [state.pagesList]);

  const savePageCallback = useCallback(
    async (page: PageDefinition): Promise<PageDefinition> => {
      const sp = spClientRef.current;
      if (!sp) {
        throw new Error('SharePoint client not initialized');
      }

      if (!state.hostname || !state.sitePath) {
        throw new Error('Site not configured');
      }

      const siteUrl = buildSiteUrl(state.hostname, state.sitePath);

      // If no pages list exists, create it first
      if (!state.pagesList) {
        console.log('[Pages] Creating LV-Pages list at', siteUrl);
        try {
          const pagesList = await createPagesList(sp, siteUrl);
          setState((prev) => ({ ...prev, pagesList }));
          console.log('[Pages] LV-Pages list created successfully');
        } catch (error) {
          console.error('[Pages] Failed to create LV-Pages list:', error);
          throw new Error('Failed to create pages list. Please check your permissions.');
        }
      }

      if (page.id) {
        // Update existing page
        await updatePage(sp, page.id, page);
        setState((prev) => ({
          ...prev,
          pages: prev.pages.map((p) => (p.id === page.id ? { ...p, ...page } : p)),
        }));
        return page;
      } else {
        // Create new page
        const newPage = await createPage(sp, page);
        setState((prev) => ({
          ...prev,
          pages: [...prev.pages, newPage],
        }));
        return newPage;
      }
    },
    [state.pagesList, state.hostname, state.sitePath]
  );

  const removePageCallback = useCallback(
    async (id: string): Promise<void> => {
      const sp = spClientRef.current;
      if (!sp) {
        throw new Error('SharePoint client not initialized');
      }

      if (!state.pagesList) {
        // No pages list, just remove from local state
        setState((prev) => ({
          ...prev,
          pages: prev.pages.filter((p) => p.id !== id),
        }));
        return;
      }

      await deletePage(sp, id);

      setState((prev) => ({
        ...prev,
        pages: prev.pages.filter((p) => p.id !== id),
      }));
    },
    [state.pagesList]
  );

  // Auto-initialize when accounts change
  useEffect(() => {
    if (accounts.length > 0) {
      initialize();
    }
  }, [accounts.length, initialize]);

  // Parse list detail configs from settings (per-list popup configurations)
  const listDetailConfigs = useMemo((): Record<string, ListDetailConfig> => {
    const json = state.settings[LIST_DETAIL_CONFIGS_KEY];
    if (!json) return {};
    try {
      const parsed = JSON.parse(json);
      if (typeof parsed === 'object' && parsed !== null) {
        return parsed as Record<string, ListDetailConfig>;
      }
      return {};
    } catch {
      return {};
    }
  }, [state.settings]);

  // Get list detail config by listId
  const getListDetailConfig = useCallback(
    (listId: string): ListDetailConfig | undefined => {
      return listDetailConfigs[listId];
    },
    [listDetailConfigs]
  );

  // Save list detail config
  const saveListDetailConfig = useCallback(
    async (config: ListDetailConfig): Promise<void> => {
      const sp = spClientRef.current;
      if (!sp) {
        throw new Error('SharePoint client not initialized');
      }

      // Get current configs and update
      const currentConfigs = { ...listDetailConfigs };
      currentConfigs[config.listId] = config;

      // Save to settings
      await setSetting(sp, LIST_DETAIL_CONFIGS_KEY, JSON.stringify(currentConfigs));

      // Update local state
      setState((prev) => ({
        ...prev,
        settings: {
          ...prev.settings,
          [LIST_DETAIL_CONFIGS_KEY]: JSON.stringify(currentConfigs),
        },
      }));
    },
    [listDetailConfigs]
  );

  const contextValue: SettingsContextValue = {
    ...state,
    spClient: spClientRef.current,
    initialize,
    configureSite,
    createList,
    changePrimarySite,
    getSetting: getSettingValue,
    updateSetting,
    createPagesList: createPagesListCallback,
    loadPages: loadPagesCallback,
    savePage: savePageCallback,
    removePage: removePageCallback,
    listDetailConfigs,
    getListDetailConfig,
    saveListDetailConfig,
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
