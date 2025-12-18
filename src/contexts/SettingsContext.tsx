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
  DEFAULT_SETTINGS_SITE_PATH,
  createSPClient,
  buildSiteUrl,
  getSite,
  findSettingsList,
  createSettingsList,
  getSettings,
  setSetting,
  findViewsList,
  createViewsList,
  getViews,
  createView,
  updateView,
  deleteView,
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
import type { ViewDefinition } from '../types/view';
import type { PageDefinition } from '../types/page';

export interface EnabledList {
  siteId: string;
  siteName: string;
  siteUrl: string;
  listId: string;
  listName: string;
}

const ENABLED_LISTS_KEY = 'EnabledLists';

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
  viewsList: SharePointList | null;
  views: ViewDefinition[];
  pagesList: SharePointList | null;
  pages: PageDefinition[];
}

interface SettingsContextValue extends SettingsState {
  spClient: SPFI | null;
  initialize: () => Promise<void>;
  configureSite: (sitePath: string, isCustom: boolean) => Promise<boolean>;
  createList: () => Promise<boolean>;
  clearSiteOverride: () => void;
  getSetting: (key: string) => string | undefined;
  updateSetting: (key: string, value: string) => Promise<void>;
  enabledLists: EnabledList[];
  // Views operations
  createViewsList: () => Promise<boolean>;
  loadViews: () => Promise<void>;
  saveView: (view: ViewDefinition) => Promise<ViewDefinition>;
  removeView: (id: string) => Promise<void>;
  // Pages operations
  createPagesList: () => Promise<boolean>;
  loadPages: () => Promise<void>;
  savePage: (page: PageDefinition) => Promise<PageDefinition>;
  removePage: (id: string) => Promise<void>;
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
    viewsList: null,
    views: [],
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
          viewsList: null,
          views: [],
          pagesList: null,
          pages: [],
        });
        return;
      }

      // Load existing settings
      const settings = await getSettings(sp);

      // Check for views list and load views
      const viewsList = await findViewsList(sp, siteUrl);
      let views: ViewDefinition[] = [];
      if (viewsList) {
        views = await getViews(sp);
      }

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
        isCustomSite,
        site,
        settingsList,
        settings,
        viewsList,
        views,
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
            viewsList: null,
            views: [],
            pagesList: null,
            pages: [],
          });
          return true; // Site found, but list needs to be created
        }

        const settings = await getSettings(sp);

        // Check for views list and load views
        const viewsList = await findViewsList(sp, siteUrl);
        let views: ViewDefinition[] = [];
        if (viewsList) {
          views = await getViews(sp);
        }

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
          isCustomSite: isCustom,
          site,
          settingsList,
          settings,
          viewsList,
          views,
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

  // Views operations
  const createViewsListCallback = useCallback(async (): Promise<boolean> => {
    if (!state.site || !state.hostname || !state.sitePath) {
      return false;
    }

    const sp = spClientRef.current;
    if (!sp) {
      return false;
    }

    try {
      const siteUrl = buildSiteUrl(state.hostname, state.sitePath);
      const viewsList = await createViewsList(sp, siteUrl);

      setState((prev) => ({
        ...prev,
        viewsList,
        views: [],
      }));

      return true;
    } catch (error) {
      console.error('Failed to create views list:', error);
      return false;
    }
  }, [state.site, state.hostname, state.sitePath]);

  const loadViewsCallback = useCallback(async (): Promise<void> => {
    const sp = spClientRef.current;
    if (!sp || !state.viewsList) {
      return;
    }

    try {
      const views = await getViews(sp);
      setState((prev) => ({ ...prev, views }));
    } catch (error) {
      console.error('Failed to load views:', error);
    }
  }, [state.viewsList]);

  const saveViewCallback = useCallback(
    async (view: ViewDefinition): Promise<ViewDefinition> => {
      const sp = spClientRef.current;
      if (!sp) {
        throw new Error('SharePoint client not initialized');
      }

      if (!state.hostname || !state.sitePath) {
        throw new Error('Site not configured');
      }

      const siteUrl = buildSiteUrl(state.hostname, state.sitePath);

      // If no views list exists, create it first
      if (!state.viewsList) {
        console.log('[Views] Creating LV-Views list at', siteUrl);
        try {
          const viewsList = await createViewsList(sp, siteUrl);
          setState((prev) => ({ ...prev, viewsList }));
          console.log('[Views] LV-Views list created successfully');
        } catch (error) {
          console.error('[Views] Failed to create LV-Views list:', error);
          throw new Error('Failed to create views list. Please check your permissions.');
        }
      }

      if (view.id) {
        // Update existing view
        await updateView(sp, view.id, view);
        setState((prev) => ({
          ...prev,
          views: prev.views.map((v) => (v.id === view.id ? { ...v, ...view } : v)),
        }));
        return view;
      } else {
        // Create new view
        const newView = await createView(sp, view);
        setState((prev) => ({
          ...prev,
          views: [...prev.views, newView],
        }));
        return newView;
      }
    },
    [state.viewsList, state.hostname, state.sitePath]
  );

  const removeViewCallback = useCallback(
    async (id: string): Promise<void> => {
      const sp = spClientRef.current;
      if (!sp) {
        throw new Error('SharePoint client not initialized');
      }

      if (!state.viewsList) {
        // No views list, just remove from local state
        setState((prev) => ({
          ...prev,
          views: prev.views.filter((v) => v.id !== id),
        }));
        return;
      }

      await deleteView(sp, id);

      setState((prev) => ({
        ...prev,
        views: prev.views.filter((v) => v.id !== id),
      }));
    },
    [state.viewsList]
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

  // Parse enabled lists from settings
  const enabledLists = useMemo((): EnabledList[] => {
    const json = state.settings[ENABLED_LISTS_KEY];
    if (!json) return [];
    try {
      const parsed = JSON.parse(json);
      if (Array.isArray(parsed) && parsed.length > 0 && typeof parsed[0] === 'object') {
        return parsed as EnabledList[];
      }
      return [];
    } catch {
      return [];
    }
  }, [state.settings]);

  const contextValue: SettingsContextValue = {
    ...state,
    spClient: spClientRef.current,
    initialize,
    configureSite,
    createList,
    clearSiteOverride,
    getSetting: getSettingValue,
    updateSetting,
    enabledLists,
    createViewsList: createViewsListCallback,
    loadViews: loadViewsCallback,
    saveView: saveViewCallback,
    removeView: removeViewCallback,
    createPagesList: createPagesListCallback,
    loadPages: loadPagesCallback,
    savePage: savePageCallback,
    removePage: removePageCallback,
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
