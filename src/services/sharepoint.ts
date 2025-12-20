import { spfi, SPFI } from '@pnp/sp';
import { SPBrowser } from '@pnp/sp/behaviors/spbrowser';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/fields';
import '@pnp/sp/views';
import '@pnp/sp/sites';
import type { IPublicClientApplication, AccountInfo } from '@azure/msal-browser';
import { getSharePointToken } from '../auth/graphClient';
import type { Queryable } from '@pnp/queryable';
import type { PageDefinition, PageItem } from '../types/page';

// Global callback for auth errors (set by AuthErrorContext)
let onAuthError: ((message: string) => void) | null = null;

/**
 * Register a callback to be called when authentication errors occur.
 * This is set by the AuthErrorContext to show the session expired modal.
 */
export function registerAuthErrorCallback(callback: (message: string) => void): void {
  onAuthError = callback;
}

/**
 * Unregister the auth error callback
 */
export function unregisterAuthErrorCallback(): void {
  onAuthError = null;
}

/**
 * Check if an error indicates an expired or invalid token
 */
function isTokenExpiredError(error: unknown): boolean {
  if (!error) return false;

  const errorStr = String(error);
  const errorMessage = error instanceof Error ? error.message : '';

  // Check error response body for JSON error
  if (typeof error === 'object' && error !== null) {
    const err = error as Record<string, unknown>;
    // Check for error_description in response
    if (typeof err.error_description === 'string') {
      if (
        err.error_description.includes('Invalid JWT') ||
        err.error_description.includes('token is expired') ||
        err.error_description.includes('Token has expired')
      ) {
        return true;
      }
    }
    // Check for nested response body
    if (err.response && typeof err.response === 'object') {
      const response = err.response as Record<string, unknown>;
      if (typeof response.error_description === 'string') {
        if (response.error_description.includes('Invalid JWT') ||
            response.error_description.includes('token is expired')) {
          return true;
        }
      }
    }
  }

  const expiredPatterns = [
    'Invalid JWT token',
    'token is expired',
    'Token has expired',
    'access_token is expired',
    'AADSTS700024',
    'AADSTS50173',
    'AADSTS500133',
  ];

  for (const pattern of expiredPatterns) {
    if (errorStr.includes(pattern) || errorMessage.includes(pattern)) {
      return true;
    }
  }

  return false;
}

/**
 * Wrap an async function to catch token errors and show the modal
 */
async function withTokenErrorHandling<T>(fn: () => Promise<T>): Promise<T> {
  try {
    return await fn();
  } catch (error) {
    if (isTokenExpiredError(error)) {
      console.error('[SharePoint] Token expired error detected:', error);
      if (onAuthError) {
        onAuthError('Your SharePoint session has expired. Please reload the app to sign in again.');
      }
    }
    throw error;
  }
}

export const DEFAULT_SETTINGS_SITE_PATH = '/sites/ListView';
const SETTINGS_LIST_NAME = 'LV-Settings';
const PAGES_LIST_NAME = 'LV-Pages';

/**
 * System list names used by ListView app - these should be hidden from users
 * Add any new system lists here
 */
export const SYSTEM_LIST_NAMES = [
  SETTINGS_LIST_NAME,
  PAGES_LIST_NAME,
] as const;

export interface SharePointSite {
  id: string;
  name: string;
  displayName: string;
  webUrl: string;
  hostname: string;
}

export interface SharePointList {
  id: string;
  name: string;
  displayName: string;
  webUrl: string;
}

export interface SettingItem {
  Id?: number;
  SettingKey: string;
  SettingValue: string;
}

/**
 * Create a bearer token auth behavior for PnPjs
 */
function MSAL(token: string) {
  return (instance: Queryable) => {
    instance.on.auth.replace(async (url: URL, init: RequestInit) => {
      init.headers = {
        ...init.headers,
        Authorization: `Bearer ${token}`,
      };
      return [url, init];
    });
    return instance;
  };
}

/**
 * Create a PnPjs SP instance for a specific site
 */
export async function createSPClient(
  msalInstance: IPublicClientApplication,
  account: AccountInfo,
  siteUrl: string
): Promise<SPFI> {
  const url = new URL(siteUrl);
  const hostname = url.hostname;
  const token = await getSharePointToken(msalInstance, account, hostname);

  return spfi(siteUrl).using(SPBrowser(), MSAL(token));
}

/**
 * Extract hostname from a SharePoint URL
 */
export function getHostnameFromUrl(url: string): string {
  return new URL(url).hostname;
}

/**
 * Build site URL from hostname and path
 */
export function buildSiteUrl(hostname: string, sitePath: string): string {
  return `https://${hostname}${sitePath}`;
}

/**
 * Check if a SharePoint site exists and is accessible
 */
export async function getSite(
  msalInstance: IPublicClientApplication,
  account: AccountInfo,
  hostname: string,
  sitePath: string
): Promise<SharePointSite | null> {
  const siteUrl = buildSiteUrl(hostname, sitePath);

  try {
    const sp = await createSPClient(msalInstance, account, siteUrl);
    const web = await withTokenErrorHandling(() => sp.web.select('Id', 'Title', 'Url')());

    return {
      id: web.Id,
      name: sitePath.split('/').pop() || '',
      displayName: web.Title,
      webUrl: web.Url,
      hostname,
    };
  } catch (error: unknown) {
    console.error('[SharePoint] Failed to get site:', error);
    // Check if it's a 404 or access denied
    if (error && typeof error === 'object' && 'status' in error) {
      const status = (error as { status: number }).status;
      if (status === 404 || status === 403) {
        return null;
      }
    }
    throw error;
  }
}

export interface CreateSiteOptions {
  title: string;
  url: string;
  description?: string;
  lcid?: number; // Language code, defaults to 1033 (English)
}

/**
 * Create a new Communication Site using PnP
 * Uses the SPSiteManager/create endpoint which is available to any user
 * with site creation permissions in the tenant
 */
export async function createCommunicationSite(
  msalInstance: IPublicClientApplication,
  account: AccountInfo,
  hostname: string,
  options: CreateSiteOptions
): Promise<SharePointSite> {
  // We need to connect to the root site to use the SPSiteManager endpoint
  const rootUrl = `https://${hostname}`;
  const sp = await createSPClient(msalInstance, account, rootUrl);

  const fullUrl = `https://${hostname}${options.url}`;

  console.log('[SharePoint] Creating communication site:', fullUrl);

  // Create the communication site using PnP
  // Note: siteDesignId should be a GUID or omitted for default template
  const result = await sp.site.createCommunicationSite(
    options.title,
    options.lcid || 1033,      // Language (English)
    false,                      // ShareByEmailEnabled
    fullUrl,
    options.description || '',
    '',                         // Classification (empty = none)
    undefined                   // SiteDesignId - undefined uses default template
  );

  console.log('[SharePoint] Site creation result:', result);

  // Wait a moment for site provisioning to complete
  await new Promise(resolve => setTimeout(resolve, 2000));

  // Verify the site was created and return its details
  const site = await getSite(msalInstance, account, hostname, options.url);

  if (!site) {
    throw new Error('Site was created but could not be accessed. Please try again in a moment.');
  }

  return site;
}

/**
 * Find the settings list if it exists
 */
export async function findSettingsList(
  sp: SPFI,
  siteUrl: string
): Promise<SharePointList | null> {
  try {
    const list = await withTokenErrorHandling(() =>
      sp.web.lists.getByTitle(SETTINGS_LIST_NAME)
        .select('Id', 'Title', 'RootFolder/ServerRelativeUrl')
        .expand('RootFolder')()
    );

    return {
      id: list.Id,
      name: SETTINGS_LIST_NAME,
      displayName: list.Title,
      webUrl: `${siteUrl}/Lists/${SETTINGS_LIST_NAME}`,
    };
  } catch (error: unknown) {
    // List not found
    if (error && typeof error === 'object' && 'status' in error) {
      const status = (error as { status: number }).status;
      if (status === 404) {
        return null;
      }
    }
    throw error;
  }
}

/**
 * Create the settings list with custom columns
 */
export async function createSettingsList(
  sp: SPFI,
  siteUrl: string
): Promise<SharePointList> {
  // Create the list
  const listAddResult = await sp.web.lists.add(
    SETTINGS_LIST_NAME,
    'ListView application settings',
    100, // Generic list template
    false // Don't enable content types
  );

  const list = sp.web.lists.getByTitle(SETTINGS_LIST_NAME);

  // Add SettingKey column (single line text, indexed)
  await list.fields.addText('SettingKey', {
    MaxLength: 255,
    Indexed: true,
  });

  // Add SettingValue column (multi-line text)
  await list.fields.addMultilineText('SettingValue', {
    NumberOfLines: 6,
    RichText: false,
  });

  // Add columns to default view
  const views = await list.views.filter("DefaultView eq true")();
  if (views.length > 0) {
    const defaultViewId = views[0].Id;
    await list.views.getById(defaultViewId).fields.add('SettingKey');
    await list.views.getById(defaultViewId).fields.add('SettingValue');
  }

  return {
    id: listAddResult.Id,
    name: SETTINGS_LIST_NAME,
    displayName: SETTINGS_LIST_NAME,
    webUrl: `${siteUrl}/Lists/${SETTINGS_LIST_NAME}`,
  };
}

/**
 * Get all settings from the list
 */
export async function getSettings(
  sp: SPFI
): Promise<Record<string, string>> {
  try {
    const items = await withTokenErrorHandling(() =>
      sp.web.lists
        .getByTitle(SETTINGS_LIST_NAME)
        .items
        .select('SettingKey', 'SettingValue')()
    );

    const settings: Record<string, string> = {};
    for (const item of items) {
      if (item.SettingKey) {
        settings[item.SettingKey] = item.SettingValue || '';
      }
    }
    return settings;
  } catch {
    return {};
  }
}

/**
 * Set a setting value
 */
export async function setSetting(
  sp: SPFI,
  key: string,
  value: string
): Promise<void> {
  const list = sp.web.lists.getByTitle(SETTINGS_LIST_NAME);

  // Check if setting exists
  const items = await list.items
    .filter(`SettingKey eq '${key}'`)
    .select('Id')();

  if (items.length > 0) {
    // Update existing
    await list.items.getById(items[0].Id).update({
      SettingValue: value,
    });
  } else {
    // Create new
    await list.items.add({
      SettingKey: key,
      SettingValue: value,
    });
  }
}

/**
 * Delete a setting
 */
export async function deleteSetting(
  sp: SPFI,
  key: string
): Promise<void> {
  const list = sp.web.lists.getByTitle(SETTINGS_LIST_NAME);

  const items = await list.items
    .filter(`SettingKey eq '${key}'`)
    .select('Id')();

  if (items.length > 0) {
    await list.items.getById(items[0].Id).delete();
  }
}

/**
 * Get the root SharePoint site URL to determine hostname
 */
export async function getRootSiteUrl(
  msalInstance: IPublicClientApplication,
  account: AccountInfo,
  tenantName: string
): Promise<string> {
  // Try common SharePoint hostname patterns
  const hostname = `${tenantName}.sharepoint.com`;
  const rootUrl = `https://${hostname}`;

  try {
    const sp = await createSPClient(msalInstance, account, rootUrl);
    const web = await sp.web.select('Url')();
    return web.Url;
  } catch {
    throw new Error('Unable to connect to SharePoint. Check your tenant name.');
  }
}

// ============================================
// List View Column Order Operations
// ============================================

/**
 * Get the default view's column order for a list
 * Returns an array of internal column names in the order they appear in the default view
 */
export async function getDefaultViewColumnOrder(
  msalInstance: IPublicClientApplication,
  account: AccountInfo,
  listWebUrl: string
): Promise<string[]> {
  try {
    // Extract the site URL and list name from the list web URL
    // listWebUrl is like: https://tenant.sharepoint.com/sites/SiteName/Lists/ListName
    const url = new URL(listWebUrl);
    const pathParts = url.pathname.split('/');
    const listsIndex = pathParts.findIndex(p => p.toLowerCase() === 'lists');

    if (listsIndex === -1 || listsIndex >= pathParts.length - 1) {
      console.warn('[SharePoint] Could not parse list URL:', listWebUrl);
      return [];
    }

    const listName = decodeURIComponent(pathParts[listsIndex + 1]);
    const sitePath = pathParts.slice(0, listsIndex).join('/');
    const siteUrl = `${url.origin}${sitePath}`;

    const sp = await createSPClient(msalInstance, account, siteUrl);

    // Get the default view and its fields
    const views = await sp.web.lists.getByTitle(listName).views.filter("DefaultView eq true")();

    if (views.length === 0) {
      return [];
    }

    const defaultViewId = views[0].Id;
    const viewFields = await sp.web.lists.getByTitle(listName).views.getById(defaultViewId).fields();

    // viewFields.Items contains the internal names of columns in order
    return viewFields.Items || [];
  } catch (error) {
    console.warn('[SharePoint] Could not get default view column order:', error);
    return [];
  }
}

// ============================================
// Pages List Operations
// ============================================

/**
 * Find the pages list if it exists
 */
export async function findPagesList(
  sp: SPFI,
  siteUrl: string
): Promise<SharePointList | null> {
  try {
    const list = await sp.web.lists.getByTitle(PAGES_LIST_NAME)
      .select('Id', 'Title', 'RootFolder/ServerRelativeUrl')
      .expand('RootFolder')();

    return {
      id: list.Id,
      name: PAGES_LIST_NAME,
      displayName: list.Title,
      webUrl: `${siteUrl}/Lists/${PAGES_LIST_NAME}`,
    };
  } catch (error: unknown) {
    if (error && typeof error === 'object' && 'status' in error) {
      const status = (error as { status: number }).status;
      if (status === 404) {
        return null;
      }
    }
    throw error;
  }
}

/**
 * Create the pages list with custom columns
 */
export async function createPagesList(
  sp: SPFI,
  siteUrl: string
): Promise<SharePointList> {
  const listAddResult = await sp.web.lists.add(
    PAGES_LIST_NAME,
    'ListView custom pages configuration',
    100,
    false
  );

  const list = sp.web.lists.getByTitle(PAGES_LIST_NAME);

  // Add PageConfig column (multi-line text for JSON storage)
  await list.fields.addMultilineText('PageConfig', {
    NumberOfLines: 10,
    RichText: false,
  });

  // Add columns to default view
  const views = await list.views.filter("DefaultView eq true")();
  if (views.length > 0) {
    const defaultViewId = views[0].Id;
    await list.views.getById(defaultViewId).fields.add('PageConfig');
  }

  return {
    id: listAddResult.Id,
    name: PAGES_LIST_NAME,
    displayName: PAGES_LIST_NAME,
    webUrl: `${siteUrl}/Lists/${PAGES_LIST_NAME}`,
  };
}

/**
 * Get all pages from the list
 */
export async function getPages(sp: SPFI): Promise<PageDefinition[]> {
  try {
    const items: PageItem[] = await withTokenErrorHandling(() =>
      sp.web.lists
        .getByTitle(PAGES_LIST_NAME)
        .items
        .select('Id', 'Title', 'PageConfig')()
    );

    return items.map((item) => {
      try {
        const config = JSON.parse(item.PageConfig || '{}') as Partial<Omit<PageDefinition, 'id' | 'name'>>;
        return {
          id: String(item.Id),
          name: item.Title,
          // Default pageType for backwards compatibility with existing pages
          pageType: config.pageType || 'lookup',
          primarySource: config.primarySource || { siteId: '', listId: '', listName: '' },
          displayColumns: config.displayColumns || [],
          searchConfig: config.searchConfig || {
            tableColumns: [],
            textSearchColumns: [],
            filterColumns: [],
          },
          relatedSections: config.relatedSections || [],
          detailLayout: config.detailLayout,
          reportLayout: config.reportLayout,
          description: config.description,
          createdAt: config.createdAt,
          updatedAt: config.updatedAt,
        };
      } catch {
        return {
          id: String(item.Id),
          name: item.Title,
          pageType: 'lookup' as const,
          primarySource: { siteId: '', listId: '', listName: '' },
          displayColumns: [],
          searchConfig: {
            tableColumns: [],
            textSearchColumns: [],
            filterColumns: [],
          },
          relatedSections: [],
        };
      }
    });
  } catch (error: unknown) {
    if (error && typeof error === 'object' && 'status' in error) {
      const status = (error as { status: number }).status;
      if (status === 404) {
        console.log('[Pages] LV-Pages list does not exist yet');
        return [];
      }
    }
    console.error('[Pages] Error fetching pages:', error);
    return [];
  }
}

/**
 * Get a single page by ID
 */
export async function getPage(sp: SPFI, id: string): Promise<PageDefinition | null> {
  try {
    const item: PageItem = await sp.web.lists
      .getByTitle(PAGES_LIST_NAME)
      .items
      .getById(parseInt(id, 10))
      .select('Id', 'Title', 'PageConfig')();

    const config = JSON.parse(item.PageConfig || '{}') as Omit<PageDefinition, 'id' | 'name'>;
    return {
      ...config,
      id: String(item.Id),
      name: item.Title,
    };
  } catch (error: unknown) {
    if (error && typeof error === 'object' && 'status' in error) {
      const status = (error as { status: number }).status;
      if (status === 404) {
        return null;
      }
    }
    console.error('[Pages] Error fetching page:', error);
    return null;
  }
}

/**
 * Create a new page
 */
export async function createPage(sp: SPFI, page: PageDefinition): Promise<PageDefinition> {
  const { name, ...config } = page;

  const result = await sp.web.lists
    .getByTitle(PAGES_LIST_NAME)
    .items
    .add({
      Title: name,
      PageConfig: JSON.stringify({
        ...config,
        createdAt: new Date().toISOString(),
        updatedAt: new Date().toISOString(),
      }),
    });

  return {
    ...page,
    id: String(result.Id),
    createdAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
  };
}

/**
 * Update an existing page
 */
export async function updatePage(
  sp: SPFI,
  id: string,
  page: Partial<PageDefinition>
): Promise<void> {
  const { name, ...config } = page;

  const updateData: Record<string, unknown> = {
    PageConfig: JSON.stringify({
      ...config,
      updatedAt: new Date().toISOString(),
    }),
  };

  if (name) {
    updateData.Title = name;
  }

  await sp.web.lists
    .getByTitle(PAGES_LIST_NAME)
    .items
    .getById(parseInt(id, 10))
    .update(updateData);
}

/**
 * Delete a page
 */
export async function deletePage(sp: SPFI, id: string): Promise<void> {
  await sp.web.lists
    .getByTitle(PAGES_LIST_NAME)
    .items
    .getById(parseInt(id, 10))
    .delete();
}

// ============================================
// Column Formatting Operations
// ============================================

export interface ColumnFormatting {
  internalName: string;
  customFormatter: string | null;
}

/**
 * Get custom column formatting for all fields in a list
 */
export async function getColumnFormatting(
  sp: SPFI,
  listId: string
): Promise<ColumnFormatting[]> {
  try {
    const fields = await sp.web.lists
      .getById(listId)
      .fields
      .select('InternalName', 'CustomFormatter')();

    return fields.map((field: { InternalName: string; CustomFormatter?: string }) => ({
      internalName: field.InternalName,
      customFormatter: field.CustomFormatter || null,
    }));
  } catch (error) {
    console.error('[SharePoint] Failed to get column formatting:', error);
    return [];
  }
}

/**
 * Parse column formatting JSON to determine if it renders as a link
 * Returns the URL field expression if it's a link formatter, null otherwise
 */
export function parseColumnFormattingForLink(customFormatter: string | null): boolean {
  if (!customFormatter) return false;

  try {
    const format = JSON.parse(customFormatter);

    // Check if root element is an anchor tag
    if (format.elmType === 'a') {
      return true;
    }

    // Check for nested anchor in children
    if (format.children && Array.isArray(format.children)) {
      return format.children.some((child: { elmType?: string }) => child.elmType === 'a');
    }

    return false;
  } catch {
    return false;
  }
}

/**
 * Get column order from the default view of a list
 * Returns array of internal column names in display order
 */
export async function getListColumnOrder(
  sp: SPFI,
  listId: string
): Promise<string[]> {
  try {
    const views = await sp.web.lists.getById(listId).views.filter("DefaultView eq true")();

    if (views.length === 0) {
      return [];
    }

    const defaultViewId = views[0].Id;
    const viewFields = await sp.web.lists.getById(listId).views.getById(defaultViewId).fields();

    return viewFields.Items || [];
  } catch (error) {
    console.warn('[SharePoint] Could not get column order:', error);
    return [];
  }
}

// ============================================
// Generic List Item CRUD Operations
// ============================================

/**
 * Create a new item in any list
 */
export async function createListItem(
  sp: SPFI,
  listId: string,
  fields: Record<string, unknown>
): Promise<{ Id: number }> {
  const result = await sp.web.lists
    .getById(listId)
    .items
    .add(fields);

  return { Id: result.Id };
}

/**
 * Update an item in any list
 */
export async function updateListItem(
  sp: SPFI,
  listId: string,
  itemId: number,
  fields: Record<string, unknown>
): Promise<void> {
  await sp.web.lists
    .getById(listId)
    .items
    .getById(itemId)
    .update(fields);
}

/**
 * Delete an item from any list
 */
export async function deleteListItem(
  sp: SPFI,
  listId: string,
  itemId: number
): Promise<void> {
  await sp.web.lists
    .getById(listId)
    .items
    .getById(itemId)
    .delete();
}
