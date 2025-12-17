import { spfi, SPFI } from '@pnp/sp';
import { SPBrowser } from '@pnp/sp/behaviors/spbrowser';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/fields';
import '@pnp/sp/views';
import type { IPublicClientApplication, AccountInfo } from '@azure/msal-browser';
import { getSharePointToken } from '../auth/graphClient';
import type { Queryable } from '@pnp/queryable';

export const DEFAULT_SETTINGS_SITE_PATH = '/sites/ListView';
const SETTINGS_LIST_NAME = 'LV-Settings';

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
    const web = await sp.web.select('Id', 'Title', 'Url')();

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

/**
 * Find the settings list if it exists
 */
export async function findSettingsList(
  sp: SPFI,
  siteUrl: string
): Promise<SharePointList | null> {
  try {
    const list = await sp.web.lists.getByTitle(SETTINGS_LIST_NAME)
      .select('Id', 'Title', 'RootFolder/ServerRelativeUrl')
      .expand('RootFolder')();

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
    const items = await sp.web.lists
      .getByTitle(SETTINGS_LIST_NAME)
      .items
      .select('SettingKey', 'SettingValue')();

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
