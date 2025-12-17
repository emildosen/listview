import { Client } from '@microsoft/microsoft-graph-client';
import type { IPublicClientApplication, AccountInfo } from '@azure/msal-browser';
import { graphScopes, getSharePointScopes } from './msalConfig';

interface SharePointRootSiteResponse {
  siteCollection?: {
    hostname?: string;
  };
}

export interface GraphSite {
  id: string;
  displayName: string;
  name: string;
  webUrl: string;
}

export interface GraphList {
  id: string;
  displayName: string;
  name: string;
  webUrl: string;
  list?: {
    hidden: boolean;
    template: string;
  };
}

export interface GraphListColumn {
  id: string;
  name: string;
  displayName: string;
  columnGroup?: string;
  hidden?: boolean;
  readOnly?: boolean;
}

export interface GraphListItem {
  id: string;
  fields: Record<string, unknown>;
}

export interface ListItemsResult {
  columns: GraphListColumn[];
  items: GraphListItem[];
}

/**
 * Get the SharePoint hostname for the tenant using Graph API
 * Calls GET /sites/root to get the root SharePoint site and extracts the hostname
 */
export async function getSharePointHostname(
  msalInstance: IPublicClientApplication,
  account: AccountInfo
): Promise<string> {
  const client = createGraphClient(msalInstance, account);

  const response: SharePointRootSiteResponse = await client
    .api('/sites/root')
    .select('siteCollection')
    .get();

  const hostname = response.siteCollection?.hostname;
  if (!hostname) {
    throw new Error('Unable to determine SharePoint hostname from tenant');
  }

  return hostname;
}

/**
 * Create Microsoft Graph client
 */
export function createGraphClient(
  msalInstance: IPublicClientApplication,
  account: AccountInfo
): Client {
  return Client.init({
    authProvider: async (done) => {
      try {
        const response = await msalInstance.acquireTokenSilent({
          scopes: graphScopes,
          account,
        });
        done(null, response.accessToken);
      } catch {
        try {
          const response = await msalInstance.acquireTokenPopup({
            scopes: graphScopes,
            account,
          });
          done(null, response.accessToken);
        } catch (popupError) {
          done(popupError as Error, null);
        }
      }
    },
  });
}

/**
 * Get a token for SharePoint REST API (different audience than Graph)
 */
export async function getSharePointToken(
  msalInstance: IPublicClientApplication,
  account: AccountInfo,
  hostname: string
): Promise<string> {
  const scopes = getSharePointScopes(hostname);

  try {
    const response = await msalInstance.acquireTokenSilent({
      scopes,
      account,
    });
    return response.accessToken;
  } catch {
    const response = await msalInstance.acquireTokenPopup({
      scopes,
      account,
    });
    return response.accessToken;
  }
}

/**
 * Get all SharePoint sites accessible to the current user
 */
export async function getAllSites(
  msalInstance: IPublicClientApplication,
  account: AccountInfo
): Promise<GraphSite[]> {
  const client = createGraphClient(msalInstance, account);
  const sites: GraphSite[] = [];

  let response = await client
    .api('/sites?search=*')
    .select('id,displayName,name,webUrl')
    .top(100)
    .get();

  sites.push(...(response.value || []));

  // Handle pagination
  while (response['@odata.nextLink']) {
    response = await client.api(response['@odata.nextLink']).get();
    sites.push(...(response.value || []));
  }

  return sites;
}

/**
 * Get all lists for a specific site
 */
export async function getSiteLists(
  msalInstance: IPublicClientApplication,
  account: AccountInfo,
  siteId: string
): Promise<GraphList[]> {
  const client = createGraphClient(msalInstance, account);

  const response = await client
    .api(`/sites/${siteId}/lists`)
    .select('id,displayName,name,webUrl,list')
    .get();

  // Filter to non-hidden generic lists (template 100)
  const lists: GraphList[] = (response.value || []).filter(
    (list: GraphList) => !list.list?.hidden && list.list?.template === 'genericList'
  );

  return lists;
}

/**
 * Get columns for a specific list
 */
export async function getListColumns(
  msalInstance: IPublicClientApplication,
  account: AccountInfo,
  siteId: string,
  listId: string
): Promise<GraphListColumn[]> {
  const client = createGraphClient(msalInstance, account);

  const response = await client
    .api(`/sites/${siteId}/lists/${listId}/columns`)
    .select('id,name,displayName,columnGroup,hidden,readOnly')
    .get();

  // Filter out hidden and system columns
  const columns: GraphListColumn[] = (response.value || []).filter(
    (col: GraphListColumn) =>
      !col.hidden &&
      col.columnGroup !== '_Hidden' &&
      !['ContentType', 'Attachments', '_UIVersionString', 'Edit', 'LinkTitleNoMenu', 'LinkTitle', 'DocIcon', 'ItemChildCount', 'FolderChildCount', 'AppAuthor', 'AppEditor'].includes(col.name)
  );

  return columns;
}

/**
 * Get list info by ID
 */
export async function getListById(
  msalInstance: IPublicClientApplication,
  account: AccountInfo,
  siteId: string,
  listId: string
): Promise<GraphList | null> {
  const client = createGraphClient(msalInstance, account);

  try {
    const list = await client
      .api(`/sites/${siteId}/lists/${listId}`)
      .select('id,displayName,name,webUrl')
      .get();

    return list;
  } catch {
    return null;
  }
}

/**
 * Get items for a specific list (top 1000)
 */
export async function getListItems(
  msalInstance: IPublicClientApplication,
  account: AccountInfo,
  siteId: string,
  listId: string
): Promise<ListItemsResult> {
  const client = createGraphClient(msalInstance, account);

  // First get columns
  const columns = await getListColumns(msalInstance, account, siteId, listId);

  // Build $expand=fields($select=...) with column names
  const columnNames = columns.map((c) => c.name).join(',');

  // Get items with pagination, up to 1000
  const items: GraphListItem[] = [];
  let response = await client
    .api(`/sites/${siteId}/lists/${listId}/items`)
    .expand(`fields($select=${columnNames})`)
    .top(250)
    .get();

  items.push(...(response.value || []));

  // Continue fetching until we have 1000 or no more pages
  while (response['@odata.nextLink'] && items.length < 1000) {
    response = await client.api(response['@odata.nextLink']).get();
    items.push(...(response.value || []));
  }

  // Limit to 1000
  return {
    columns,
    items: items.slice(0, 1000),
  };
}
