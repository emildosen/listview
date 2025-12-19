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
  // Default value configured in SharePoint
  defaultValue?: { value?: string; formula?: string };
  // Type-specific properties from Graph API (only one will be present)
  text?: { allowMultipleLines?: boolean; maxLength?: number };
  boolean?: Record<string, never>;  // Empty object indicates boolean column
  number?: { minimum?: number; maximum?: number };
  dateTime?: { format?: string };
  // Lookup column metadata - present if this is a lookup column
  lookup?: {
    listId: string;       // The list this lookup points to
    columnName: string;   // The column in the target list
    allowMultipleValues?: boolean;
  };
  // Choice column metadata - present if this is a choice column
  choice?: {
    choices: string[];
    allowMultipleValues?: boolean;
  };
  // Hyperlink or picture column - present if this is a URL column
  hyperlinkOrPicture?: {
    isPicture?: boolean;
  };
}

// Form field configuration from ContentType columnPositions (ordered for forms)
export interface FormFieldConfig {
  id: string;
  name: string;
  displayName: string;
  required: boolean;
  hidden: boolean;
  readOnly: boolean;
  text?: { allowMultipleLines?: boolean; maxLength?: number };
  boolean?: Record<string, never>;
  number?: { minimum?: number; maximum?: number };
  dateTime?: { format?: string };
  lookup?: { listId: string; columnName: string; allowMultipleValues?: boolean };
  choice?: { choices: string[]; allowMultipleValues?: boolean };
  hyperlinkOrPicture?: { isPicture?: boolean };
  defaultValue?: { value?: string; formula?: string };
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
    .select('id,name,displayName,columnGroup,hidden,readOnly,defaultValue,text,boolean,number,dateTime,lookup,choice,hyperlinkOrPicture')
    .get();

  // Filter out system columns but keep hidden columns (marked with hidden: true for UI to filter)
  const columns: GraphListColumn[] = (response.value || [])
    .filter(
      (col: GraphListColumn & { lookup?: { listId?: string; columnName?: string; allowMultipleValues?: boolean } }) =>
        col.columnGroup !== '_Hidden' &&
        !['ContentType', 'Attachments', '_UIVersionString', 'Edit', 'LinkTitleNoMenu', 'LinkTitle', 'DocIcon', 'ItemChildCount', 'FolderChildCount', 'AppAuthor', 'AppEditor'].includes(col.name)
    )
    .map((col: GraphListColumn) => {
      // Include type metadata from Graph API
      const result: GraphListColumn = {
        id: col.id,
        name: col.name,
        displayName: col.displayName,
        columnGroup: col.columnGroup,
        hidden: col.hidden,
        readOnly: col.readOnly,
      };

      // Default value
      if (col.defaultValue) result.defaultValue = col.defaultValue;

      // Type-specific properties (only one will be present)
      if (col.text) result.text = col.text;
      if (col.boolean) result.boolean = col.boolean;
      if (col.number) result.number = col.number;
      if (col.dateTime) result.dateTime = col.dateTime;

      if (col.lookup?.listId) {
        result.lookup = {
          listId: col.lookup.listId,
          columnName: col.lookup.columnName || 'Title',
          allowMultipleValues: col.lookup.allowMultipleValues,
        };
      }

      if (col.choice?.choices) {
        result.choice = {
          choices: col.choice.choices,
          allowMultipleValues: col.choice.allowMultipleValues,
        };
      }

      if (col.hyperlinkOrPicture) {
        result.hyperlinkOrPicture = {
          isPicture: col.hyperlinkOrPicture.isPicture,
        };
      }

      return result;
    });

  return columns;
}

/**
 * Get form field configuration from ContentType columns.
 * Returns fields in the order they appear in SharePoint's default edit/new forms.
 */
export async function getFormFieldConfig(
  msalInstance: IPublicClientApplication,
  account: AccountInfo,
  siteId: string,
  listId: string
): Promise<{ contentTypeId: string; fields: FormFieldConfig[] }> {
  const client = createGraphClient(msalInstance, account);

  // Step 1: Get content types for this list
  const contentTypesResponse = await client
    .api(`/sites/${siteId}/lists/${listId}/contentTypes`)
    .select('id,name,hidden')
    .top(10)
    .get();

  // Find the default content type (usually "Item", first non-hidden one)
  const contentTypes = contentTypesResponse.value || [];
  const defaultContentType = contentTypes.find(
    (ct: { hidden?: boolean; name?: string }) =>
      !ct.hidden && ct.name !== 'Folder' && !ct.name?.startsWith('_')
  ) || contentTypes[0];

  if (!defaultContentType) {
    throw new Error('No content type found for list');
  }

  const contentTypeId = defaultContentType.id;
  // URL-encode the contentTypeId since it can contain special characters like 0x0100...
  const encodedContentTypeId = encodeURIComponent(contentTypeId);

  // Step 2: Get columns for this content type (they come in form order)
  const columnsResponse = await client
    .api(`/sites/${siteId}/lists/${listId}/contentTypes/${encodedContentTypeId}/columns`)
    .select('id,name,displayName,hidden,readOnly,required,defaultValue,text,boolean,number,dateTime,lookup,choice,hyperlinkOrPicture')
    .get();

  const columns = columnsResponse.value || [];

  // Map to FormFieldConfig, preserving order from columnPositions
  const fields: FormFieldConfig[] = columns.map(
    (col: GraphListColumn & { required?: boolean }) => {
      const field: FormFieldConfig = {
        id: col.id,
        name: col.name,
        displayName: col.displayName,
        required: col.required ?? false,
        hidden: col.hidden ?? false,
        readOnly: col.readOnly ?? false,
      };

      if (col.defaultValue) field.defaultValue = col.defaultValue;
      if (col.text) field.text = col.text;
      if (col.boolean) field.boolean = col.boolean;
      if (col.number) field.number = col.number;
      if (col.dateTime) field.dateTime = col.dateTime;

      if (col.lookup?.listId) {
        field.lookup = {
          listId: col.lookup.listId,
          columnName: col.lookup.columnName || 'Title',
          allowMultipleValues: col.lookup.allowMultipleValues,
        };
      }

      if (col.choice?.choices) {
        field.choice = {
          choices: col.choice.choices,
          allowMultipleValues: col.choice.allowMultipleValues,
        };
      }

      if (col.hyperlinkOrPicture) {
        field.hyperlinkOrPicture = {
          isPicture: col.hyperlinkOrPicture.isPicture,
        };
      }

      return field;
    }
  );

  return { contentTypeId, fields };
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
  // For lookup columns, also request the LookupId field (e.g., StudentLookupId for Student column)
  const columnNames: string[] = [];
  for (const col of columns) {
    columnNames.push(col.name);
    if (col.lookup) {
      // SharePoint stores lookup IDs in a separate field: {ColumnName}LookupId
      columnNames.push(`${col.name}LookupId`);
    }
  }

  // Get items with pagination, up to 1000
  const items: GraphListItem[] = [];
  let response = await client
    .api(`/sites/${siteId}/lists/${listId}/items`)
    .expand(`fields($select=${columnNames.join(',')})`)
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

// SharePoint URL resolution types and utilities

export type SharePointResourceType = 'file' | 'page' | 'list-item' | 'list-attachment' | 'folder' | 'generic';

export interface SharePointUrlInfo {
  displayName: string;
  type: SharePointResourceType;
  resolved: boolean; // true if resolved via Graph API, false if parsed from URL
}

/**
 * Check if a string value starts with a SharePoint URL
 */
export function isSharePointUrl(value: string): boolean {
  return typeof value === 'string' &&
    /^https:\/\/[^/]+\.sharepoint\.com/i.test(value.trim());
}

/**
 * Parse a SharePoint URL to extract display info without making API calls
 */
export function parseSharePointUrl(url: string): SharePointUrlInfo {
  try {
    const urlObj = new URL(url.trim());
    const path = decodeURIComponent(urlObj.pathname);
    const searchParams = urlObj.searchParams;

    // Check for list item display form: /Lists/{ListName}/DispForm.aspx?ID=X
    const listItemMatch = path.match(/\/Lists\/([^/]+)\/(?:DispForm|EditForm|NewForm)\.aspx/i);
    if (listItemMatch) {
      const listName = listItemMatch[1].replace(/%20/g, ' ');
      const itemId = searchParams.get('ID');
      return {
        displayName: itemId ? `${listName} #${itemId}` : listName,
        type: 'list-item',
        resolved: false,
      };
    }

    // Check for SitePages: /sites/{site}/SitePages/{PageName}.aspx
    const pageMatch = path.match(/\/SitePages\/([^/]+)\.aspx$/i);
    if (pageMatch) {
      // Convert hyphens to spaces and clean up
      const pageName = pageMatch[1]
        .replace(/-/g, ' ')
        .replace(/%20/g, ' ');
      return {
        displayName: pageName,
        type: 'page',
        resolved: false,
      };
    }

    // Check for list attachments: /Lists/{ListName}/Attachments/{ItemId}/{Filename}
    // These are NOT stored in drives, so we handle them separately
    const listAttachmentMatch = path.match(/\/Lists\/[^/]+\/Attachments\/\d+\/([^/]+\.[a-zA-Z0-9]{2,5})$/i);
    if (listAttachmentMatch) {
      return {
        displayName: listAttachmentMatch[1],
        type: 'list-attachment',
        resolved: true, // Mark as resolved since we can't fetch more info via drive API
      };
    }

    // Check for files in document libraries (has file extension)
    const fileExtMatch = path.match(/\/([^/]+\.[a-zA-Z0-9]{2,5})$/);
    if (fileExtMatch) {
      return {
        displayName: fileExtMatch[1],
        type: 'file',
        resolved: false,
      };
    }

    // Check for folders in Shared Documents or other libraries
    const folderMatch = path.match(/\/(Shared%20Documents|Documents|[^/]+)\/([^/]+)\/?$/);
    if (folderMatch && !folderMatch[2].includes('.')) {
      return {
        displayName: folderMatch[2].replace(/%20/g, ' '),
        type: 'folder',
        resolved: false,
      };
    }

    // Generic: use last path segment
    const segments = path.split('/').filter(Boolean);
    const lastSegment = segments[segments.length - 1] || url;
    return {
      displayName: lastSegment.replace(/%20/g, ' '),
      type: 'generic',
      resolved: false,
    };
  } catch {
    // If URL parsing fails, return the original URL
    return {
      displayName: url,
      type: 'generic',
      resolved: false,
    };
  }
}

// Cache for resolved SharePoint URLs
const resolvedUrlCache = new Map<string, SharePointUrlInfo>();

/**
 * Resolve a SharePoint URL to get display info via Graph API
 * Falls back to URL parsing if API call fails
 */
export async function resolveSharePointUrl(
  msalInstance: IPublicClientApplication,
  account: AccountInfo,
  url: string
): Promise<SharePointUrlInfo> {
  const trimmedUrl = url.trim();

  // Check cache first
  const cached = resolvedUrlCache.get(trimmedUrl);
  if (cached) {
    return cached;
  }

  // Parse URL first to get immediate result and determine type
  const parsed = parseSharePointUrl(trimmedUrl);

  // Only attempt Graph API resolution for files
  if (parsed.type === 'file') {
    try {
      const urlObj = new URL(trimmedUrl);
      const hostname = urlObj.hostname;
      const path = decodeURIComponent(urlObj.pathname);

      // Extract site path: /sites/{siteName} or just the root
      const siteMatch = path.match(/^(\/sites\/[^/]+)/);
      const sitePath = siteMatch ? siteMatch[1] : '';

      // Extract file path after document library
      // Common patterns: /Shared Documents/, /Documents/, /sites/{site}/{library}/
      const docLibMatch = path.match(/(?:\/Shared%20Documents|\/Documents|\/sites\/[^/]+\/[^/]+)(\/.*)/);
      if (docLibMatch) {
        const filePath = docLibMatch[1];

        const client = createGraphClient(msalInstance, account);

        // Get site ID
        const siteResponse = await client
          .api(`/sites/${hostname}:${sitePath || '/'}`)
          .select('id')
          .get();

        if (siteResponse?.id) {
          // Try to get file metadata
          const driveItem = await client
            .api(`/sites/${siteResponse.id}/drive/root:${filePath}`)
            .select('name')
            .get();

          if (driveItem?.name) {
            const result: SharePointUrlInfo = {
              displayName: driveItem.name,
              type: 'file',
              resolved: true,
            };
            resolvedUrlCache.set(trimmedUrl, result);
            return result;
          }
        }
      }
    } catch {
      // API call failed, fall back to parsed result
    }
  }

  // Cache and return the parsed result
  resolvedUrlCache.set(trimmedUrl, parsed);
  return parsed;
}
