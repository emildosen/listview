import { Client } from '@microsoft/microsoft-graph-client';
import type { IPublicClientApplication, AccountInfo } from '@azure/msal-browser';
import { graphScopes, getSharePointScopes } from './msalConfig';

interface SharePointRootSiteResponse {
  siteCollection?: {
    hostname?: string;
  };
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
