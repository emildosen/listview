import { useState, useEffect } from 'react';
import { useMsal } from '@azure/msal-react';
import { Link } from '@fluentui/react-components';
import {
  DocumentRegular,
  DocumentTextRegular,
  FolderRegular,
  LinkRegular,
  DocumentTableRegular,
} from '@fluentui/react-icons';
import {
  isSharePointUrl,
  parseSharePointUrl,
  resolveSharePointUrl,
  type SharePointUrlInfo,
  type SharePointResourceType,
} from '../../auth/graphClient';

interface SharePointLinkProps {
  url: string;
  inline?: boolean;
  stopPropagation?: boolean;
}

// Map resource types to icons
function getIconForType(type: SharePointResourceType) {
  switch (type) {
    case 'file':
      return <DocumentRegular style={{ fontSize: '12px' }} />;
    case 'page':
      return <DocumentTextRegular style={{ fontSize: '12px' }} />;
    case 'folder':
      return <FolderRegular style={{ fontSize: '12px' }} />;
    case 'list-item':
      return <DocumentTableRegular style={{ fontSize: '12px' }} />;
    case 'generic':
    default:
      return <LinkRegular style={{ fontSize: '12px' }} />;
  }
}

// Get the href for the link, adding web=1 for files to open in browser
function getHref(url: string, type: SharePointResourceType): string {
  if (type === 'file') {
    // Add web=1 parameter to open files in browser instead of downloading
    try {
      const urlObj = new URL(url.trim());
      urlObj.searchParams.set('web', '1');
      return urlObj.toString();
    } catch {
      return url;
    }
  }
  return url;
}

/**
 * SharePointLink - Renders a SharePoint URL as a clickable link with:
 * - Appropriate icon based on resource type
 * - Display name parsed from URL or resolved via Graph API
 * - Async resolution with immediate display of parsed name
 */
export function SharePointLink({ url, inline = true, stopPropagation = true }: SharePointLinkProps) {
  const { instance, accounts } = useMsal();
  const account = accounts[0];

  // Parse URL immediately for instant display
  const [info, setInfo] = useState<SharePointUrlInfo>(() => parseSharePointUrl(url));

  useEffect(() => {
    // Attempt async resolution if we have auth
    if (account && isSharePointUrl(url)) {
      resolveSharePointUrl(instance, account, url)
        .then((resolved) => {
          // Only update if the resolved name is different
          if (resolved.displayName !== info.displayName) {
            setInfo(resolved);
          }
        })
        .catch(() => {
          // Keep the parsed result on error
        });
    }
  }, [url, instance, account, info.displayName]);

  const handleClick = (e: React.MouseEvent) => {
    if (stopPropagation) {
      e.stopPropagation();
    }
  };

  return (
    <Link
      href={getHref(url, info.type)}
      target="_blank"
      rel="noopener noreferrer"
      inline={inline}
      onClick={handleClick}
      style={{ display: 'inline-flex', alignItems: 'center', gap: '4px' }}
    >
      {getIconForType(info.type)}
      {info.displayName}
    </Link>
  );
}

