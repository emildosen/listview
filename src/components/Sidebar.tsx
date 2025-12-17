import { useState, useRef, useEffect } from 'react';
import { useMsal } from '@azure/msal-react';
import { useLocation, Link } from 'react-router-dom';
import { useTheme } from '../contexts/ThemeContext';
import { useSettings } from '../contexts/SettingsContext';
import { graphScopes } from '../auth/msalConfig';

function Sidebar() {
  const { instance, accounts } = useMsal();
  const { setupStatus, enabledLists, views } = useSettings();
  const { theme, setTheme } = useTheme();
  const location = useLocation();
  const [dropdownOpen, setDropdownOpen] = useState(false);
  const [listsExpanded, setListsExpanded] = useState(true);
  const [viewsExpanded, setViewsExpanded] = useState(true);
  const [profilePicture, setProfilePicture] = useState<string | null>(null);
  const dropdownRef = useRef<HTMLDivElement>(null);

  const account = accounts[0];
  const isReady = setupStatus === 'ready';

  // Close dropdown when clicking outside
  useEffect(() => {
    function handleClickOutside(event: MouseEvent) {
      if (dropdownRef.current && !dropdownRef.current.contains(event.target as Node)) {
        setDropdownOpen(false);
      }
    }
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  // Fetch profile picture
  useEffect(() => {
    async function fetchProfilePicture() {
      if (!account) return;

      try {
        const tokenResponse = await instance.acquireTokenSilent({
          scopes: graphScopes,
          account: account,
        });

        const response = await fetch('https://graph.microsoft.com/v1.0/me/photo/$value', {
          headers: {
            Authorization: `Bearer ${tokenResponse.accessToken}`,
          },
        });

        if (response.ok) {
          const blob = await response.blob();
          const url = URL.createObjectURL(blob);
          setProfilePicture(url);
        }
      } catch (error) {
        // Silently fail - will show initials instead
        console.debug('Could not fetch profile picture:', error);
      }
    }

    fetchProfilePicture();

    // Cleanup blob URL on unmount
    return () => {
      if (profilePicture) {
        URL.revokeObjectURL(profilePicture);
      }
    };
  }, [account, instance]);

  const handleSignOut = async () => {
    try {
      await instance.logoutPopup({
        postLogoutRedirectUri: window.location.origin,
        account: instance.getActiveAccount(),
      });
    } catch (error) {
      console.error('Logout failed:', error);
    }
  };

  // Get user initials for avatar
  const getInitials = () => {
    if (!account) return '?';
    const name = account.name || account.username || '';
    const parts = name.split(' ').filter(Boolean);
    if (parts.length >= 2) {
      return (parts[0][0] + parts[parts.length - 1][0]).toUpperCase();
    }
    return name.substring(0, 2).toUpperCase() || '?';
  };

  const isActive = (path: string) => {
    if (path === '/app') {
      return location.pathname === '/app';
    }
    return location.pathname.startsWith(path);
  };

  return (
    <aside className="w-64 bg-base-200 border-r border-base-300 flex flex-col h-screen fixed left-0 top-0">
      {/* Logo */}
      <div className="p-4 border-b border-base-300">
        <Link to="/app" className="text-xl font-bold tracking-tight">
          ListView
        </Link>
      </div>

      {/* Navigation */}
      <nav className="flex-1 p-4">
        <ul className="menu gap-1">
          <li>
            <Link
              to="/app"
              className={isActive('/app') ? 'active' : ''}
            >
              <svg
                xmlns="http://www.w3.org/2000/svg"
                fill="none"
                viewBox="0 0 24 24"
                strokeWidth={1.5}
                stroke="currentColor"
                className="w-5 h-5"
              >
                <path
                  strokeLinecap="round"
                  strokeLinejoin="round"
                  d="m2.25 12 8.954-8.955c.44-.439 1.152-.439 1.591 0L21.75 12M4.5 9.75v10.125c0 .621.504 1.125 1.125 1.125H9.75v-4.875c0-.621.504-1.125 1.125-1.125h2.25c.621 0 1.125.504 1.125 1.125V21h4.125c.621 0 1.125-.504 1.125-1.125V9.75M8.25 21h8.25"
                />
              </svg>
              Home
            </Link>
          </li>
        </ul>

        {/* Views Section */}
        {isReady && (
          <div className="mt-4">
            <div className="group flex items-center px-3 py-2">
              <button
                onClick={() => setViewsExpanded(!viewsExpanded)}
                className="flex-1 text-left text-sm font-medium text-base-content/70"
              >
                Views
              </button>
              <div className="flex items-center gap-1 opacity-0 group-hover:opacity-100 transition-opacity">
                <Link
                  to="/app/views/new"
                  className="text-base-content/50 hover:text-primary transition-colors"
                  title="Create View"
                  onClick={(e) => e.stopPropagation()}
                >
                  <svg
                    xmlns="http://www.w3.org/2000/svg"
                    fill="none"
                    viewBox="0 0 24 24"
                    strokeWidth={1.5}
                    stroke="currentColor"
                    className="w-4 h-4"
                  >
                    <path strokeLinecap="round" strokeLinejoin="round" d="M12 4.5v15m7.5-7.5h-15" />
                  </svg>
                </Link>
                <Link
                  to="/app/views"
                  className="text-base-content/50 hover:text-primary transition-colors"
                  title="Manage Views"
                  onClick={(e) => e.stopPropagation()}
                >
                  <svg
                    xmlns="http://www.w3.org/2000/svg"
                    fill="none"
                    viewBox="0 0 24 24"
                    strokeWidth={1.5}
                    stroke="currentColor"
                    className="w-4 h-4"
                  >
                    <path
                      strokeLinecap="round"
                      strokeLinejoin="round"
                      d="m16.862 4.487 1.687-1.688a1.875 1.875 0 1 1 2.652 2.652L10.582 16.07a4.5 4.5 0 0 1-1.897 1.13L6 18l.8-2.685a4.5 4.5 0 0 1 1.13-1.897l8.932-8.931Zm0 0L19.5 7.125"
                    />
                  </svg>
                </Link>
              </div>
            </div>
            {viewsExpanded && views.length > 0 && (
              <div className="ml-5 border-l border-base-300 pl-1">
                {views.map((view) => {
                  const viewPath = `/app/views/${view.id}`;
                  const isViewActive = location.pathname === viewPath;
                  return (
                    <Link
                      key={view.id}
                      to={viewPath}
                      className={`block px-3 py-1.5 mx-1 text-sm rounded-lg hover:bg-base-300 transition-colors ${
                        isViewActive ? 'bg-base-300 text-base-content' : 'text-base-content/70'
                      }`}
                      title={view.description || view.name}
                    >
                      <span className="truncate block">{view.name}</span>
                    </Link>
                  );
                })}
              </div>
            )}
          </div>
        )}

        {/* Lists Section */}
        {isReady && (
          <div className="mt-4">
            <div className="group flex items-center px-3 py-2">
              <button
                onClick={() => setListsExpanded(!listsExpanded)}
                className="flex-1 text-left text-sm font-medium text-base-content/70"
              >
                Lists
              </button>
              <Link
                to="/app/lists"
                className="opacity-0 group-hover:opacity-100 text-base-content/50 hover:text-primary transition-all"
                title="Manage Lists"
                onClick={(e) => e.stopPropagation()}
              >
                <svg
                  xmlns="http://www.w3.org/2000/svg"
                  fill="none"
                  viewBox="0 0 24 24"
                  strokeWidth={1.5}
                  stroke="currentColor"
                  className="w-4 h-4"
                >
                  <path
                    strokeLinecap="round"
                    strokeLinejoin="round"
                    d="m16.862 4.487 1.687-1.688a1.875 1.875 0 1 1 2.652 2.652L10.582 16.07a4.5 4.5 0 0 1-1.897 1.13L6 18l.8-2.685a4.5 4.5 0 0 1 1.13-1.897l8.932-8.931Zm0 0L19.5 7.125"
                  />
                </svg>
              </Link>
            </div>
            {listsExpanded && enabledLists.length > 0 && (
              <div className="ml-5 border-l border-base-300 pl-1">
                {enabledLists.map((list) => {
                  const listPath = `/app/lists/${encodeURIComponent(list.siteId)}/${encodeURIComponent(list.listId)}`;
                  const isListActive = location.pathname === listPath;
                  return (
                    <Link
                      key={`${list.siteId}:${list.listId}`}
                      to={listPath}
                      className={`block px-3 py-1.5 mx-1 text-sm rounded-lg hover:bg-base-300 transition-colors ${
                        isListActive ? 'bg-base-300 text-base-content' : 'text-base-content/70'
                      }`}
                      title={`${list.listName} (${list.siteName})`}
                    >
                      <span className="truncate block">{list.listName}</span>
                    </Link>
                  );
                })}
              </div>
            )}
          </div>
        )}
      </nav>

      {/* User Profile */}
      <div className="p-4 border-t border-base-300" ref={dropdownRef}>
        <div className="relative">
          <button
            onClick={() => setDropdownOpen(!dropdownOpen)}
            className="w-full flex items-center gap-3 p-2 rounded-lg hover:bg-base-300 transition-colors"
          >
            {/* Avatar with profile picture or initials */}
            {profilePicture ? (
              <img
                src={profilePicture}
                alt=""
                className="w-10 h-10 rounded-full object-cover"
              />
            ) : (
              <div className="w-10 h-10 rounded-full bg-primary text-primary-content flex items-center justify-center font-medium text-sm">
                {getInitials()}
              </div>
            )}
            <div className="flex-1 text-left min-w-0">
              <p className="font-medium text-sm truncate">
                {account?.name || 'User'}
              </p>
              <p className="text-xs text-base-content/60 truncate">
                {account?.username}
              </p>
            </div>
            <svg
              xmlns="http://www.w3.org/2000/svg"
              fill="none"
              viewBox="0 0 24 24"
              strokeWidth={1.5}
              stroke="currentColor"
              className={`w-4 h-4 transition-transform ${dropdownOpen ? 'rotate-180' : ''}`}
            >
              <path strokeLinecap="round" strokeLinejoin="round" d="m4.5 15.75 7.5-7.5 7.5 7.5" />
            </svg>
          </button>

          {/* Dropdown Menu */}
          {dropdownOpen && (
            <div className="absolute bottom-full left-0 right-0 mb-2 bg-base-100 border border-base-300 rounded-lg shadow-lg overflow-hidden">
              {/* Theme Switcher */}
              <div className="p-3">
                <p className="text-xs text-base-content/60 mb-2">Theme</p>
                <div className="flex gap-1">
                  <button
                    onClick={() => setTheme('light')}
                    className={`flex-1 btn btn-sm ${theme === 'light' ? 'btn-primary' : 'btn-ghost'}`}
                  >
                    <svg
                      xmlns="http://www.w3.org/2000/svg"
                      fill="none"
                      viewBox="0 0 24 24"
                      strokeWidth={1.5}
                      stroke="currentColor"
                      className="w-4 h-4"
                    >
                      <path
                        strokeLinecap="round"
                        strokeLinejoin="round"
                        d="M12 3v2.25m6.364.386-1.591 1.591M21 12h-2.25m-.386 6.364-1.591-1.591M12 18.75V21m-4.773-4.227-1.591 1.591M5.25 12H3m4.227-4.773L5.636 5.636M15.75 12a3.75 3.75 0 1 1-7.5 0 3.75 3.75 0 0 1 7.5 0Z"
                      />
                    </svg>
                    Light
                  </button>
                  <button
                    onClick={() => setTheme('dark')}
                    className={`flex-1 btn btn-sm ${theme === 'dark' ? 'btn-primary' : 'btn-ghost'}`}
                  >
                    <svg
                      xmlns="http://www.w3.org/2000/svg"
                      fill="none"
                      viewBox="0 0 24 24"
                      strokeWidth={1.5}
                      stroke="currentColor"
                      className="w-4 h-4"
                    >
                      <path
                        strokeLinecap="round"
                        strokeLinejoin="round"
                        d="M21.752 15.002A9.72 9.72 0 0 1 18 15.75c-5.385 0-9.75-4.365-9.75-9.75 0-1.33.266-2.597.748-3.752A9.753 9.753 0 0 0 3 11.25C3 16.635 7.365 21 12.75 21a9.753 9.753 0 0 0 9.002-5.998Z"
                      />
                    </svg>
                    Dark
                  </button>
                </div>
              </div>

              {/* Divider */}
              <div className="border-t border-base-300" />

              {/* Settings */}
              {isReady && (
                <Link
                  to="/app/settings"
                  className="p-3 hover:bg-base-200 transition-colors flex items-center gap-2"
                >
                  <svg
                    xmlns="http://www.w3.org/2000/svg"
                    fill="none"
                    viewBox="0 0 24 24"
                    strokeWidth={1.5}
                    stroke="currentColor"
                    className="w-5 h-5"
                  >
                    <path
                      strokeLinecap="round"
                      strokeLinejoin="round"
                      d="M9.594 3.94c.09-.542.56-.94 1.11-.94h2.593c.55 0 1.02.398 1.11.94l.213 1.281c.063.374.313.686.645.87.074.04.147.083.22.127.325.196.72.257 1.075.124l1.217-.456a1.125 1.125 0 0 1 1.37.49l1.296 2.247a1.125 1.125 0 0 1-.26 1.431l-1.003.827c-.293.241-.438.613-.43.992a7.723 7.723 0 0 1 0 .255c-.008.378.137.75.43.991l1.004.827c.424.35.534.955.26 1.43l-1.298 2.247a1.125 1.125 0 0 1-1.369.491l-1.217-.456c-.355-.133-.75-.072-1.076.124a6.47 6.47 0 0 1-.22.128c-.331.183-.581.495-.644.869l-.213 1.281c-.09.543-.56.94-1.11.94h-2.594c-.55 0-1.019-.398-1.11-.94l-.213-1.281c-.062-.374-.312-.686-.644-.87a6.52 6.52 0 0 1-.22-.127c-.325-.196-.72-.257-1.076-.124l-1.217.456a1.125 1.125 0 0 1-1.369-.49l-1.297-2.247a1.125 1.125 0 0 1 .26-1.431l1.004-.827c.292-.24.437-.613.43-.991a6.932 6.932 0 0 1 0-.255c.007-.38-.138-.751-.43-.992l-1.004-.827a1.125 1.125 0 0 1-.26-1.43l1.297-2.247a1.125 1.125 0 0 1 1.37-.491l1.216.456c.356.133.751.072 1.076-.124.072-.044.146-.086.22-.128.332-.183.582-.495.644-.869l.214-1.28Z"
                    />
                    <path
                      strokeLinecap="round"
                      strokeLinejoin="round"
                      d="M15 12a3 3 0 1 1-6 0 3 3 0 0 1 6 0Z"
                    />
                  </svg>
                  Settings
                </Link>
              )}

              {/* Sign Out */}
              <button
                onClick={handleSignOut}
                className="w-full p-3 text-left hover:bg-base-200 transition-colors flex items-center gap-2 text-error"
              >
                <svg
                  xmlns="http://www.w3.org/2000/svg"
                  fill="none"
                  viewBox="0 0 24 24"
                  strokeWidth={1.5}
                  stroke="currentColor"
                  className="w-5 h-5"
                >
                  <path
                    strokeLinecap="round"
                    strokeLinejoin="round"
                    d="M8.25 9V5.25A2.25 2.25 0 0 1 10.5 3h6a2.25 2.25 0 0 1 2.25 2.25v13.5A2.25 2.25 0 0 1 16.5 21h-6a2.25 2.25 0 0 1-2.25-2.25V15m-3 0-3-3m0 0 3-3m-3 3H15"
                  />
                </svg>
                Sign out
              </button>
            </div>
          )}
        </div>
      </div>
    </aside>
  );
}

export default Sidebar;
