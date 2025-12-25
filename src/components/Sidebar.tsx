import { useState, useRef, useEffect } from 'react';
import { useMsal } from '@azure/msal-react';
import { useLocation, Link, useNavigate } from 'react-router-dom';
import {
  makeStyles,
  tokens,
  Button,
  Avatar,
  Text,
  Divider,
  mergeClasses,
} from '@fluentui/react-components';
import {
  HomeRegular,
  AddRegular,
  EditRegular,
  ChevronUpRegular,
  WeatherSunnyRegular,
  WeatherMoonRegular,
  SettingsRegular,
  SignOutRegular,
} from '@fluentui/react-icons';
import { useTheme } from '../contexts/ThemeContext';
import { useSettings } from '../contexts/SettingsContext';
import { graphScopes } from '../auth/msalConfig';
import Logo from './Logo';

const useStyles = makeStyles({
  sidebar: {
    width: '256px',
    backgroundColor: tokens.colorNeutralBackground2,
    borderRight: `1px solid ${tokens.colorNeutralStroke1}`,
    display: 'flex',
    flexDirection: 'column',
    height: '100vh',
    position: 'fixed',
    left: 0,
    top: 0,
  },
  logo: {
    padding: '16px',
  },
  logoLink: {
    display: 'flex',
    alignItems: 'center',
    gap: '10px',
    fontFamily: 'Anta, sans-serif',
    fontSize: tokens.fontSizeBase600,
    fontWeight: 400,
    letterSpacing: '-0.02em',
    textDecoration: 'none',
    color: tokens.colorNeutralForeground1,
  },
  nav: {
    flex: 1,
    padding: '16px',
    overflowY: 'auto',
  },
  navDark: {
    scrollbarColor: '#333 #121212',
    '&::-webkit-scrollbar': {
      width: '8px',
    },
    '&::-webkit-scrollbar-track': {
      background: '#121212',
    },
    '&::-webkit-scrollbar-thumb': {
      background: '#333',
      borderRadius: '4px',
    },
    '&::-webkit-scrollbar-thumb:hover': {
      background: '#444',
    },
  },
  menuList: {
    listStyle: 'none',
    margin: 0,
    padding: 0,
    display: 'flex',
    flexDirection: 'column',
    gap: '4px',
  },
  menuItem: {
    display: 'flex',
    alignItems: 'center',
    gap: '10px',
    padding: '10px 12px',
    borderRadius: tokens.borderRadiusMedium,
    textDecoration: 'none',
    color: tokens.colorNeutralForeground1,
    fontSize: tokens.fontSizeBase400,
    cursor: 'pointer',
    transitionProperty: 'background-color',
    transitionDuration: tokens.durationNormal,
    ':hover': {
      backgroundColor: tokens.colorNeutralBackground3,
    },
  },
  menuItemActive: {
    backgroundColor: tokens.colorNeutralBackground3,
  },
  section: {
    marginTop: '16px',
  },
  sectionHeader: {
    display: 'flex',
    alignItems: 'center',
    padding: '10px 12px',
  },
  sectionTitle: {
    flex: 1,
    fontSize: tokens.fontSizeBase300,
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground2,
    background: 'none',
    border: 'none',
    cursor: 'pointer',
    textAlign: 'left',
    padding: 0,
  },
  sectionActions: {
    display: 'flex',
    alignItems: 'center',
    gap: '4px',
    opacity: 0,
    transitionProperty: 'opacity',
    transitionDuration: tokens.durationNormal,
  },
  sectionHeaderHover: {
    ':hover': {
      '& .section-actions': {
        opacity: 1,
      },
    },
  },
  sectionActionLink: {
    color: tokens.colorNeutralForeground3,
    display: 'flex',
    alignItems: 'center',
    transitionProperty: 'color',
    transitionDuration: tokens.durationNormal,
    ':hover': {
      color: tokens.colorBrandForeground1,
    },
  },
  sectionContent: {
    marginLeft: '20px',
    paddingLeft: '4px',
  },
  sectionLink: {
    display: 'block',
    padding: '8px 12px',
    margin: '0 4px',
    fontSize: tokens.fontSizeBase300,
    borderRadius: tokens.borderRadiusMedium,
    textDecoration: 'none',
    color: tokens.colorNeutralForeground2,
    whiteSpace: 'nowrap',
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    transitionProperty: 'background-color',
    transitionDuration: tokens.durationNormal,
    ':hover': {
      backgroundColor: tokens.colorNeutralBackground3,
    },
  },
  sectionLinkActive: {
    backgroundColor: tokens.colorNeutralBackground3,
    color: tokens.colorNeutralForeground1,
  },
  userProfile: {
    padding: '16px',
  },
  profileButton: {
    width: '100%',
    display: 'flex',
    alignItems: 'center',
    gap: '12px',
    padding: '8px',
    borderRadius: tokens.borderRadiusMedium,
    border: 'none',
    background: 'none',
    cursor: 'pointer',
    transitionProperty: 'background-color',
    transitionDuration: tokens.durationNormal,
    ':hover': {
      backgroundColor: tokens.colorNeutralBackground3,
    },
  },
  profileInfo: {
    flex: 1,
    textAlign: 'left',
    minWidth: 0,
  },
  profileName: {
    display: 'block',
    fontWeight: tokens.fontWeightMedium,
    fontSize: tokens.fontSizeBase200,
    whiteSpace: 'nowrap',
    overflow: 'hidden',
    textOverflow: 'ellipsis',
  },
  profileNameDark: {
    color: '#ffffff',
  },
  profileEmail: {
    display: 'block',
    fontSize: tokens.fontSizeBase100,
    color: tokens.colorNeutralForeground2,
    whiteSpace: 'nowrap',
    overflow: 'hidden',
    textOverflow: 'ellipsis',
  },
  chevron: {
    transitionProperty: 'transform',
    transitionDuration: tokens.durationNormal,
  },
  chevronRotated: {
    transform: 'rotate(180deg)',
  },
  dropdown: {
    position: 'absolute',
    bottom: '100%',
    left: 0,
    right: 0,
    marginBottom: '8px',
    backgroundColor: tokens.colorNeutralBackground1,
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    borderRadius: tokens.borderRadiusMedium,
    boxShadow: tokens.shadow16,
    overflow: 'hidden',
  },
  dropdownSection: {
    padding: '12px',
  },
  dropdownLabel: {
    display: 'block',
    fontSize: tokens.fontSizeBase100,
    color: tokens.colorNeutralForeground2,
    marginBottom: '8px',
  },
  themeButtons: {
    display: 'flex',
    gap: '4px',
  },
  dropdownItem: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
    padding: '10px 12px',
    width: '100%',
    border: 'none',
    background: 'none',
    cursor: 'pointer',
    textDecoration: 'none',
    color: tokens.colorNeutralForeground1,
    fontSize: tokens.fontSizeBase200,
    transitionProperty: 'background-color',
    transitionDuration: tokens.durationNormal,
    ':hover': {
      backgroundColor: tokens.colorNeutralBackground2,
    },
  },
  signOutItem: {
    color: tokens.colorPaletteRedForeground1,
  },
  relativeContainer: {
    position: 'relative',
  },
  addPageButton: {
    marginTop: '2px',
    width: '100%',
    justifyContent: 'flex-start',
    gap: '8px',
    padding: '6px 12px',
    minWidth: 'auto',
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground3,
    ':focus-visible': {
      outlineColor: tokens.colorNeutralStroke1,
      outlineWidth: '1px',
    },
  },
});

function Sidebar() {
  const styles = useStyles();
  const { instance, accounts } = useMsal();
  const { setupStatus, pages } = useSettings();
  const { theme, setTheme } = useTheme();
  const location = useLocation();
  const navigate = useNavigate();
  const [dropdownOpen, setDropdownOpen] = useState(false);
  const [pagesExpanded, setPagesExpanded] = useState(true);
  const [profilePicture, setProfilePicture] = useState<string | null>(null);
  const profilePictureUrlRef = useRef<string | null>(null);
  const dropdownRef = useRef<HTMLDivElement>(null);
  const [sectionHover, setSectionHover] = useState<string | null>(null);

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
          profilePictureUrlRef.current = url;
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
      if (profilePictureUrlRef.current) {
        URL.revokeObjectURL(profilePictureUrlRef.current);
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
    <aside className={styles.sidebar}>
      {/* Logo */}
      <div className={styles.logo}>
        <Link to="/app" className={styles.logoLink}>
          <Logo size={28} />
          ListView
        </Link>
      </div>

      {/* Navigation */}
      <nav className={mergeClasses(styles.nav, theme === 'dark' && styles.navDark)}>
        <ul className={styles.menuList}>
          <li>
            <Link
              to="/app"
              className={mergeClasses(
                styles.menuItem,
                isActive('/app') && styles.menuItemActive
              )}
            >
              <HomeRegular fontSize={22} />
              Home
            </Link>
          </li>
        </ul>
        <Button
          appearance="transparent"
          size="small"
          icon={<AddRegular />}
          className={styles.addPageButton}
          onClick={() => navigate('/app/pages/new')}
        >
          Add Page
        </Button>

        {/* Pages Section */}
        {isReady && (
          <div className={styles.section}>
            <div
              className={styles.sectionHeader}
              onMouseEnter={() => setSectionHover('pages')}
              onMouseLeave={() => setSectionHover(null)}
            >
              <button
                onClick={() => setPagesExpanded(!pagesExpanded)}
                className={styles.sectionTitle}
              >
                Pages
              </button>
              <div
                className={styles.sectionActions}
                style={{ opacity: sectionHover === 'pages' ? 1 : 0 }}
              >
                <Link
                  to="/app/pages/new"
                  className={styles.sectionActionLink}
                  title="Create Page"
                  onClick={(e) => e.stopPropagation()}
                >
                  <AddRegular fontSize={16} />
                </Link>
                <Link
                  to="/app/pages"
                  className={styles.sectionActionLink}
                  title="Manage Pages"
                  onClick={(e) => e.stopPropagation()}
                >
                  <EditRegular fontSize={16} />
                </Link>
              </div>
            </div>
            {pagesExpanded && pages.length > 0 && (
              <div className={styles.sectionContent}>
                {pages.map((page) => {
                  const pagePath = `/app/pages/${page.id}`;
                  const isPageActive = location.pathname === pagePath;
                  return (
                    <Link
                      key={page.id}
                      to={pagePath}
                      className={mergeClasses(
                        styles.sectionLink,
                        isPageActive && styles.sectionLinkActive
                      )}
                      title={page.description || page.name}
                    >
                      {page.name}
                    </Link>
                  );
                })}
              </div>
            )}
          </div>
        )}

      </nav>

      {/* User Profile */}
      <div className={styles.userProfile} ref={dropdownRef}>
        <div className={styles.relativeContainer}>
          <button
            onClick={() => setDropdownOpen(!dropdownOpen)}
            className={styles.profileButton}
          >
            <Avatar
              image={profilePicture ? { src: profilePicture } : undefined}
              name={account?.name || 'User'}
              initials={!profilePicture ? getInitials() : undefined}
              size={40}
              color="brand"
            />
            <div className={styles.profileInfo}>
              <Text className={mergeClasses(styles.profileName, theme === 'dark' && styles.profileNameDark)}>
                {account?.name || 'User'}
              </Text>
              <Text className={styles.profileEmail}>
                {account?.username}
              </Text>
            </div>
            <ChevronUpRegular
              fontSize={16}
              className={mergeClasses(
                styles.chevron,
                dropdownOpen && styles.chevronRotated
              )}
            />
          </button>

          {/* Dropdown Menu */}
          {dropdownOpen && (
            <div className={styles.dropdown}>
              {/* Theme Switcher */}
              <div className={styles.dropdownSection}>
                <Text className={styles.dropdownLabel}>Theme</Text>
                <div className={styles.themeButtons}>
                  <Button
                    appearance={theme === 'light' ? 'primary' : 'subtle'}
                    size="small"
                    icon={<WeatherSunnyRegular />}
                    onClick={() => setTheme('light')}
                    style={{ flex: 1 }}
                  >
                    Light
                  </Button>
                  <Button
                    appearance={theme === 'dark' ? 'primary' : 'subtle'}
                    size="small"
                    icon={<WeatherMoonRegular />}
                    onClick={() => setTheme('dark')}
                    style={{ flex: 1 }}
                  >
                    Dark
                  </Button>
                </div>
              </div>

              <Divider />

              {/* Settings */}
              {isReady && (
                <Link
                  to="/app/settings"
                  className={styles.dropdownItem}
                  onClick={() => setDropdownOpen(false)}
                >
                  <SettingsRegular fontSize={16} />
                  Settings
                </Link>
              )}

              {/* Sign Out */}
              <button
                onClick={handleSignOut}
                className={mergeClasses(styles.dropdownItem, styles.signOutItem)}
              >
                <SignOutRegular fontSize={16} />
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
