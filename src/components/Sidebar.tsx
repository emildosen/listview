import { useState, useRef, useEffect, useMemo } from 'react';
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
  Input,
  Menu,
  MenuTrigger,
  MenuPopover,
  MenuList,
  MenuItem,
} from '@fluentui/react-components';
import {
  HomeRegular,
  AddRegular,
  ChevronUpRegular,
  ChevronDownRegular,
  ChevronRightRegular,
  WeatherSunnyRegular,
  WeatherMoonRegular,
  SettingsRegular,
  SignOutRegular,
  ReOrderDotsVerticalRegular,
  EditRegular,
  DeleteRegular,
  MoreHorizontalRegular,
} from '@fluentui/react-icons';
import {
  DndContext,
  closestCenter,
  KeyboardSensor,
  PointerSensor,
  useSensor,
  useSensors,
  type DragEndEvent,
} from '@dnd-kit/core';
import {
  arrayMove,
  SortableContext,
  sortableKeyboardCoordinates,
  useSortable,
  verticalListSortingStrategy,
} from '@dnd-kit/sortable';
import { CSS } from '@dnd-kit/utilities';
import { useTheme } from '../contexts/ThemeContext';
import { useSettings } from '../contexts/SettingsContext';
import { graphScopes } from '../auth/msalConfig';
import { getPageIcon } from '../utils/iconMap';
import type { Section, PageDefinition } from '../types/page';
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
  },
  menuItemActive: {
    backgroundColor: tokens.colorNeutralBackground3,
  },
  menuItemHover: {
    backgroundColor: tokens.colorNeutralBackground3,
  },
  // Page link styles
  pageLink: {
    display: 'flex',
    alignItems: 'center',
    gap: '10px',
    padding: '8px 12px',
    borderRadius: tokens.borderRadiusMedium,
    textDecoration: 'none',
    color: tokens.colorNeutralForeground2,
    fontSize: tokens.fontSizeBase300,
    whiteSpace: 'nowrap',
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    transitionProperty: 'background-color',
    transitionDuration: tokens.durationNormal,
  },
  pageLinkActive: {
    backgroundColor: tokens.colorNeutralBackground3,
    color: tokens.colorNeutralForeground1,
  },
  pageLinkHover: {
    backgroundColor: tokens.colorNeutralBackground3,
  },
  pageIcon: {
    flexShrink: 0,
    color: tokens.colorNeutralForeground3,
  },
  pageName: {
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap',
  },
  // Section styles
  sectionHeader: {
    display: 'flex',
    alignItems: 'center',
    padding: '8px 8px 8px 0',
    marginTop: '12px',
    borderRadius: tokens.borderRadiusMedium,
    cursor: 'pointer',
    transitionProperty: 'background-color',
    transitionDuration: tokens.durationNormal,
  },
  sectionHeaderHover: {
    backgroundColor: tokens.colorNeutralBackground3,
  },
  sectionDragHandle: {
    cursor: 'grab',
    color: tokens.colorNeutralForeground3,
    marginRight: '2px',
    opacity: 0,
    transitionProperty: 'opacity',
    transitionDuration: tokens.durationNormal,
  },
  sectionDragHandleVisible: {
    opacity: 1,
  },
  sectionChevron: {
    color: tokens.colorNeutralForeground3,
    marginRight: '6px',
  },
  sectionTitle: {
    flex: 1,
    fontSize: tokens.fontSizeBase300,
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground2,
    textTransform: 'uppercase',
    letterSpacing: '0.02em',
  },
  sectionTitleInput: {
    flex: 1,
  },
  sectionActions: {
    opacity: 0,
    transitionProperty: 'opacity',
    transitionDuration: tokens.durationNormal,
  },
  sectionActionsVisible: {
    opacity: 1,
  },
  sectionContent: {
    marginLeft: '4px',
    paddingLeft: '8px',
    borderLeft: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  sectionDragging: {
    opacity: 0.5,
  },
  // User profile
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
  },
  profileButtonHover: {
    backgroundColor: tokens.colorNeutralBackground3,
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
  },
  dropdownItemHover: {
    backgroundColor: tokens.colorNeutralBackground2,
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
  },
});

// Sortable Section component
interface SortableSectionProps {
  section: Section;
  pages: PageDefinition[];
  isExpanded: boolean;
  onToggle: () => void;
  onRename: (name: string) => void;
  onDelete: () => void;
  currentPath: string;
}

function SortableSection({
  section,
  pages,
  isExpanded,
  onToggle,
  onRename,
  onDelete,
  currentPath,
}: SortableSectionProps) {
  const styles = useStyles();
  const [isHovered, setIsHovered] = useState(false);
  const [isEditing, setIsEditing] = useState(false);
  const [editName, setEditName] = useState(section.name);

  const {
    attributes,
    listeners,
    setNodeRef,
    transform,
    transition,
    isDragging,
  } = useSortable({ id: section.id });

  const style = {
    transform: CSS.Transform.toString(transform),
    transition,
  };

  const handleSaveEdit = () => {
    if (editName.trim() && editName !== section.name) {
      onRename(editName.trim());
    }
    setIsEditing(false);
  };

  const handleKeyDown = (e: React.KeyboardEvent) => {
    if (e.key === 'Enter') {
      handleSaveEdit();
    } else if (e.key === 'Escape') {
      setEditName(section.name);
      setIsEditing(false);
    }
  };

  return (
    <div
      ref={setNodeRef}
      style={style}
      className={isDragging ? styles.sectionDragging : undefined}
    >
      <div
        className={mergeClasses(
          styles.sectionHeader,
          isHovered && styles.sectionHeaderHover
        )}
        onMouseEnter={() => setIsHovered(true)}
        onMouseLeave={() => setIsHovered(false)}
        onClick={onToggle}
      >
        <div
          {...attributes}
          {...listeners}
          className={mergeClasses(
            styles.sectionDragHandle,
            isHovered && styles.sectionDragHandleVisible
          )}
          onClick={(e) => e.stopPropagation()}
        >
          <ReOrderDotsVerticalRegular fontSize={16} />
        </div>
        <span className={styles.sectionChevron}>
          {isExpanded ? (
            <ChevronDownRegular fontSize={12} />
          ) : (
            <ChevronRightRegular fontSize={12} />
          )}
        </span>
        {isEditing ? (
          <Input
            size="small"
            value={editName}
            onChange={(_, data) => setEditName(data.value)}
            onBlur={handleSaveEdit}
            onKeyDown={handleKeyDown}
            onClick={(e) => e.stopPropagation()}
            autoFocus
            className={styles.sectionTitleInput}
          />
        ) : (
          <span className={styles.sectionTitle}>{section.name}</span>
        )}
        <div
          className={mergeClasses(
            styles.sectionActions,
            isHovered && styles.sectionActionsVisible
          )}
          onClick={(e) => e.stopPropagation()}
        >
          <Menu>
            <MenuTrigger disableButtonEnhancement>
              <Button
                appearance="subtle"
                size="small"
                icon={<MoreHorizontalRegular fontSize={16} />}
              />
            </MenuTrigger>
            <MenuPopover>
              <MenuList>
                <MenuItem
                  icon={<EditRegular />}
                  onClick={() => setIsEditing(true)}
                >
                  Rename
                </MenuItem>
                <MenuItem icon={<DeleteRegular />} onClick={onDelete}>
                  Delete
                </MenuItem>
              </MenuList>
            </MenuPopover>
          </Menu>
        </div>
      </div>
      {isExpanded && pages.length > 0 && (
        <div className={styles.sectionContent}>
          {pages.map((page) => (
            <PageLink
              key={page.id}
              page={page}
              isActive={currentPath === `/app/pages/${page.id}`}
            />
          ))}
        </div>
      )}
    </div>
  );
}

// Page link component
interface PageLinkProps {
  page: PageDefinition;
  isActive: boolean;
}

function PageLink({ page, isActive }: PageLinkProps) {
  const styles = useStyles();
  const [isHovered, setIsHovered] = useState(false);
  const Icon = getPageIcon(page.icon);

  return (
    <Link
      to={`/app/pages/${page.id}`}
      className={mergeClasses(
        styles.pageLink,
        isActive && styles.pageLinkActive,
        isHovered && !isActive && styles.pageLinkHover
      )}
      onMouseEnter={() => setIsHovered(true)}
      onMouseLeave={() => setIsHovered(false)}
      title={page.description || page.name}
    >
      <Icon fontSize={18} className={styles.pageIcon} />
      <span className={styles.pageName}>{page.name}</span>
    </Link>
  );
}

function Sidebar() {
  const styles = useStyles();
  const { instance, accounts } = useMsal();
  const { setupStatus, pages, sections, saveSection, removeSection, reorderSections } = useSettings();
  const { theme, setTheme } = useTheme();
  const location = useLocation();
  const navigate = useNavigate();
  const [dropdownOpen, setDropdownOpen] = useState(false);
  const [profilePicture, setProfilePicture] = useState<string | null>(null);
  const profilePictureUrlRef = useRef<string | null>(null);
  const dropdownRef = useRef<HTMLDivElement>(null);
  const [expandedSections, setExpandedSections] = useState<Record<string, boolean>>({});
  const [homeHovered, setHomeHovered] = useState(false);
  const [profileHovered, setProfileHovered] = useState(false);

  const account = accounts[0];
  const isReady = setupStatus === 'ready';

  // Sorted sections
  const sortedSections = useMemo(() => {
    return Object.values(sections).sort((a, b) => a.order - b.order);
  }, [sections]);

  // Pages grouped by section
  const unsectionedPages = useMemo(() => {
    return pages.filter((p) => !p.sectionId);
  }, [pages]);

  const pagesBySection = useMemo(() => {
    const grouped: Record<string, PageDefinition[]> = {};
    pages.forEach((page) => {
      if (page.sectionId) {
        if (!grouped[page.sectionId]) {
          grouped[page.sectionId] = [];
        }
        grouped[page.sectionId].push(page);
      }
    });
    return grouped;
  }, [pages]);

  // dnd-kit sensors
  const sensors = useSensors(
    useSensor(PointerSensor),
    useSensor(KeyboardSensor, {
      coordinateGetter: sortableKeyboardCoordinates,
    })
  );

  // Handle section drag end
  const handleDragEnd = async (event: DragEndEvent) => {
    const { active, over } = event;
    if (over && active.id !== over.id) {
      const oldIndex = sortedSections.findIndex((s) => s.id === active.id);
      const newIndex = sortedSections.findIndex((s) => s.id === over.id);
      const newOrder = arrayMove(sortedSections, oldIndex, newIndex);
      await reorderSections(newOrder.map((s) => s.id));
    }
  };

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
        console.debug('Could not fetch profile picture:', error);
      }
    }

    fetchProfilePicture();

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

  const toggleSection = (sectionId: string) => {
    setExpandedSections((prev) => ({
      ...prev,
      [sectionId]: !prev[sectionId],
    }));
  };

  const handleRenameSection = async (section: Section, newName: string) => {
    await saveSection({ ...section, name: newName });
  };

  const handleDeleteSection = async (sectionId: string) => {
    await removeSection(sectionId);
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
                isActive('/app') && styles.menuItemActive,
                homeHovered && !isActive('/app') && styles.menuItemHover
              )}
              onMouseEnter={() => setHomeHovered(true)}
              onMouseLeave={() => setHomeHovered(false)}
            >
              <HomeRegular fontSize={22} />
              Home
            </Link>
          </li>
        </ul>

        {/* Unsectioned pages (above sections) */}
        {isReady && unsectionedPages.length > 0 && (
          <div style={{ marginTop: '8px' }}>
            {unsectionedPages.map((page) => (
              <PageLink
                key={page.id}
                page={page}
                isActive={location.pathname === `/app/pages/${page.id}`}
              />
            ))}
          </div>
        )}

        {/* Custom sections with drag-drop */}
        {isReady && sortedSections.length > 0 && (
          <DndContext
            sensors={sensors}
            collisionDetection={closestCenter}
            onDragEnd={handleDragEnd}
          >
            <SortableContext
              items={sortedSections.map((s) => s.id)}
              strategy={verticalListSortingStrategy}
            >
              {sortedSections.map((section) => (
                <SortableSection
                  key={section.id}
                  section={section}
                  pages={pagesBySection[section.id] || []}
                  isExpanded={expandedSections[section.id] !== false}
                  onToggle={() => toggleSection(section.id)}
                  onRename={(name) => handleRenameSection(section, name)}
                  onDelete={() => handleDeleteSection(section.id)}
                  currentPath={location.pathname}
                />
              ))}
            </SortableContext>
          </DndContext>
        )}

        {/* Add Page button */}
        <Button
          appearance="transparent"
          size="small"
          icon={<AddRegular />}
          className={styles.addPageButton}
          onClick={() => navigate('/app/pages/new')}
        >
          Add Page
        </Button>
      </nav>

      {/* User Profile */}
      <div className={styles.userProfile} ref={dropdownRef}>
        <div className={styles.relativeContainer}>
          <button
            onClick={() => setDropdownOpen(!dropdownOpen)}
            className={mergeClasses(
              styles.profileButton,
              profileHovered && styles.profileButtonHover
            )}
            onMouseEnter={() => setProfileHovered(true)}
            onMouseLeave={() => setProfileHovered(false)}
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
