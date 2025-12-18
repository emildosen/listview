import { useMsal } from '@azure/msal-react';
import { useEffect, useState, useCallback } from 'react';
import { Link, useNavigate } from 'react-router-dom';
import {
  makeStyles,
  mergeClasses,
  tokens,
  Button,
  Text,
  Title2,
  Badge,
  Spinner,
  Input,
  Checkbox,
  Breadcrumb,
  BreadcrumbItem,
  BreadcrumbDivider,
  MessageBar,
  MessageBarBody,
  Table,
  TableHeader,
  TableRow,
  TableHeaderCell,
  TableBody,
  TableCell,
  TableCellLayout,
} from '@fluentui/react-components';
import {
  SearchRegular,
  DismissRegular,
  WarningRegular,
  DatabaseRegular,
  ArrowLeftRegular,
} from '@fluentui/react-icons';
import { useSettings } from '../contexts/SettingsContext';
import { useTheme } from '../contexts/ThemeContext';
import {
  getAllSites,
  getSiteLists,
  type GraphSite,
  type GraphList,
} from '../auth/graphClient';
import { SYSTEM_LIST_NAMES } from '../services/sharepoint';

export interface ListRow {
  siteId: string;
  siteName: string;
  siteUrl: string;
  listId: string;
  listName: string;
}

export const ENABLED_LISTS_KEY = 'EnabledLists';

const useStyles = makeStyles({
  container: {
    padding: '32px',
    flex: 1,
  },
  breadcrumb: {
    marginBottom: '24px',
  },
  breadcrumbLink: {
    textDecoration: 'none',
    color: 'inherit',
  },
  content: {
    maxWidth: '896px',
  },
  header: {
    display: 'flex',
    alignItems: 'flex-start',
    justifyContent: 'space-between',
    marginBottom: '24px',
  },
  description: {
    color: tokens.colorNeutralForeground2,
    marginTop: '4px',
  },
  searchWrapper: {
    marginBottom: '16px',
    position: 'relative',
  },
  searchInput: {
    width: '100%',
  },
  // Azure style: sharp edges, subtle shadow
  card: {
    backgroundColor: tokens.colorNeutralBackground1,
    borderRadius: '2px',
    border: '1px solid transparent',
    borderImage: 'linear-gradient(135deg, rgba(0,0,0,0.05) 0%, rgba(0,0,0,0.15) 100%) 1',
    boxShadow: '0 1px 2px rgba(0, 0, 0, 0.08), 0 2px 4px rgba(0, 0, 0, 0.04)',
    marginBottom: '16px',
    overflow: 'hidden',
  },
  cardDark: {
    backgroundColor: '#1a1a1a',
    borderImage: 'none',
    border: '1px solid #333333',
  },
  cardBody: {
    padding: '48px',
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    textAlign: 'center',
  },
  emptyIcon: {
    color: tokens.colorNeutralForeground3,
    marginBottom: '16px',
  },
  emptyText: {
    color: tokens.colorNeutralForeground2,
    marginBottom: '8px',
  },
  emptySubtext: {
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase200,
  },
  footer: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    marginTop: '32px',
    paddingTop: '24px',
    borderTop: `1px solid ${tokens.colorNeutralStroke1}`,
  },
  footerActions: {
    display: 'flex',
    alignItems: 'center',
    gap: '12px',
  },
  unsavedText: {
    color: tokens.colorPaletteYellowForeground1,
    fontSize: tokens.fontSizeBase200,
  },
  tableHeader: {
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  tableHeaderCell: {
    fontWeight: tokens.fontWeightSemibold,
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground2,
    textTransform: 'uppercase',
    letterSpacing: '0.02em',
  },
  tableRow: {
    cursor: 'pointer',
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    ':hover': {
      backgroundColor: tokens.colorNeutralBackground1Hover,
    },
    ':last-child': {
      borderBottom: 'none',
    },
  },
  tableRowSelected: {
    backgroundColor: tokens.colorBrandBackground2,
  },
  siteName: {
    color: tokens.colorNeutralForeground2,
  },
  messageBar: {
    marginBottom: '16px',
  },
});

function DataPage() {
  const styles = useStyles();
  const { theme } = useTheme();
  const { instance, accounts } = useMsal();
  const navigate = useNavigate();
  const { getSetting, updateSetting } = useSettings();
  const [lists, setLists] = useState<ListRow[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [selectedLists, setSelectedLists] = useState<Set<string>>(new Set());
  const [savedLists, setSavedLists] = useState<Set<string>>(new Set());
  const [saving, setSaving] = useState(false);
  const [searchQuery, setSearchQuery] = useState('');

  const account = accounts[0];

  const getListKey = (siteId: string, listId: string) => `${siteId}:${listId}`;

  // Load all sites and their lists
  useEffect(() => {
    if (!account) return;

    const loadData = async () => {
      setLoading(true);
      setError(null);

      try {
        // Fetch all sites
        const sites = await getAllSites(instance, account);

        // Fetch lists for each site in parallel
        const listsPromises = sites.map(async (site: GraphSite) => {
          try {
            const siteLists = await getSiteLists(instance, account, site.id);
            return siteLists.map((list: GraphList) => ({
              siteId: site.id,
              siteName: site.displayName || site.name,
              siteUrl: site.webUrl,
              listId: list.id,
              listName: list.displayName || list.name,
            }));
          } catch {
            // Skip sites where we can't fetch lists
            return [];
          }
        });

        const listsArrays = await Promise.all(listsPromises);
        const allLists = listsArrays.flat();

        // Filter out system lists used by ListView app
        const userLists = allLists.filter(
          (list) => !SYSTEM_LIST_NAMES.includes(list.listName as typeof SYSTEM_LIST_NAMES[number])
        );

        setLists(userLists);

        // Load saved enabled lists from settings
        const savedJson = getSetting(ENABLED_LISTS_KEY);
        if (savedJson) {
          try {
            const saved = JSON.parse(savedJson);
            // Handle both new format (array of objects) and old format (array of keys)
            let keys: string[];
            if (Array.isArray(saved) && saved.length > 0 && typeof saved[0] === 'object') {
              keys = (saved as ListRow[]).map((l) => getListKey(l.siteId, l.listId));
            } else {
              keys = saved as string[];
            }
            const savedSet = new Set(keys);
            setSelectedLists(savedSet);
            setSavedLists(savedSet);
          } catch {
            // Invalid JSON, ignore
          }
        }
      } catch (err) {
        setError(err instanceof Error ? err.message : 'Failed to load data');
      } finally {
        setLoading(false);
      }
    };

    loadData();
  }, [instance, account, getSetting]);

  const handleToggleList = useCallback((siteId: string, listId: string) => {
    const key = getListKey(siteId, listId);
    setSelectedLists((prev) => {
      const next = new Set(prev);
      if (next.has(key)) {
        next.delete(key);
      } else {
        next.add(key);
      }
      return next;
    });
  }, []);

  const handleSelectAll = useCallback(() => {
    if (selectedLists.size === lists.length) {
      setSelectedLists(new Set());
    } else {
      setSelectedLists(new Set(lists.map((l) => getListKey(l.siteId, l.listId))));
    }
  }, [lists, selectedLists.size]);

  const handleSave = useCallback(async () => {
    setSaving(true);
    try {
      // Store full list info, not just keys
      const enabledListObjects = lists.filter((list) =>
        selectedLists.has(getListKey(list.siteId, list.listId))
      );
      await updateSetting(ENABLED_LISTS_KEY, JSON.stringify(enabledListObjects));
      setSavedLists(new Set(selectedLists));
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Failed to save');
    } finally {
      setSaving(false);
    }
  }, [selectedLists, lists, updateSetting]);

  const handleCancel = useCallback(() => {
    setSelectedLists(new Set(savedLists));
  }, [savedLists]);

  const hasChanges =
    selectedLists.size !== savedLists.size ||
    [...selectedLists].some((key) => !savedLists.has(key));

  // Filter lists based on search query
  const filteredLists = lists.filter((list) => {
    if (!searchQuery.trim()) return true;
    const query = searchQuery.toLowerCase();
    return (
      list.listName.toLowerCase().includes(query) ||
      list.siteName.toLowerCase().includes(query)
    );
  });

  return (
    <div className={styles.container}>
      {/* Breadcrumb */}
      <Breadcrumb className={styles.breadcrumb}>
        <BreadcrumbItem>
          <Link to="/app" className={styles.breadcrumbLink}>
            Home
          </Link>
        </BreadcrumbItem>
        <BreadcrumbDivider />
        <BreadcrumbItem>
          <Text weight="semibold">Lists</Text>
        </BreadcrumbItem>
      </Breadcrumb>

      <div className={styles.content}>
        <div className={styles.header}>
          <div>
            <Title2 as="h1">Manage Lists</Title2>
            <Text className={styles.description}>
              Select which SharePoint lists to enable.
            </Text>
          </div>
          {selectedLists.size > 0 && (
            <Badge appearance="filled" color="brand" size="large">
              {selectedLists.size} list{selectedLists.size !== 1 ? 's' : ''} selected
            </Badge>
          )}
        </div>

        {/* Search Bar */}
        {!loading && lists.length > 0 && (
          <div className={styles.searchWrapper}>
            <Input
              className={styles.searchInput}
              placeholder="Search lists..."
              value={searchQuery}
              onChange={(_e, data) => setSearchQuery(data.value)}
              contentBefore={<SearchRegular />}
              contentAfter={
                searchQuery ? (
                  <Button
                    appearance="subtle"
                    size="small"
                    icon={<DismissRegular />}
                    onClick={() => setSearchQuery('')}
                  />
                ) : undefined
              }
            />
          </div>
        )}

        {/* Loading State */}
        {loading && (
          <div className={mergeClasses(styles.card, theme === 'dark' && styles.cardDark)}>
            <div className={styles.cardBody}>
              <Spinner size="large" />
              <Text className={styles.emptyText} style={{ marginTop: '16px' }}>
                Loading sites and lists...
              </Text>
            </div>
          </div>
        )}

        {/* Error State */}
        {error && !loading && (
          <MessageBar intent="error" className={styles.messageBar}>
            <MessageBarBody>
              <WarningRegular /> {error}
            </MessageBarBody>
          </MessageBar>
        )}

        {/* No Lists */}
        {!loading && !error && lists.length === 0 && (
          <div className={mergeClasses(styles.card, theme === 'dark' && styles.cardDark)}>
            <div className={styles.cardBody}>
              <DatabaseRegular fontSize={48} className={styles.emptyIcon} />
              <Text className={styles.emptyText}>No lists found</Text>
              <Text className={styles.emptySubtext}>
                No SharePoint lists available
              </Text>
            </div>
          </div>
        )}

        {/* No Search Results */}
        {!loading && lists.length > 0 && filteredLists.length === 0 && searchQuery && (
          <div className={mergeClasses(styles.card, theme === 'dark' && styles.cardDark)}>
            <div className={styles.cardBody}>
              <SearchRegular fontSize={48} className={styles.emptyIcon} />
              <Text className={styles.emptyText}>No lists match "{searchQuery}"</Text>
              <Button
                appearance="subtle"
                size="small"
                onClick={() => setSearchQuery('')}
                style={{ marginTop: '8px' }}
              >
                Clear search
              </Button>
            </div>
          </div>
        )}

        {/* Lists Table */}
        {!loading && filteredLists.length > 0 && (
          <div className={mergeClasses(styles.card, theme === 'dark' && styles.cardDark)}>
            <Table>
              <TableHeader className={styles.tableHeader}>
                <TableRow>
                  <TableHeaderCell style={{ width: '48px' }}>
                    <Checkbox
                      checked={selectedLists.size === lists.length && lists.length > 0}
                      onChange={handleSelectAll}
                    />
                  </TableHeaderCell>
                  <TableHeaderCell className={styles.tableHeaderCell}>List Name</TableHeaderCell>
                  <TableHeaderCell className={styles.tableHeaderCell}>Site</TableHeaderCell>
                </TableRow>
              </TableHeader>
              <TableBody>
                {filteredLists.map((list) => {
                  const key = getListKey(list.siteId, list.listId);
                  const isSelected = selectedLists.has(key);
                  return (
                    <TableRow
                      key={key}
                      className={mergeClasses(styles.tableRow, isSelected && styles.tableRowSelected)}
                      onClick={() => handleToggleList(list.siteId, list.listId)}
                    >
                      <TableCell>
                        <Checkbox
                          checked={isSelected}
                          onChange={() => handleToggleList(list.siteId, list.listId)}
                          onClick={(e) => e.stopPropagation()}
                        />
                      </TableCell>
                      <TableCell>
                        <TableCellLayout>
                          <Text weight="medium">{list.listName}</Text>
                        </TableCellLayout>
                      </TableCell>
                      <TableCell>
                        <Text className={styles.siteName}>{list.siteName}</Text>
                      </TableCell>
                    </TableRow>
                  );
                })}
              </TableBody>
            </Table>
          </div>
        )}

        {/* Action Buttons */}
        <div className={styles.footer}>
          <Button appearance="subtle" icon={<ArrowLeftRegular />} onClick={() => navigate('/app')}>
            Back
          </Button>

          <div className={styles.footerActions}>
            {hasChanges && (
              <Text className={styles.unsavedText}>Unsaved changes</Text>
            )}
            <Button
              appearance="subtle"
              onClick={handleCancel}
              disabled={!hasChanges || saving}
            >
              Cancel
            </Button>
            <Button
              appearance="primary"
              onClick={handleSave}
              disabled={!hasChanges || saving}
              icon={saving ? <Spinner size="tiny" /> : undefined}
            >
              {saving ? 'Saving...' : 'Save'}
            </Button>
          </div>
        </div>
      </div>
    </div>
  );
}

export default DataPage;
