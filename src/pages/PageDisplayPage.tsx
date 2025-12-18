import { useState, useEffect, useMemo, useCallback } from 'react';
import { useParams, Link, useNavigate } from 'react-router-dom';
import { useMsal } from '@azure/msal-react';
import {
  makeStyles,
  mergeClasses,
  tokens,
  Text,
  Title1,
  Spinner,
  Button,
  Breadcrumb,
  BreadcrumbItem,
  BreadcrumbDivider,
  MessageBar,
  MessageBarBody,
} from '@fluentui/react-components';
import { DocumentTextRegular, DataPieRegular } from '@fluentui/react-icons';
import { useSettings } from '../contexts/SettingsContext';
import { useTheme } from '../contexts/ThemeContext';
import { getListItems, type GraphListColumn, type GraphListItem } from '../auth/graphClient';
import SearchPanel from '../components/PageDisplay/SearchPanel';
import TableView from '../components/PageDisplay/TableView';
import ItemDetailModal from '../components/PageDisplay/ItemDetailModal';
import type { PageDefinition } from '../types/page';

const useStyles = makeStyles({
  container: {
    padding: '32px',
    flex: 1,
  },
  breadcrumb: {
    marginBottom: '16px',
  },
  breadcrumbLink: {
    textDecoration: 'none',
    color: 'inherit',
  },
  loadingContainer: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    padding: '48px',
  },
  backLink: {
    marginTop: '16px',
  },
  messageBar: {
    marginBottom: '16px',
  },
  mainContent: {
    height: 'calc(100% - 4rem)',
  },
  twoPanelLayout: {
    display: 'flex',
    gap: '24px',
    height: 'calc(100% - 4rem)',
  },
  searchPanelContainer: {
    width: '320px',
    flexShrink: 0,
  },
  detailPanelContainer: {
    flex: 1,
    overflow: 'auto',
  },
  emptyState: {
    flex: 1,
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    justifyContent: 'center',
    gap: '12px',
    color: tokens.colorNeutralForeground2,
    backgroundColor: tokens.colorNeutralBackground2,
    borderRadius: '8px',
  },
  emptyStateDark: {
    backgroundColor: '#1a1a1a',
  },
  emptyIcon: {
    opacity: 0.3,
  },
  reportContainer: {
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    justifyContent: 'center',
    height: 'calc(100vh - 200px)',
    gap: '16px',
  },
  reportIcon: {
    color: tokens.colorNeutralForeground3,
    opacity: 0.5,
  },
  reportTitle: {
    color: tokens.colorNeutralForeground1,
    textAlign: 'center',
  },
  reportSubtitle: {
    color: tokens.colorNeutralForeground3,
    textAlign: 'center',
  },
});

function PageDisplayPage() {
  const styles = useStyles();
  const { theme } = useTheme();
  const { pageId } = useParams<{ pageId: string }>();
  const navigate = useNavigate();
  const { instance, accounts } = useMsal();
  const { pages, spClient, savePage } = useSettings();

  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [items, setItems] = useState<GraphListItem[]>([]);
  const [columns, setColumns] = useState<GraphListColumn[]>([]);
  const [modalItem, setModalItem] = useState<GraphListItem | null>(null);
  const [filters, setFilters] = useState<Record<string, string>>({});
  const [searchText, setSearchText] = useState('');

  // Find page config
  const page = useMemo((): PageDefinition | undefined => {
    if (!pageId) return undefined;
    return pages.find((p) => p.id === pageId);
  }, [pageId, pages]);

  const account = accounts[0];

  // Load primary list data
  const loadData = useCallback(async () => {
    if (!page || !account || !page.primarySource?.siteId || !page.primarySource?.listId) {
      setLoading(false);
      return;
    }

    setLoading(true);
    setError(null);

    try {
      const result = await getListItems(
        instance,
        account,
        page.primarySource.siteId,
        page.primarySource.listId
      );
      setColumns(result.columns);
      setItems(result.items);
    } catch (err) {
      console.error('Failed to load data:', err);
      setError(err instanceof Error ? err.message : 'Failed to load data');
    } finally {
      setLoading(false);
    }
  }, [page, instance, account]);

  useEffect(() => {
    loadData();
  }, [loadData]);

  // Filter items based on search and filters
  const filteredItems = useMemo(() => {
    if (!page?.searchConfig) return items;

    return items.filter((item) => {
      // Apply dropdown filters
      for (const [column, value] of Object.entries(filters)) {
        if (value && String(item.fields[column] || '') !== value) {
          return false;
        }
      }

      // Apply text search
      if (searchText && page.searchConfig.textSearchColumns.length > 0) {
        const searchLower = searchText.toLowerCase();
        const matchesSearch = page.searchConfig.textSearchColumns.some((col) => {
          const fieldValue = item.fields[col];
          if (fieldValue === null || fieldValue === undefined) return false;
          return String(fieldValue).toLowerCase().includes(searchLower);
        });
        if (!matchesSearch) return false;
      }

      return true;
    });
  }, [items, filters, searchText, page?.searchConfig]);

  // Handle page configuration updates
  const handlePageUpdate = useCallback(async (updatedPage: PageDefinition) => {
    await savePage(updatedPage);
  }, [savePage]);

  // Handle item selection from SearchPanel (for inline mode)
  const handleSelectItem = useCallback((itemId: string | null) => {
    if (itemId) {
      const item = items.find(i => i.id === itemId);
      if (item) setModalItem(item);
    }
  }, [items]);

  // Loading state for pages
  if (pages.length === 0) {
    return (
      <div className={styles.container}>
        <div className={styles.loadingContainer}>
          <Spinner size="large" />
        </div>
      </div>
    );
  }

  // Page not found
  if (!page) {
    return (
      <div className={styles.container}>
        <Breadcrumb className={styles.breadcrumb}>
          <BreadcrumbItem>
            <Link to="/app" className={styles.breadcrumbLink}>
              Home
            </Link>
          </BreadcrumbItem>
          <BreadcrumbDivider />
          <BreadcrumbItem>
            <Link to="/app/pages" className={styles.breadcrumbLink}>
              Pages
            </Link>
          </BreadcrumbItem>
          <BreadcrumbDivider />
          <BreadcrumbItem>
            <Text weight="semibold">Not Found</Text>
          </BreadcrumbItem>
        </Breadcrumb>
        <MessageBar intent="error" className={styles.messageBar}>
          <MessageBarBody>Page not found</MessageBarBody>
        </MessageBar>
        <div className={styles.backLink}>
          <Button appearance="subtle" onClick={() => navigate('/app/pages')}>
            Back to Pages
          </Button>
        </div>
      </div>
    );
  }

  // Page not configured (only applies to lookup pages)
  if (page.pageType === 'lookup' && (!page.primarySource?.siteId || !page.primarySource?.listId)) {
    return (
      <div className={styles.container}>
        <Breadcrumb className={styles.breadcrumb}>
          <BreadcrumbItem>
            <Link to="/app" className={styles.breadcrumbLink}>
              Home
            </Link>
          </BreadcrumbItem>
          <BreadcrumbDivider />
          <BreadcrumbItem>
            <Link to="/app/pages" className={styles.breadcrumbLink}>
              Pages
            </Link>
          </BreadcrumbItem>
          <BreadcrumbDivider />
          <BreadcrumbItem>
            <Text weight="semibold">{page.name}</Text>
          </BreadcrumbItem>
        </Breadcrumb>
        <MessageBar intent="warning" className={styles.messageBar}>
          <MessageBarBody>
            This page is not fully configured. Please edit the page to select a primary list.
          </MessageBarBody>
        </MessageBar>
        <div className={styles.backLink}>
          <Button appearance="primary" onClick={() => navigate(`/app/pages/${page.id}/edit`)}>
            Configure Page
          </Button>
        </div>
      </div>
    );
  }

  // Report page display
  if (page.pageType === 'report') {
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
            <Link to="/app/pages" className={styles.breadcrumbLink}>
              Pages
            </Link>
          </BreadcrumbItem>
          <BreadcrumbDivider />
          <BreadcrumbItem>
            <Text weight="semibold">{page.name}</Text>
          </BreadcrumbItem>
        </Breadcrumb>

        {/* Report Page Content */}
        <div className={styles.reportContainer}>
          <DataPieRegular fontSize={64} className={styles.reportIcon} />
          <Title1 className={styles.reportTitle}>{page.name}</Title1>
          {page.description && (
            <Text className={styles.reportSubtitle}>{page.description}</Text>
          )}
        </div>
      </div>
    );
  }

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
          <Link to="/app/pages" className={styles.breadcrumbLink}>
            Pages
          </Link>
        </BreadcrumbItem>
        <BreadcrumbDivider />
        <BreadcrumbItem>
          <Text weight="semibold">{page.name}</Text>
        </BreadcrumbItem>
      </Breadcrumb>

      {/* Error State */}
      {error && (
        <MessageBar intent="error" className={styles.messageBar}>
          <MessageBarBody>{error}</MessageBarBody>
        </MessageBar>
      )}

      {/* Loading State */}
      {loading && (
        <div className={styles.loadingContainer}>
          <Spinner size="large" />
        </div>
      )}

      {/* Main Content */}
      {!loading && !error && (
        <>
          {page.searchConfig?.displayMode === 'table' ? (
            /* Table View */
            <div className={styles.mainContent}>
              <TableView
                page={page}
                columns={columns}
                items={filteredItems}
                filters={filters}
                searchText={searchText}
                onFilterChange={setFilters}
                onSearchChange={setSearchText}
                spClient={spClient}
                onPageUpdate={handlePageUpdate}
              />
            </div>
          ) : (
            /* Inline View - Two Panel Layout */
            <div className={styles.twoPanelLayout}>
              {/* Search Panel */}
              <div className={styles.searchPanelContainer}>
                <SearchPanel
                  page={page}
                  columns={columns}
                  items={filteredItems}
                  filters={filters}
                  searchText={searchText}
                  selectedItemId={null}
                  onFilterChange={setFilters}
                  onSearchChange={setSearchText}
                  onSelectItem={handleSelectItem}
                />
              </div>

              {/* Empty State (click an item to open details) */}
              <div className={mergeClasses(styles.emptyState, theme === 'dark' && styles.emptyStateDark)}>
                <DocumentTextRegular fontSize={48} className={styles.emptyIcon} />
                <Text>Select an item to view details</Text>
              </div>
            </div>
          )}

          {/* Item Detail Modal (for both modes) */}
          {modalItem && page && (
            <ItemDetailModal
              page={page}
              columns={columns}
              item={modalItem}
              spClient={spClient}
              onClose={() => setModalItem(null)}
              onPageUpdate={handlePageUpdate}
            />
          )}
        </>
      )}
    </div>
  );
}

export default PageDisplayPage;
