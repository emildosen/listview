import { useState, useEffect, useMemo, useCallback } from 'react';
import { useParams, Link, useNavigate } from 'react-router-dom';
import { useMsal } from '@azure/msal-react';
import {
  makeStyles,
  Text,
  Spinner,
  Button,
  Breadcrumb,
  BreadcrumbItem,
  BreadcrumbDivider,
  MessageBar,
  MessageBarBody,
} from '@fluentui/react-components';
import { useSettings } from '../contexts/SettingsContext';
import { getListItems, type GraphListColumn, type GraphListItem } from '../auth/graphClient';
import SearchPanel from '../components/PageDisplay/SearchPanel';
import DetailPanel from '../components/PageDisplay/DetailPanel';
import TableView from '../components/PageDisplay/TableView';
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
});

function PageDisplayPage() {
  const styles = useStyles();
  const { pageId } = useParams<{ pageId: string }>();
  const navigate = useNavigate();
  const { instance, accounts } = useMsal();
  const { pages, spClient } = useSettings();

  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [items, setItems] = useState<GraphListItem[]>([]);
  const [columns, setColumns] = useState<GraphListColumn[]>([]);
  const [selectedItemId, setSelectedItemId] = useState<string | null>(null);
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

  // Selected item
  const selectedItem = useMemo(() => {
    if (!selectedItemId) return null;
    return items.find((item) => item.id === selectedItemId) || null;
  }, [items, selectedItemId]);

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

  // Page not configured
  if (!page.primarySource?.siteId || !page.primarySource?.listId) {
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
                  selectedItemId={selectedItemId}
                  onFilterChange={setFilters}
                  onSearchChange={setSearchText}
                  onSelectItem={setSelectedItemId}
                />
              </div>

              {/* Detail Panel */}
              <div className={styles.detailPanelContainer}>
                <DetailPanel
                  page={page}
                  columns={columns}
                  item={selectedItem}
                  spClient={spClient}
                />
              </div>
            </div>
          )}
        </>
      )}
    </div>
  );
}

export default PageDisplayPage;
