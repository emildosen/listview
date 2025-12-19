import { useState, useEffect, useMemo, useCallback } from 'react';
import { useParams, Link as RouterLink, useNavigate } from 'react-router-dom';
import { useMsal } from '@azure/msal-react';
import {
  makeStyles,
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
import { SettingsRegular } from '@fluentui/react-icons';
import { useSettings } from '../contexts/SettingsContext';
import { getListItems, type GraphListColumn, type GraphListItem } from '../auth/graphClient';
import TableView from '../components/PageDisplay/TableView';
import ReportPageCanvas from '../components/PageDisplay/ReportPageCanvas';
import ReportCustomizeDrawer from '../components/PageDisplay/ReportCustomizeDrawer';
import LookupCustomizeDrawer from '../components/PageDisplay/LookupCustomizeDrawer';
import type { PageDefinition, ReportLayoutConfig, AnyWebPartConfig } from '../types/page';

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
  reportHeader: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'flex-start',
    marginBottom: '24px',
  },
  reportHeaderInfo: {
    display: 'flex',
    flexDirection: 'column',
    gap: '4px',
  },
  reportHeaderTitle: {
    margin: 0,
  },
  reportHeaderDescription: {
    color: tokens.colorNeutralForeground3,
    margin: 0,
  },
});

function PageDisplayPage() {
  const styles = useStyles();
  const { pageId } = useParams<{ pageId: string }>();
  const navigate = useNavigate();
  const { instance, accounts } = useMsal();
  const { pages, savePage } = useSettings();

  const [loading, setLoading] = useState(true);
  const [initialLoadComplete, setInitialLoadComplete] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [items, setItems] = useState<GraphListItem[]>([]);
  const [columns, setColumns] = useState<GraphListColumn[]>([]);
  const [filters, setFilters] = useState<Record<string, string>>({});
  const [searchText, setSearchText] = useState('');
  const [customizeDrawerOpen, setCustomizeDrawerOpen] = useState(false);

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
      setInitialLoadComplete(true);
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

  // Handle report page save
  const handleReportPageSave = useCallback(async (updatedPage: PageDefinition) => {
    await savePage(updatedPage);
  }, [savePage]);

  // Handle web part config change
  const handleWebPartConfigChange = useCallback(
    async (sectionId: string, columnId: string, config: AnyWebPartConfig) => {
      if (!page?.reportLayout) return;

      const updatedSections = page.reportLayout.sections.map((section) => {
        if (section.id !== sectionId) return section;
        return {
          ...section,
          columns: section.columns.map((column) => {
            if (column.id !== columnId) return column;
            return { ...column, webPart: config };
          }),
        };
      });

      const updatedPage = {
        ...page,
        reportLayout: { ...page.reportLayout, sections: updatedSections },
      };
      await savePage(updatedPage);
    },
    [page, savePage]
  );

  // Handle lookup page save
  const handleLookupPageSave = useCallback(async (updatedPage: PageDefinition) => {
    await savePage(updatedPage);
  }, [savePage]);

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
            <RouterLink to="/app" className={styles.breadcrumbLink}>
              Home
            </RouterLink>
          </BreadcrumbItem>
          <BreadcrumbDivider />
          <BreadcrumbItem>
            <RouterLink to="/app/pages" className={styles.breadcrumbLink}>
              Pages
            </RouterLink>
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
            <RouterLink to="/app" className={styles.breadcrumbLink}>
              Home
            </RouterLink>
          </BreadcrumbItem>
          <BreadcrumbDivider />
          <BreadcrumbItem>
            <RouterLink to="/app/pages" className={styles.breadcrumbLink}>
              Pages
            </RouterLink>
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
    // Default layout: one full-width empty section
    const defaultLayout: ReportLayoutConfig = {
      sections: [{
        id: 'default-section',
        layout: 'one-column',
        columns: [{ id: 'default-col', webPart: null }],
      }],
    };

    const currentLayout = page.reportLayout || defaultLayout;

    return (
      <div className={styles.container}>
        {/* Report Page Header */}
        <div className={styles.reportHeader}>
          <div className={styles.reportHeaderInfo}>
            <Title1 className={styles.reportHeaderTitle}>{page.name}</Title1>
            {page.description && (
              <Text className={styles.reportHeaderDescription}>{page.description}</Text>
            )}
          </div>
          <Button
            appearance="subtle"
            icon={<SettingsRegular />}
            onClick={() => setCustomizeDrawerOpen(true)}
          >
            Customize
          </Button>
        </div>

        {/* Report Page Canvas */}
        <ReportPageCanvas
          layout={currentLayout}
          onWebPartConfigChange={handleWebPartConfigChange}
        />

        {/* Customize Drawer */}
        <ReportCustomizeDrawer
          page={page}
          open={customizeDrawerOpen}
          onClose={() => setCustomizeDrawerOpen(false)}
          onSave={handleReportPageSave}
        />
      </div>
    );
  }

  return (
    <div className={styles.container}>
      {/* Lookup Page Header */}
      <div className={styles.reportHeader}>
        <div className={styles.reportHeaderInfo}>
          <Title1 className={styles.reportHeaderTitle}>{page.name}</Title1>
          {page.description && (
            <Text className={styles.reportHeaderDescription}>{page.description}</Text>
          )}
        </div>
        <Button
          appearance="subtle"
          icon={<SettingsRegular />}
          onClick={() => setCustomizeDrawerOpen(true)}
        >
          Customize
        </Button>
      </div>

      {/* Error State */}
      {error && (
        <MessageBar intent="error" className={styles.messageBar}>
          <MessageBarBody>{error}</MessageBarBody>
        </MessageBar>
      )}

      {/* Loading State - only show on initial load, not refreshes */}
      {loading && !initialLoadComplete && (
        <div className={styles.loadingContainer}>
          <Spinner size="large" />
        </div>
      )}

      {/* Main Content - keep mounted during refreshes to preserve modal state */}
      {initialLoadComplete && !error && (
        <div className={styles.mainContent}>
          <TableView
            page={page}
            columns={columns}
            items={filteredItems}
            filters={filters}
            searchText={searchText}
            onFilterChange={setFilters}
            onSearchChange={setSearchText}
            onItemCreated={loadData}
            onItemUpdated={loadData}
            onItemDeleted={loadData}
          />
        </div>
      )}

      {/* Customize Drawer */}
      <LookupCustomizeDrawer
        page={page}
        open={customizeDrawerOpen}
        onClose={() => setCustomizeDrawerOpen(false)}
        onSave={handleLookupPageSave}
      />
    </div>
  );
}

export default PageDisplayPage;
