import { useState, useEffect, useCallback } from 'react';
import {
  makeStyles,
  tokens,
  Table,
  TableHeader,
  TableRow,
  TableHeaderCell,
  TableBody,
  TableCell,
  Text,
  Spinner,
  mergeClasses,
} from '@fluentui/react-components';
import { TableRegular } from '@fluentui/react-icons';
import { useMsal } from '@azure/msal-react';
import { useTheme } from '../../../contexts/ThemeContext';
import type { ListItemsWebPartConfig, AnyWebPartConfig } from '../../../types/page';
import type { GraphListColumn, GraphListItem } from '../../../auth/graphClient';
import { fetchListWebPartData } from '../../../services/webPartData';
import WebPartHeader from './WebPartHeader';
import WebPartSettingsDrawer from './WebPartSettingsDrawer';

const useStyles = makeStyles({
  container: {
    backgroundColor: tokens.colorNeutralBackground1,
    borderRadius: '2px',
    border: '1px solid transparent',
    borderImage: 'linear-gradient(135deg, rgba(0,0,0,0.05) 0%, rgba(0,0,0,0.15) 100%) 1',
    boxShadow: '0 1px 2px rgba(0, 0, 0, 0.08), 0 2px 4px rgba(0, 0, 0, 0.04)',
    overflow: 'hidden',
  },
  containerDark: {
    backgroundColor: '#1a1a1a',
    borderImage: 'none',
    border: '1px solid #333333',
  },
  tableWrapper: {
    overflowX: 'auto',
    maxHeight: '400px',
    overflowY: 'auto',
  },
  table: {
    minWidth: '100%',
  },
  headerCell: {
    fontWeight: tokens.fontWeightSemibold,
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground2,
    textTransform: 'uppercase',
    letterSpacing: '0.02em',
  },
  emptyState: {
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    justifyContent: 'center',
    padding: '48px 24px',
    gap: '12px',
  },
  emptyIcon: {
    color: tokens.colorNeutralForeground3,
    fontSize: '32px',
  },
  emptyText: {
    color: tokens.colorNeutralForeground3,
    textAlign: 'center',
    fontSize: tokens.fontSizeBase200,
  },
  loadingContainer: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    padding: '48px 24px',
  },
  errorText: {
    color: tokens.colorPaletteRedForeground1,
    padding: '16px',
    textAlign: 'center',
  },
  footer: {
    padding: '8px 16px',
    borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground3,
  },
  footerDark: {
    borderTop: '1px solid #333333',
  },
});

interface ListItemsWebPartProps {
  config: ListItemsWebPartConfig;
  onConfigChange?: (config: AnyWebPartConfig) => void;
}

export default function ListItemsWebPart({ config, onConfigChange }: ListItemsWebPartProps) {
  const { theme } = useTheme();
  const styles = useStyles();
  const { instance, accounts } = useMsal();
  const account = accounts[0];

  const [items, setItems] = useState<GraphListItem[]>([]);
  const [columns, setColumns] = useState<GraphListColumn[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [totalCount, setTotalCount] = useState(0);
  const [settingsOpen, setSettingsOpen] = useState(false);

  const isConfigured = Boolean(config.dataSource?.siteId && config.dataSource?.listId);

  // Load data when config changes
  useEffect(() => {
    async function loadData() {
      if (!isConfigured || !account) {
        setItems([]);
        setColumns([]);
        return;
      }

      setLoading(true);
      setError(null);

      try {
        const result = await fetchListWebPartData(instance, account, config);
        setItems(result.items);
        setColumns(result.columns);
        setTotalCount(result.totalCount);
      } catch (err) {
        console.error('Failed to load web part data:', err);
        setError(err instanceof Error ? err.message : 'Failed to load data');
      } finally {
        setLoading(false);
      }
    }

    loadData();
  }, [config, instance, account, isConfigured]);

  const handleSettingsClick = useCallback(() => {
    setSettingsOpen(true);
  }, []);

  const handleSettingsSave = useCallback(
    (updatedConfig: AnyWebPartConfig) => {
      onConfigChange?.(updatedConfig);
      setSettingsOpen(false);
    },
    [onConfigChange]
  );

  // Get display value for a cell
  const getDisplayValue = (item: GraphListItem, columnName: string): string => {
    const value = item.fields[columnName];
    const column = columns.find((c) => c.name === columnName);

    if (value === null || value === undefined) return '-';

    // Handle lookup columns
    if (column?.lookup && typeof value === 'object') {
      return (value as { LookupValue?: string }).LookupValue || '-';
    }

    // Handle boolean
    if (column?.boolean) {
      return value ? 'Yes' : 'No';
    }

    // Handle dates
    if (column?.dateTime && typeof value === 'string') {
      return new Date(value).toLocaleDateString();
    }

    return String(value);
  };

  // Determine which columns to display
  const displayColumns = config.displayColumns && config.displayColumns.length > 0
    ? config.displayColumns
    : columns.slice(0, 5).map((c) => ({ internalName: c.name, displayName: c.displayName }));

  return (
    <div className={mergeClasses(styles.container, theme === 'dark' && styles.containerDark)}>
      <WebPartHeader
        title={config.title}
        isConfigured={isConfigured}
        onSettingsClick={handleSettingsClick}
      />

      {/* Loading state */}
      {loading && (
        <div className={styles.loadingContainer}>
          <Spinner size="small" label="Loading data..." />
        </div>
      )}

      {/* Error state */}
      {error && !loading && <Text className={styles.errorText}>{error}</Text>}

      {/* Empty/Not configured state */}
      {!loading && !error && !isConfigured && (
        <div className={styles.emptyState}>
          <TableRegular className={styles.emptyIcon} />
          <Text className={styles.emptyText}>
            Click the settings icon to configure this web part
          </Text>
        </div>
      )}

      {/* No data state */}
      {!loading && !error && isConfigured && items.length === 0 && (
        <div className={styles.emptyState}>
          <Text className={styles.emptyText}>No items found</Text>
        </div>
      )}

      {/* Data table */}
      {!loading && !error && isConfigured && items.length > 0 && (
        <>
          <div className={styles.tableWrapper}>
            <Table className={styles.table}>
              <TableHeader>
                <TableRow>
                  {displayColumns.map((col) => (
                    <TableHeaderCell key={col.internalName} className={styles.headerCell}>
                      {col.displayName}
                    </TableHeaderCell>
                  ))}
                </TableRow>
              </TableHeader>
              <TableBody>
                {items.map((item) => (
                  <TableRow key={item.id}>
                    {displayColumns.map((col) => (
                      <TableCell key={col.internalName}>
                        {getDisplayValue(item, col.internalName)}
                      </TableCell>
                    ))}
                  </TableRow>
                ))}
              </TableBody>
            </Table>
          </div>
          <div className={mergeClasses(styles.footer, theme === 'dark' && styles.footerDark)}>
            Showing {items.length} of {totalCount} items
          </div>
        </>
      )}

      {/* Settings Drawer */}
      <WebPartSettingsDrawer
        webPart={config}
        open={settingsOpen}
        onClose={() => setSettingsOpen(false)}
        onSave={handleSettingsSave}
      />
    </div>
  );
}
