import { useMsal } from '@azure/msal-react';
import { useEffect, useState, useMemo, useCallback } from 'react';
import { Link, useParams, useNavigate } from 'react-router-dom';
import {
  makeStyles,
  mergeClasses,
  tokens,
  Button,
  Text,
  Title2,
  Spinner,
  Breadcrumb,
  BreadcrumbItem,
  BreadcrumbDivider,
  MessageBar,
  MessageBarBody,
  DataGrid,
  DataGridHeader,
  DataGridRow,
  DataGridHeaderCell,
  DataGridBody,
  DataGridCell,
  createTableColumn,
  TableCellLayout,
} from '@fluentui/react-components';
import type { TableColumnDefinition } from '@fluentui/react-components';
import {
  WarningRegular,
  DatabaseRegular,
  ArrowLeftRegular,
} from '@fluentui/react-icons';
import { useTheme } from '../contexts/ThemeContext';
import {
  getListById,
  getListItems,
  type GraphList,
  type GraphListColumn,
  type GraphListItem,
} from '../auth/graphClient';
import { getDefaultViewColumnOrder } from '../services/sharepoint';

interface RowData {
  _id: string;
  [key: string]: unknown;
}

const useStyles = makeStyles({
  container: {
    padding: '32px',
    flex: 1,
    display: 'flex',
    flexDirection: 'column',
  },
  breadcrumb: {
    marginBottom: '24px',
  },
  breadcrumbLink: {
    textDecoration: 'none',
    color: 'inherit',
  },
  content: {
    flex: 1,
    display: 'flex',
    flexDirection: 'column',
    minHeight: 0,
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
  // Azure style: sharp edges, subtle shadow
  card: {
    backgroundColor: tokens.colorNeutralBackground1,
    borderRadius: '2px',
    border: '1px solid transparent',
    borderImage: 'linear-gradient(135deg, rgba(0,0,0,0.05) 0%, rgba(0,0,0,0.15) 100%) 1',
    boxShadow: '0 1px 2px rgba(0, 0, 0, 0.08), 0 2px 4px rgba(0, 0, 0, 0.04)',
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
    justifyContent: 'center',
    textAlign: 'center',
    flex: 1,
  },
  emptyIcon: {
    color: tokens.colorNeutralForeground3,
    marginBottom: '16px',
  },
  emptyText: {
    color: tokens.colorNeutralForeground2,
  },
  footer: {
    marginTop: '32px',
    paddingTop: '24px',
    borderTop: `1px solid ${tokens.colorNeutralStroke1}`,
  },
  tableCard: {
    flex: 1,
    minHeight: 0,
    overflow: 'hidden',
    display: 'flex',
    flexDirection: 'column',
  },
  gridWrapper: {
    flex: 1,
    minHeight: 0,
    overflow: 'auto',
  },
  rowCount: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground2,
    marginTop: '8px',
  },
  messageBar: {
    marginBottom: '16px',
  },
});

function ListViewPage() {
  const styles = useStyles();
  const { theme } = useTheme();
  const { siteId, listId } = useParams<{ siteId: string; listId: string }>();
  const { instance, accounts } = useMsal();
  const navigate = useNavigate();
  const [list, setList] = useState<GraphList | null>(null);
  const [columns, setColumns] = useState<GraphListColumn[]>([]);
  const [items, setItems] = useState<GraphListItem[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  const account = accounts[0];

  useEffect(() => {
    if (!account || !siteId || !listId) return;

    const loadData = async () => {
      setLoading(true);
      setError(null);

      try {
        // Fetch list info and items in parallel
        const [listInfo, listData] = await Promise.all([
          getListById(instance, account, siteId, listId),
          getListItems(instance, account, siteId, listId),
        ]);

        if (!listInfo) {
          setError('List not found or access denied');
          return;
        }

        setList(listInfo);

        // Get default view column order
        let orderedColumns = listData.columns;
        if (listInfo.webUrl) {
          try {
            const columnOrder = await getDefaultViewColumnOrder(instance, account, listInfo.webUrl);
            if (columnOrder.length > 0) {
              // Sort columns according to default view order
              // Columns in the view come first (in order), then remaining columns
              const orderMap = new Map(columnOrder.map((name, index) => [name, index]));
              orderedColumns = [...listData.columns].sort((a, b) => {
                const aOrder = orderMap.get(a.name);
                const bOrder = orderMap.get(b.name);
                // If both are in the view, sort by view order
                if (aOrder !== undefined && bOrder !== undefined) {
                  return aOrder - bOrder;
                }
                // If only a is in view, a comes first
                if (aOrder !== undefined) return -1;
                // If only b is in view, b comes first
                if (bOrder !== undefined) return 1;
                // Neither in view, maintain original order
                return 0;
              });
            }
          } catch (err) {
            console.warn('Could not get default view column order:', err);
          }
        }

        setColumns(orderedColumns);
        setItems(listData.items);
      } catch (err) {
        setError(err instanceof Error ? err.message : 'Failed to load list');
      } finally {
        setLoading(false);
      }
    };

    loadData();
  }, [instance, account, siteId, listId]);

  // Format cell values for display
  const formatCellValue = useCallback((value: unknown): string => {
    if (value === null || value === undefined) return '-';
    if (typeof value === 'boolean') return value ? 'Yes' : 'No';
    if (typeof value === 'object') {
      // Handle lookup fields and other complex types
      if ('LookupValue' in (value as Record<string, unknown>)) {
        return String((value as Record<string, unknown>).LookupValue || '-');
      }
      if ('DisplayValue' in (value as Record<string, unknown>)) {
        return String((value as Record<string, unknown>).DisplayValue || '-');
      }
      return JSON.stringify(value);
    }
    return String(value);
  }, []);

  // Convert items to row data
  const rowData = useMemo((): RowData[] => {
    return items.map((item) => ({
      _id: item.id,
      ...item.fields,
    }));
  }, [items]);

  // Generate Fluent UI DataGrid column definitions
  const columnDefs = useMemo((): TableColumnDefinition<RowData>[] => {
    return columns.map((col) =>
      createTableColumn<RowData>({
        columnId: col.name,
        compare: (a, b) => {
          const aVal = String(a[col.name] ?? '');
          const bVal = String(b[col.name] ?? '');
          return aVal.localeCompare(bVal);
        },
        renderHeaderCell: () => col.displayName,
        renderCell: (item) => (
          <TableCellLayout truncate>
            {formatCellValue(item[col.name])}
          </TableCellLayout>
        ),
      })
    );
  }, [columns, formatCellValue]);

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
          <Link to="/app/lists" className={styles.breadcrumbLink}>
            Lists
          </Link>
        </BreadcrumbItem>
        <BreadcrumbDivider />
        <BreadcrumbItem>
          <Text weight="semibold">{list?.displayName || 'List'}</Text>
        </BreadcrumbItem>
      </Breadcrumb>

      <div className={styles.content}>
        {/* Header */}
        <div className={styles.header}>
          <div>
            <Title2 as="h1">{list?.displayName || 'List'}</Title2>
            {!loading && items.length > 0 && (
              <Text className={styles.description}>
                {items.length} item{items.length !== 1 ? 's' : ''}
                {items.length === 1000 && ' (limited to 1000)'}
              </Text>
            )}
          </div>
        </div>

        {/* Loading State */}
        {loading && (
          <div className={mergeClasses(styles.card, theme === 'dark' && styles.cardDark)} style={{ flex: 1 }}>
            <div className={styles.cardBody}>
              <Spinner size="large" />
              <Text className={styles.emptyText} style={{ marginTop: '16px' }}>
                Loading list data...
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

        {/* Empty State */}
        {!loading && !error && items.length === 0 && (
          <div className={mergeClasses(styles.card, theme === 'dark' && styles.cardDark)} style={{ flex: 1 }}>
            <div className={styles.cardBody}>
              <DatabaseRegular fontSize={48} className={styles.emptyIcon} />
              <Text className={styles.emptyText}>No items in this list</Text>
            </div>
          </div>
        )}

        {/* Fluent UI DataGrid */}
        {!loading && !error && items.length > 0 && (
          <div className={mergeClasses(styles.card, styles.tableCard, theme === 'dark' && styles.cardDark)}>
            <div className={styles.gridWrapper}>
              <DataGrid
                items={rowData}
                columns={columnDefs}
                sortable
                resizableColumns
                getRowId={(item) => item._id}
              >
                <DataGridHeader>
                  <DataGridRow>
                    {({ renderHeaderCell }) => (
                      <DataGridHeaderCell>{renderHeaderCell()}</DataGridHeaderCell>
                    )}
                  </DataGridRow>
                </DataGridHeader>
                <DataGridBody<RowData>>
                  {({ item, rowId }) => (
                    <DataGridRow<RowData> key={rowId}>
                      {({ renderCell }) => (
                        <DataGridCell>{renderCell(item)}</DataGridCell>
                      )}
                    </DataGridRow>
                  )}
                </DataGridBody>
              </DataGrid>
            </div>
          </div>
        )}

        {!loading && !error && items.length > 0 && (
          <Text className={styles.rowCount}>
            {items.length} row{items.length !== 1 ? 's' : ''} total
          </Text>
        )}

        {/* Back Button */}
        <div className={styles.footer}>
          <Button appearance="subtle" icon={<ArrowLeftRegular />} onClick={() => navigate('/app')}>
            Back
          </Button>
        </div>
      </div>
    </div>
  );
}

export default ListViewPage;
