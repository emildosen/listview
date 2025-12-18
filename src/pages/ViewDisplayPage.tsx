import { useState, useEffect, useMemo } from 'react';
import { useParams, Link, useNavigate } from 'react-router-dom';
import { useMsal } from '@azure/msal-react';
import {
  makeStyles,
  tokens,
  Text,
  Title2,
  Badge,
  Button,
  Spinner,
  Card,
  Breadcrumb,
  BreadcrumbItem,
  BreadcrumbDivider,
  MessageBar,
  MessageBarBody,
  DataGrid,
  DataGridHeader,
  DataGridHeaderCell,
  DataGridBody,
  DataGridRow,
  DataGridCell,
  TableCellLayout,
  createTableColumn,
} from '@fluentui/react-components';
import type { TableColumnDefinition } from '@fluentui/react-components';
import {
  SettingsRegular,
  WarningRegular,
  DatabaseRegular,
  ArrowLeftRegular,
} from '@fluentui/react-icons';
import { useSettings } from '../contexts/SettingsContext';
import { getListItems, type GraphListItem, type GraphListColumn } from '../auth/graphClient';

import type { ViewDefinition, ViewFilter, FilterOperator, ViewSorting } from '../types/view';

interface RowData {
  _sourceListId: string;
  _sourceListName: string;
  _itemId?: string;  // SharePoint item ID for JOIN matching
  [key: string]: unknown;
}

interface AggregateResult {
  [key: string]: number | string;
}

// Represents a lookup relationship between two sources
interface LookupRelationship {
  childListId: string;       // The list that HAS the lookup column
  parentListId: string;      // The list that is LOOKED UP TO
  lookupColumnName: string;  // The column name in the child list (e.g., "Student")
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
    flex: '1',
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
  titleRow: {
    display: 'flex',
    alignItems: 'center',
    gap: '12px',
  },
  description: {
    color: tokens.colorNeutralForeground2,
    marginTop: '4px',
  },
  meta: {
    display: 'flex',
    alignItems: 'center',
    gap: '16px',
    marginTop: '8px',
    fontSize: '14px',
    color: tokens.colorNeutralForeground2,
  },
  loadingCard: {
    flex: '1',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
  },
  loadingContent: {
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    gap: '16px',
  },
  emptyCard: {
    flex: '1',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
  },
  emptyContent: {
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    gap: '8px',
    color: tokens.colorNeutralForeground3,
  },
  emptyIcon: {
    fontSize: '48px',
    marginBottom: '8px',
    color: tokens.colorNeutralForeground4,
  },
  dataGrid: {
    minWidth: '100%',
  },
  rowCount: {
    marginTop: '8px',
    fontSize: '14px',
    color: tokens.colorNeutralForeground2,
  },
  footer: {
    marginTop: '32px',
    paddingTop: '24px',
    borderTop: `1px solid ${tokens.colorNeutralStroke1}`,
  },
  backLink: {
    marginTop: '16px',
  },
  messageBar: {
    marginBottom: '16px',
  },
});

function ViewDisplayPage() {
  const styles = useStyles();
  const { viewId } = useParams<{ viewId: string }>();
  const navigate = useNavigate();
  const { instance, accounts } = useMsal();
  const { views } = useSettings();
  const account = accounts[0];

  const [view, setView] = useState<ViewDefinition | null>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [rawData, setRawData] = useState<RowData[]>([]);
  const [lookupRelationships, setLookupRelationships] = useState<LookupRelationship[]>([]);

  // Find the view
  useEffect(() => {
    const found = views.find((v) => v.id === viewId);
    setView(found || null);
  }, [viewId, views]);

  // Load data from all sources
  useEffect(() => {
    if (!view || !account) {
      setLoading(false);
      return;
    }

    const loadData = async () => {
      setLoading(true);
      setError(null);

      try {
        const allRows: RowData[] = [];
        const colsByList = new Map<string, GraphListColumn[]>();
        const sourceListIds = new Set(view.sources.map(s => s.listId));

        for (const source of view.sources) {
          const result = await getListItems(instance, account, source.siteId, source.listId);

          // Store column metadata for this list
          colsByList.set(source.listId, result.columns);

          // Map items to row data, including the item ID for JOIN matching
          const rows = result.items.map((item: GraphListItem) => ({
            _sourceListId: source.listId,
            _sourceListName: source.listName,
            _itemId: item.id,
            ...item.fields,
          }));

          allRows.push(...rows);
        }

        // Detect lookup relationships between sources
        const relationships: LookupRelationship[] = [];
        for (const [listId, columns] of colsByList) {
          for (const col of columns) {
            if (col.lookup?.listId && sourceListIds.has(col.lookup.listId)) {
              // This list has a lookup column pointing to another source in the view
              relationships.push({
                childListId: listId,
                parentListId: col.lookup.listId,
                lookupColumnName: col.name,
              });
            }
          }
        }

        setLookupRelationships(relationships);
        setRawData(allRows);
      } catch (err) {
        console.error('Failed to load view data:', err);
        setError(err instanceof Error ? err.message : 'Failed to load data');
      } finally {
        setLoading(false);
      }
    };

    loadData();
  }, [view, instance, account]);

  // Apply filters
  const applyFilter = (row: RowData, filter: ViewFilter): boolean => {
    const value = row[filter.column];
    const filterValue = filter.value;

    const stringValue = String(value ?? '').toLowerCase();
    const stringFilterValue = filterValue.toLowerCase();

    const operators: Record<FilterOperator, () => boolean> = {
      eq: () => stringValue === stringFilterValue,
      ne: () => stringValue !== stringFilterValue,
      gt: () => {
        const numValue = parseFloat(String(value));
        const numFilter = parseFloat(filterValue);
        return !isNaN(numValue) && !isNaN(numFilter) && numValue > numFilter;
      },
      lt: () => {
        const numValue = parseFloat(String(value));
        const numFilter = parseFloat(filterValue);
        return !isNaN(numValue) && !isNaN(numFilter) && numValue < numFilter;
      },
      contains: () => stringValue.includes(stringFilterValue),
    };

    return operators[filter.operator]();
  };

  // Process data based on view mode
  const processedData = useMemo(() => {
    if (!view || rawData.length === 0) return [];

    // Convert value to number, treating booleans as 1/0
    const toNumber = (v: unknown): number => {
      if (typeof v === 'boolean') return v ? 1 : 0;
      const num = parseFloat(String(v));
      return isNaN(num) ? NaN : num;
    };

    // Multi-level sort function
    const applySorting = <T extends Record<string, unknown>>(
      data: T[],
      sortRules: ViewSorting | undefined
    ): T[] => {
      if (!sortRules || sortRules.length === 0) return data;

      return [...data].sort((a, b) => {
        for (const rule of sortRules) {
          const aVal = a[rule.column];
          const bVal = b[rule.column];

          // Handle nulls
          if (aVal === null || aVal === undefined) {
            if (bVal === null || bVal === undefined) continue;
            return rule.direction === 'asc' ? 1 : -1;
          }
          if (bVal === null || bVal === undefined) {
            return rule.direction === 'asc' ? -1 : 1;
          }

          // Try numeric comparison first
          const aNum = toNumber(aVal);
          const bNum = toNumber(bVal);

          let cmp: number;
          if (!isNaN(aNum) && !isNaN(bNum)) {
            cmp = aNum - bNum;
          } else {
            // Fall back to string comparison
            const aStr = String(aVal).toLowerCase();
            const bStr = String(bVal).toLowerCase();
            cmp = aStr.localeCompare(bStr, undefined, { numeric: true });
          }

          if (cmp !== 0) {
            return rule.direction === 'asc' ? cmp : -cmp;
          }
        }
        return 0;
      });
    };

    // Compute aggregation for a set of rows
    const computeAggregation = (
      rows: RowData[],
      col: { internalName: string; aggregation?: string }
    ): number => {
      const values = rows
        .map((row) => row[col.internalName])
        .filter((v) => v !== null && v !== undefined);

      switch (col.aggregation) {
        case 'count':
          return values.length;
        case 'sum': {
          const numValues = values.map(toNumber).filter((n) => !isNaN(n));
          return numValues.reduce((a, b) => a + b, 0);
        }
        case 'avg': {
          const numValues = values.map(toNumber).filter((n) => !isNaN(n));
          return numValues.length > 0
            ? Math.round((numValues.reduce((a, b) => a + b, 0) / numValues.length) * 100) / 100
            : 0;
        }
        case 'min': {
          const numValues = values.map(toNumber).filter((n) => !isNaN(n));
          return numValues.length > 0 ? Math.min(...numValues) : 0;
        }
        case 'max': {
          const numValues = values.map(toNumber).filter((n) => !isNaN(n));
          return numValues.length > 0 ? Math.max(...numValues) : 0;
        }
        default:
          return 0;
      }
    };

    // Compute aggregation for a column that comes from joined child rows
    const computeJoinedAggregation = (
      rows: RowData[],
      col: { internalName: string; aggregation?: string },
      childListId: string
    ): number => {
      // Collect all values from the child rows stored in _childRows_<listId>
      const allChildValues: unknown[] = [];

      for (const row of rows) {
        const childRows = row[`_childRows_${childListId}`] as RowData[] | undefined;
        if (childRows && Array.isArray(childRows)) {
          for (const childRow of childRows) {
            const value = childRow[col.internalName];
            if (value !== null && value !== undefined) {
              allChildValues.push(value);
            }
          }
        }
      }

      switch (col.aggregation) {
        case 'count':
          return allChildValues.length;
        case 'sum': {
          const numValues = allChildValues.map(toNumber).filter((n) => !isNaN(n));
          return numValues.reduce((a, b) => a + b, 0);
        }
        case 'avg': {
          const numValues = allChildValues.map(toNumber).filter((n) => !isNaN(n));
          return numValues.length > 0
            ? Math.round((numValues.reduce((a, b) => a + b, 0) / numValues.length) * 100) / 100
            : 0;
        }
        case 'min': {
          const numValues = allChildValues.map(toNumber).filter((n) => !isNaN(n));
          return numValues.length > 0 ? Math.min(...numValues) : 0;
        }
        case 'max': {
          const numValues = allChildValues.map(toNumber).filter((n) => !isNaN(n));
          return numValues.length > 0 ? Math.max(...numValues) : 0;
        }
        default:
          return 0;
      }
    };

    // Helper to get LookupId - checks both object form and separate LookupId field
    const getLookupId = (row: RowData, lookupColumnName: string): string | null => {
      // First check for the separate {ColumnName}LookupId field (SharePoint's format)
      const lookupIdField = `${lookupColumnName}LookupId`;
      if (row[lookupIdField] !== undefined && row[lookupIdField] !== null) {
        return String(row[lookupIdField]);
      }

      // Fall back to checking if the value is an object with LookupId property
      const value = row[lookupColumnName];
      if (value && typeof value === 'object' && 'LookupId' in value) {
        return String((value as { LookupId: unknown }).LookupId);
      }

      return null;
    };

    // For aggregate views with lookup relationships, perform JOINs
    let dataToProcess = rawData;

    if (view.mode === 'aggregate' && lookupRelationships.length > 0) {
      // Separate rows by source list
      const rowsByList = new Map<string, RowData[]>();
      for (const row of rawData) {
        const listId = row._sourceListId;
        if (!rowsByList.has(listId)) {
          rowsByList.set(listId, []);
        }
        rowsByList.get(listId)!.push(row);
      }

      // Determine parent and child lists based on relationships
      // Parent list = the one being looked up TO
      // Child list = the one with the lookup column
      const childListIds = new Set(lookupRelationships.map(r => r.childListId));
      const parentListIds = new Set(lookupRelationships.map(r => r.parentListId));

      // Find the primary parent list (lookup target that's not also a child)
      const primaryParents = [...parentListIds].filter(id => !childListIds.has(id));

      if (primaryParents.length > 0) {
        const parentListId = primaryParents[0];
        const parentRows = rowsByList.get(parentListId) || [];

        // Find relationships where this is the parent
        const childRelations = lookupRelationships.filter(r => r.parentListId === parentListId);

        // Create joined rows: for each parent, collect all related child rows
        const joinedRows: RowData[] = [];

        for (const parentRow of parentRows) {
          // Start with parent row data
          const joinedRow: RowData = { ...parentRow };

          // For each child relationship, find matching child rows
          for (const rel of childRelations) {
            const childRows = rowsByList.get(rel.childListId) || [];

            const matchingChildren = childRows.filter(childRow => {
              const lookupId = getLookupId(childRow, rel.lookupColumnName);
              return lookupId === parentRow._itemId;
            });

            // Store matching child rows for aggregation
            joinedRow[`_childRows_${rel.childListId}`] = matchingChildren;
          }

          joinedRows.push(joinedRow);
        }

        dataToProcess = joinedRows;
      }
    }

    // Apply filters (to parent rows in joined case)
    let filteredData = dataToProcess;
    if (view.filters && view.filters.length > 0) {
      filteredData = dataToProcess.filter((row) =>
        view.filters!.every((filter) => applyFilter(row, filter))
      );
    }

    if (view.mode === 'aggregate') {
      const groupBy = view.groupBy || [];

      // Determine which columns come from parent vs child lists
      const childListIds = new Set(lookupRelationships.map(r => r.childListId));

      if (groupBy.length === 0) {
        // No grouping - single aggregate row
        const result: AggregateResult = {};
        for (const col of view.columns) {
          if (col.aggregation) {
            // Check if this column is from a child list (needs special handling)
            if (childListIds.has(col.sourceListId)) {
              result[col.internalName] = computeJoinedAggregation(filteredData, col, col.sourceListId);
            } else {
              result[col.internalName] = computeAggregation(filteredData, col);
            }
          } else {
            result[col.internalName] = '';
          }
        }
        return [result];
      }

      // Group data by groupBy columns
      const groups = new Map<string, RowData[]>();
      for (const row of filteredData) {
        const key = groupBy.map((col) => String(row[col] ?? '')).join('|||');
        if (!groups.has(key)) {
          groups.set(key, []);
        }
        groups.get(key)!.push(row);
      }

      // Compute aggregations per group
      const results: AggregateResult[] = [];
      for (const [, groupRows] of groups) {
        const result: AggregateResult = {};

        for (const col of view.columns) {
          if (groupBy.includes(col.internalName)) {
            // This is a group-by column - use the value from the first row
            result[col.internalName] = groupRows[0][col.internalName] as string | number;
          } else if (col.aggregation) {
            // Check if this column is from a child list (needs special handling)
            if (childListIds.has(col.sourceListId)) {
              result[col.internalName] = computeJoinedAggregation(groupRows, col, col.sourceListId);
            } else {
              result[col.internalName] = computeAggregation(groupRows, col);
            }
          } else {
            // Column without aggregation in aggregate mode - use first value
            result[col.internalName] = groupRows[0][col.internalName] as string | number;
          }
        }

        results.push(result);
      }

      // Apply user-defined sorting to aggregate results
      return applySorting(results, view.sorting);
    }

    // Union mode - apply sorting
    return applySorting(filteredData, view.sorting);
  }, [view, rawData, lookupRelationships]);

  // Generate DataGrid column definitions
  const columnDefs = useMemo((): TableColumnDefinition<Record<string, unknown>>[] => {
    if (!view) return [];

    return view.columns.map((col) => {
      const isGroupByCol = view.mode === 'aggregate' && view.groupBy?.includes(col.internalName);
      const headerLabel = col.displayName + (col.aggregation && !isGroupByCol ? ` (${col.aggregation})` : '');

      return createTableColumn<Record<string, unknown>>({
        columnId: col.internalName,
        compare: (a, b) => {
          const aVal = String(a[col.internalName] ?? '');
          const bVal = String(b[col.internalName] ?? '');
          return aVal.localeCompare(bVal);
        },
        renderHeaderCell: () => headerLabel,
        renderCell: (item) => (
          <TableCellLayout truncate>
            {formatCellValue(item[col.internalName])}
          </TableCellLayout>
        ),
      });
    });
  }, [view]);

  if (!view) {
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
            <Link to="/app/views" className={styles.breadcrumbLink}>
              Views
            </Link>
          </BreadcrumbItem>
          <BreadcrumbDivider />
          <BreadcrumbItem>
            <Text weight="semibold">Not Found</Text>
          </BreadcrumbItem>
        </Breadcrumb>
        <MessageBar intent="error" className={styles.messageBar}>
          <MessageBarBody>View not found</MessageBarBody>
        </MessageBar>
        <div className={styles.backLink}>
          <Button appearance="subtle" icon={<ArrowLeftRegular />} onClick={() => navigate('/app/views')}>
            Back to Views
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
          <Link to="/app/views" className={styles.breadcrumbLink}>
            Views
          </Link>
        </BreadcrumbItem>
        <BreadcrumbDivider />
        <BreadcrumbItem>
          <Text weight="semibold">{view.name}</Text>
        </BreadcrumbItem>
      </Breadcrumb>

      <div className={styles.content}>
        {/* Header */}
        <div className={styles.header}>
          <div>
            <div className={styles.titleRow}>
              <Title2 as="h1">{view.name}</Title2>
              <Badge
                appearance="filled"
                color={view.mode === 'aggregate' ? 'important' : 'brand'}
              >
                {view.mode === 'aggregate' ? 'Aggregate' : 'Union'}
              </Badge>
            </div>
            {view.description && (
              <Text className={styles.description}>{view.description}</Text>
            )}
            <div className={styles.meta}>
              <span>
                {view.sources.length} source{view.sources.length !== 1 ? 's' : ''}
              </span>
              <span>
                {view.columns.length} column{view.columns.length !== 1 ? 's' : ''}
              </span>
              {view.filters && view.filters.length > 0 && (
                <span>
                  {view.filters.length} filter{view.filters.length !== 1 ? 's' : ''}
                </span>
              )}
            </div>
          </div>
          <Button
            appearance="outline"
            size="small"
            icon={<SettingsRegular />}
            onClick={() => navigate(`/app/views/${view.id}/edit`)}
          >
            Edit View
          </Button>
        </div>

        {/* Loading State */}
        {loading && (
          <Card className={styles.loadingCard}>
            <div className={styles.loadingContent}>
              <Spinner size="large" />
              <Text className={styles.description}>Loading data from sources...</Text>
            </div>
          </Card>
        )}

        {/* Error State */}
        {error && !loading && (
          <MessageBar intent="error" className={styles.messageBar}>
            <MessageBarBody>
              <WarningRegular style={{ marginRight: '8px' }} />
              {error}
            </MessageBarBody>
          </MessageBar>
        )}

        {/* No Data */}
        {!loading && !error && processedData.length === 0 && (
          <Card className={styles.emptyCard}>
            <div className={styles.emptyContent}>
              <DatabaseRegular className={styles.emptyIcon} />
              <Text>No data found</Text>
              <Text size={200}>
                {view.filters && view.filters.length > 0
                  ? 'Try adjusting the filters'
                  : 'The source lists may be empty'}
              </Text>
            </div>
          </Card>
        )}

        {/* DataGrid Data Table */}
        {!loading && !error && processedData.length > 0 && (
          <div>
            <DataGrid
              items={processedData}
              columns={columnDefs}
              sortable
              resizableColumns
              className={styles.dataGrid}
            >
              <DataGridHeader>
                <DataGridRow>
                  {({ renderHeaderCell }) => (
                    <DataGridHeaderCell>{renderHeaderCell()}</DataGridHeaderCell>
                  )}
                </DataGridRow>
              </DataGridHeader>
              <DataGridBody<Record<string, unknown>>>
                {({ item, rowId }) => (
                  <DataGridRow<Record<string, unknown>> key={rowId}>
                    {({ renderCell }) => (
                      <DataGridCell>{renderCell(item)}</DataGridCell>
                    )}
                  </DataGridRow>
                )}
              </DataGridBody>
            </DataGrid>
            <Text className={styles.rowCount}>
              {processedData.length} row{processedData.length !== 1 ? 's' : ''} total
            </Text>
          </div>
        )}

        {/* Back Button */}
        <div className={styles.footer}>
          <Button appearance="subtle" icon={<ArrowLeftRegular />} onClick={() => navigate('/app/views')}>
            Back to Views
          </Button>
        </div>
      </div>
    </div>
  );
}

function formatCellValue(value: unknown): string {
  if (value === null || value === undefined) {
    return '-';
  }
  if (typeof value === 'boolean') {
    return value ? 'Yes' : 'No';
  }
  if (typeof value === 'object') {
    // Handle complex types like dates or lookups
    if ('DisplayValue' in (value as Record<string, unknown>)) {
      return String((value as Record<string, unknown>).DisplayValue);
    }
    return JSON.stringify(value);
  }
  return String(value);
}

export default ViewDisplayPage;
