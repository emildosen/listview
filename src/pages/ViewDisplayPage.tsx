import { useState, useEffect, useMemo } from 'react';
import { useParams, Link } from 'react-router-dom';
import { useMsal } from '@azure/msal-react';
import { AgGridReact } from 'ag-grid-react';
import { ModuleRegistry, AllCommunityModule, themeQuartz, colorSchemeDark } from 'ag-grid-community';
import type { ColDef, ValueFormatterParams } from 'ag-grid-community';
import { useSettings } from '../contexts/SettingsContext';
import { useTheme } from '../contexts/ThemeContext';
import { getListItems, type GraphListItem } from '../auth/graphClient';

// Register AG Grid modules
ModuleRegistry.registerModules([AllCommunityModule]);
import type { ViewDefinition, ViewFilter, FilterOperator, ViewSorting } from '../types/view';

interface RowData {
  _sourceListId: string;
  _sourceListName: string;
  [key: string]: unknown;
}

interface AggregateResult {
  [key: string]: number | string;
}

function ViewDisplayPage() {
  const { viewId } = useParams<{ viewId: string }>();
  const { instance, accounts } = useMsal();
  const { views } = useSettings();
  const { theme } = useTheme();
  const account = accounts[0];

  const [view, setView] = useState<ViewDefinition | null>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [rawData, setRawData] = useState<RowData[]>([]);

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

        for (const source of view.sources) {
          const result = await getListItems(instance, account, source.siteId, source.listId);

          // Map items to row data
          const rows = result.items.map((item: GraphListItem) => ({
            _sourceListId: source.listId,
            _sourceListName: source.listName,
            ...item.fields,
          }));

          allRows.push(...rows);
        }

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

  // Process data based on view mode
  const processedData = useMemo(() => {
    if (!view || rawData.length === 0) return [];

    // Apply filters
    let filteredData = rawData;
    if (view.filters && view.filters.length > 0) {
      filteredData = rawData.filter((row) =>
        view.filters!.every((filter) => applyFilter(row, filter))
      );
    }

    if (view.mode === 'aggregate') {
      const groupBy = view.groupBy || [];

      if (groupBy.length === 0) {
        // No grouping - single aggregate row
        const result: AggregateResult = {};
        for (const col of view.columns) {
          if (col.aggregation) {
            result[col.internalName] = computeAggregation(filteredData, col);
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
            // This is an aggregated column
            result[col.internalName] = computeAggregation(groupRows, col);
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
  }, [view, rawData]);

  // Generate AG Grid column definitions
  const columnDefs = useMemo((): ColDef[] => {
    if (!view) return [];

    const cols: ColDef[] = [];

    // Add view columns
    for (const col of view.columns) {
      const isGroupByCol = view.mode === 'aggregate' && view.groupBy?.includes(col.internalName);
      const colDef: ColDef = {
        headerName: col.displayName + (col.aggregation && !isGroupByCol ? ` (${col.aggregation})` : ''),
        field: col.internalName,
        sortable: true,
        filter: true,
        resizable: true,
        valueFormatter: (params: ValueFormatterParams) => formatCellValue(params.value),
      };
      cols.push(colDef);
    }

    return cols;
  }, [view]);

  // AG Grid default column settings
  const defaultColDef = useMemo((): ColDef => ({
    flex: 1,
    minWidth: 100,
    resizable: true,
  }), []);

  // AG Grid theme based on current app theme
  const gridTheme = useMemo(() => {
    return theme === 'dark' ? themeQuartz.withPart(colorSchemeDark) : themeQuartz;
  }, [theme]);

  if (!view) {
    return (
      <div className="p-8">
        <div className="text-sm breadcrumbs mb-6">
          <ul>
            <li>
              <Link to="/app">Home</Link>
            </li>
            <li>
              <Link to="/app/views">Views</Link>
            </li>
            <li>Not Found</li>
          </ul>
        </div>
        <div className="alert alert-error">
          <span>View not found</span>
        </div>
        <div className="mt-4">
          <Link to="/app/views" className="btn btn-ghost">
            Back to Views
          </Link>
        </div>
      </div>
    );
  }

  return (
    <div className="p-8 h-full flex flex-col">
      {/* Breadcrumb */}
      <div className="text-sm breadcrumbs mb-6">
        <ul>
          <li>
            <Link to="/app">Home</Link>
          </li>
          <li>
            <Link to="/app/views">Views</Link>
          </li>
          <li>{view.name}</li>
        </ul>
      </div>

      <div className="flex-1 flex flex-col min-h-0">
        {/* Header */}
        <div className="flex items-start justify-between mb-6">
          <div>
            <div className="flex items-center gap-3">
              <h1 className="text-2xl font-bold">{view.name}</h1>
              <span
                className={`badge ${
                  view.mode === 'aggregate' ? 'badge-secondary' : 'badge-primary'
                }`}
              >
                {view.mode === 'aggregate' ? 'Aggregate' : 'Union'}
              </span>
            </div>
            {view.description && (
              <p className="text-base-content/60 mt-1">{view.description}</p>
            )}
            <div className="flex items-center gap-4 mt-2 text-sm text-base-content/60">
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
          <Link to={`/app/views/${view.id}/edit`} className="btn btn-outline btn-sm">
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
                d="M9.594 3.94c.09-.542.56-.94 1.11-.94h2.593c.55 0 1.02.398 1.11.94l.213 1.281c.063.374.313.686.645.87.074.04.147.083.22.127.325.196.72.257 1.075.124l1.217-.456a1.125 1.125 0 0 1 1.37.49l1.296 2.247a1.125 1.125 0 0 1-.26 1.431l-1.003.827c-.293.241-.438.613-.43.992a7.723 7.723 0 0 1 0 .255c-.008.378.137.75.43.991l1.004.827c.424.35.534.955.26 1.43l-1.298 2.247a1.125 1.125 0 0 1-1.369.491l-1.217-.456c-.355-.133-.75-.072-1.076.124a6.47 6.47 0 0 1-.22.128c-.331.183-.581.495-.644.869l-.213 1.281c-.09.543-.56.94-1.11.94h-2.594c-.55 0-1.019-.398-1.11-.94l-.213-1.281c-.062-.374-.312-.686-.644-.87a6.52 6.52 0 0 1-.22-.127c-.325-.196-.72-.257-1.076-.124l-1.217.456a1.125 1.125 0 0 1-1.369-.49l-1.297-2.247a1.125 1.125 0 0 1 .26-1.431l1.004-.827c.292-.24.437-.613.43-.991a6.932 6.932 0 0 1 0-.255c.007-.38-.138-.751-.43-.992l-1.004-.827a1.125 1.125 0 0 1-.26-1.43l1.297-2.247a1.125 1.125 0 0 1 1.37-.491l1.216.456c.356.133.751.072 1.076-.124.072-.044.146-.086.22-.128.332-.183.582-.495.644-.869l.214-1.28Z"
              />
              <path
                strokeLinecap="round"
                strokeLinejoin="round"
                d="M15 12a3 3 0 1 1-6 0 3 3 0 0 1 6 0Z"
              />
            </svg>
            Edit View
          </Link>
        </div>

        {/* Loading State */}
        {loading && (
          <div className="card bg-base-200 flex-1">
            <div className="card-body items-center justify-center">
              <span className="loading loading-spinner loading-lg text-primary" />
              <p className="text-base-content/60 mt-4">Loading data from sources...</p>
            </div>
          </div>
        )}

        {/* Error State */}
        {error && !loading && (
          <div className="alert alert-error mb-4">
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
                d="M12 9v3.75m9-.75a9 9 0 1 1-18 0 9 9 0 0 1 18 0Zm-9 3.75h.008v.008H12v-.008Z"
              />
            </svg>
            <span>{error}</span>
          </div>
        )}

        {/* No Data */}
        {!loading && !error && processedData.length === 0 && (
          <div className="card bg-base-200 flex-1">
            <div className="card-body items-center justify-center">
              <svg
                xmlns="http://www.w3.org/2000/svg"
                fill="none"
                viewBox="0 0 24 24"
                strokeWidth={1.5}
                stroke="currentColor"
                className="w-12 h-12 text-base-content/30 mb-4"
              >
                <path
                  strokeLinecap="round"
                  strokeLinejoin="round"
                  d="M20.25 6.375c0 2.278-3.694 4.125-8.25 4.125S3.75 8.653 3.75 6.375m16.5 0c0-2.278-3.694-4.125-8.25-4.125S3.75 4.097 3.75 6.375m16.5 0v11.25c0 2.278-3.694 4.125-8.25 4.125s-8.25-1.847-8.25-4.125V6.375m16.5 0v3.75m-16.5-3.75v3.75m16.5 0v3.75C20.25 16.153 16.556 18 12 18s-8.25-1.847-8.25-4.125v-3.75m16.5 0c0 2.278-3.694 4.125-8.25 4.125s-8.25-1.847-8.25-4.125"
                />
              </svg>
              <p className="text-base-content/60">No data found</p>
              <p className="text-sm text-base-content/40">
                {view.filters && view.filters.length > 0
                  ? 'Try adjusting the filters'
                  : 'The source lists may be empty'}
              </p>
            </div>
          </div>
        )}

        {/* AG Grid Data Table */}
        {!loading && !error && processedData.length > 0 && (
          <div>
            <AgGridReact
              theme={gridTheme}
              rowData={processedData}
              columnDefs={columnDefs}
              defaultColDef={defaultColDef}
              domLayout="autoHeight"
              animateRows={true}
              pagination={true}
              paginationPageSize={50}
              paginationPageSizeSelector={[25, 50, 100, 200]}
              suppressMovableColumns={false}
              enableCellTextSelection={true}
            />
            <p className="text-sm text-base-content/60 mt-2">
              {processedData.length} row{processedData.length !== 1 ? 's' : ''} total
            </p>
          </div>
        )}

        {/* Back Button */}
        <div className="mt-8 pt-6 border-t border-base-300">
          <Link to="/app/views" className="btn btn-ghost">
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
                d="M10.5 19.5 3 12m0 0 7.5-7.5M3 12h18"
              />
            </svg>
            Back to Views
          </Link>
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
