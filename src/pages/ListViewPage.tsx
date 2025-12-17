import { useMsal } from '@azure/msal-react';
import { useEffect, useState, useMemo } from 'react';
import { Link, useParams } from 'react-router-dom';
import { AgGridReact } from 'ag-grid-react';
import { ModuleRegistry, AllCommunityModule, themeQuartz, colorSchemeDark } from 'ag-grid-community';
import type { ColDef, ValueFormatterParams } from 'ag-grid-community';
import { useTheme } from '../contexts/ThemeContext';
import {
  getListById,
  getListItems,
  type GraphList,
  type GraphListColumn,
  type GraphListItem,
} from '../auth/graphClient';
import { getDefaultViewColumnOrder } from '../services/sharepoint';

// Register AG Grid modules
ModuleRegistry.registerModules([AllCommunityModule]);

interface RowData {
  [key: string]: unknown;
}

function ListViewPage() {
  const { siteId, listId } = useParams<{ siteId: string; listId: string }>();
  const { instance, accounts } = useMsal();
  const { theme } = useTheme();
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
  const formatCellValue = (value: unknown): string => {
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
  };

  // Convert items to row data
  const rowData = useMemo((): RowData[] => {
    return items.map((item) => ({
      _id: item.id,
      ...item.fields,
    }));
  }, [items]);

  // Generate AG Grid column definitions
  const columnDefs = useMemo((): ColDef[] => {
    return columns.map((col) => ({
      headerName: col.displayName,
      field: col.name,
      sortable: true,
      filter: true,
      resizable: true,
      valueFormatter: (params: ValueFormatterParams) => formatCellValue(params.value),
    }));
  }, [columns]);

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

  return (
    <div className="p-8 h-full flex flex-col">
      {/* Breadcrumb */}
      <div className="text-sm breadcrumbs mb-6">
        <ul>
          <li>
            <Link to="/app">Home</Link>
          </li>
          <li>
            <Link to="/app/lists">Lists</Link>
          </li>
          <li>{list?.displayName || 'List'}</li>
        </ul>
      </div>

      <div className="flex-1 flex flex-col min-h-0">
        {/* Header */}
        <div className="flex items-start justify-between mb-6">
          <div>
            <h1 className="text-2xl font-bold mb-1">{list?.displayName || 'List'}</h1>
            {!loading && items.length > 0 && (
              <p className="text-base-content/60">
                {items.length} item{items.length !== 1 ? 's' : ''}
                {items.length === 1000 && ' (limited to 1000)'}
              </p>
            )}
          </div>
        </div>

        {/* Loading State */}
        {loading && (
          <div className="card bg-base-200 flex-1">
            <div className="card-body items-center justify-center">
              <span className="loading loading-spinner loading-lg text-primary" />
              <p className="text-base-content/60 mt-4">Loading list data...</p>
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

        {/* Empty State */}
        {!loading && !error && items.length === 0 && (
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
              <p className="text-base-content/60">No items in this list</p>
            </div>
          </div>
        )}

        {/* AG Grid Data Table */}
        {!loading && !error && items.length > 0 && (
          <div>
            <AgGridReact
              theme={gridTheme}
              rowData={rowData}
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
              {items.length} row{items.length !== 1 ? 's' : ''} total
            </p>
          </div>
        )}

        {/* Back Button */}
        <div className="mt-8 pt-6 border-t border-base-300">
          <Link to="/app" className="btn btn-ghost">
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
            Back
          </Link>
        </div>
      </div>
    </div>
  );
}

export default ListViewPage;
