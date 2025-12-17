import { useMsal } from '@azure/msal-react';
import { useEffect, useState } from 'react';
import { Link, useParams } from 'react-router-dom';
import {
  getListById,
  getListItems,
  type GraphList,
  type GraphListColumn,
  type GraphListItem,
} from '../auth/graphClient';

function ListViewPage() {
  const { siteId, listId } = useParams<{ siteId: string; listId: string }>();
  const { instance, accounts } = useMsal();
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
        setColumns(listData.columns);
        setItems(listData.items);
      } catch (err) {
        setError(err instanceof Error ? err.message : 'Failed to load list');
      } finally {
        setLoading(false);
      }
    };

    loadData();
  }, [instance, account, siteId, listId]);

  const formatCellValue = (value: unknown): string => {
    if (value === null || value === undefined) return '';
    if (typeof value === 'boolean') return value ? 'Yes' : 'No';
    if (typeof value === 'object') {
      // Handle lookup fields and other complex types
      if ('LookupValue' in (value as Record<string, unknown>)) {
        return String((value as Record<string, unknown>).LookupValue || '');
      }
      return JSON.stringify(value);
    }
    return String(value);
  };

  return (
    <div className="p-8">
      {/* Breadcrumb */}
      <div className="text-sm breadcrumbs mb-6">
        <ul>
          <li>
            <Link to="/app">Home</Link>
          </li>
          <li>
            <Link to="/app/data">Data</Link>
          </li>
          <li>{list?.displayName || 'List'}</li>
        </ul>
      </div>

      <div>
        <div className="flex items-start justify-between mb-6">
          <div>
            <h1 className="text-2xl font-bold mb-1">{list?.displayName || 'List'}</h1>
            {!loading && items.length > 0 && (
              <p className="text-base-content/60">
                Showing {items.length} item{items.length !== 1 ? 's' : ''}
                {items.length === 1000 && ' (limited to 1000)'}
              </p>
            )}
          </div>
        </div>

        {/* Loading State */}
        {loading && (
          <div className="card bg-base-200">
            <div className="card-body items-center text-center py-12">
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
          <div className="card bg-base-200">
            <div className="card-body items-center text-center py-12">
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
                  d="M3.375 19.5h17.25m-17.25 0a1.125 1.125 0 0 1-1.125-1.125M3.375 19.5h7.5c.621 0 1.125-.504 1.125-1.125m-9.75 0V5.625m0 12.75v-1.5c0-.621.504-1.125 1.125-1.125m18.375 2.625V5.625m0 12.75c0 .621-.504 1.125-1.125 1.125m1.125-1.125v-1.5c0-.621-.504-1.125-1.125-1.125m0 3.75h-7.5A1.125 1.125 0 0 1 12 18.375m9.75-12.75c0-.621-.504-1.125-1.125-1.125H3.375c-.621 0-1.125.504-1.125 1.125m19.5 0v1.5c0 .621-.504 1.125-1.125 1.125M2.25 5.625v1.5c0 .621.504 1.125 1.125 1.125m0 0h17.25m-17.25 0h7.5c.621 0 1.125.504 1.125 1.125M3.375 8.25c-.621 0-1.125.504-1.125 1.125v1.5c0 .621.504 1.125 1.125 1.125m17.25-3.75h-7.5c-.621 0-1.125.504-1.125 1.125m8.625-1.125c.621 0 1.125.504 1.125 1.125v1.5c0 .621-.504 1.125-1.125 1.125"
                />
              </svg>
              <p className="text-base-content/60">No items in this list</p>
            </div>
          </div>
        )}

        {/* Spreadsheet Table */}
        {!loading && !error && items.length > 0 && (
          <div className="card bg-base-200 overflow-hidden">
            <div className="overflow-x-auto">
              <table className="table table-pin-rows table-pin-cols">
                <thead>
                  <tr>
                    <th className="bg-base-300 z-10">#</th>
                    {columns.map((col) => (
                      <th key={col.id} className="bg-base-300">
                        {col.displayName}
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {items.map((item, index) => (
                    <tr key={item.id} className="hover:bg-base-300/50">
                      <th className="bg-base-200 font-normal text-base-content/60">
                        {index + 1}
                      </th>
                      {columns.map((col) => (
                        <td key={col.id} className="max-w-xs truncate">
                          {formatCellValue(item.fields[col.name])}
                        </td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
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
