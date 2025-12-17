import { useMsal } from '@azure/msal-react';
import { useEffect, useState, useCallback } from 'react';
import { Link } from 'react-router-dom';
import { useSettings } from '../contexts/SettingsContext';
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
  listId: string;
  listName: string;
}

export const ENABLED_LISTS_KEY = 'EnabledLists';

function DataPage() {
  const { instance, accounts } = useMsal();
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
    <div className="p-8">
      {/* Breadcrumb */}
      <div className="text-sm breadcrumbs mb-6">
        <ul>
          <li>
            <Link to="/app">Home</Link>
          </li>
          <li>Lists</li>
        </ul>
      </div>

      <div className="max-w-4xl">
        <div className="flex items-start justify-between mb-6">
          <div>
            <h1 className="text-2xl font-bold mb-1">Manage Lists</h1>
            <p className="text-base-content/60">
              Select which SharePoint lists to enable.
            </p>
          </div>
          {selectedLists.size > 0 && (
            <div className="badge badge-primary badge-lg">
              {selectedLists.size} list{selectedLists.size !== 1 ? 's' : ''} selected
            </div>
          )}
        </div>

        {/* Search Bar */}
        {!loading && lists.length > 0 && (
          <div className="mb-4">
            <label className="input input-bordered flex items-center gap-2">
              <svg
                xmlns="http://www.w3.org/2000/svg"
                fill="none"
                viewBox="0 0 24 24"
                strokeWidth={1.5}
                stroke="currentColor"
                className="w-5 h-5 text-base-content/50"
              >
                <path
                  strokeLinecap="round"
                  strokeLinejoin="round"
                  d="m21 21-5.197-5.197m0 0A7.5 7.5 0 1 0 5.196 5.196a7.5 7.5 0 0 0 10.607 10.607Z"
                />
              </svg>
              <input
                type="text"
                placeholder="Search lists..."
                className="grow bg-transparent border-none outline-none"
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
              />
              {searchQuery && (
                <button
                  onClick={() => setSearchQuery('')}
                  className="text-base-content/50 hover:text-base-content"
                >
                  <svg
                    xmlns="http://www.w3.org/2000/svg"
                    fill="none"
                    viewBox="0 0 24 24"
                    strokeWidth={1.5}
                    stroke="currentColor"
                    className="w-5 h-5"
                  >
                    <path strokeLinecap="round" strokeLinejoin="round" d="M6 18 18 6M6 6l12 12" />
                  </svg>
                </button>
              )}
            </label>
          </div>
        )}

        {/* Loading State */}
        {loading && (
          <div className="card bg-base-200">
            <div className="card-body items-center text-center py-12">
              <span className="loading loading-spinner loading-lg text-primary" />
              <p className="text-base-content/60 mt-4">Loading sites and lists...</p>
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

        {/* No Lists */}
        {!loading && !error && lists.length === 0 && (
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
                  d="M20.25 6.375c0 2.278-3.694 4.125-8.25 4.125S3.75 8.653 3.75 6.375m16.5 0c0-2.278-3.694-4.125-8.25-4.125S3.75 4.097 3.75 6.375m16.5 0v11.25c0 2.278-3.694 4.125-8.25 4.125s-8.25-1.847-8.25-4.125V6.375m16.5 0v3.75m-16.5-3.75v3.75m16.5 0v3.75C20.25 16.153 16.556 18 12 18s-8.25-1.847-8.25-4.125v-3.75m16.5 0c0 2.278-3.694 4.125-8.25 4.125s-8.25-1.847-8.25-4.125"
                />
              </svg>
              <p className="text-base-content/60">No lists found</p>
              <p className="text-sm text-base-content/40">
                No SharePoint lists available
              </p>
            </div>
          </div>
        )}

        {/* No Search Results */}
        {!loading && lists.length > 0 && filteredLists.length === 0 && searchQuery && (
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
                  d="m21 21-5.197-5.197m0 0A7.5 7.5 0 1 0 5.196 5.196a7.5 7.5 0 0 0 10.607 10.607Z"
                />
              </svg>
              <p className="text-base-content/60">No lists match "{searchQuery}"</p>
              <button
                onClick={() => setSearchQuery('')}
                className="btn btn-ghost btn-sm mt-2"
              >
                Clear search
              </button>
            </div>
          </div>
        )}

        {/* Lists Table */}
        {!loading && filteredLists.length > 0 && (
          <div className="card bg-base-200">
            <div className="overflow-x-auto">
              <table className="table">
                <thead>
                  <tr>
                    <th>
                      <label>
                        <input
                          type="checkbox"
                          className="checkbox checkbox-sm"
                          checked={selectedLists.size === lists.length && lists.length > 0}
                          onChange={handleSelectAll}
                        />
                      </label>
                    </th>
                    <th>List Name</th>
                    <th>Site</th>
                  </tr>
                </thead>
                <tbody>
                  {filteredLists.map((list) => {
                    const key = getListKey(list.siteId, list.listId);
                    const isSelected = selectedLists.has(key);
                    return (
                      <tr
                        key={key}
                        className={isSelected ? 'bg-primary/5' : 'hover:bg-base-300/50'}
                      >
                        <th>
                          <label>
                            <input
                              type="checkbox"
                              className="checkbox checkbox-sm checkbox-primary"
                              checked={isSelected}
                              onChange={() => handleToggleList(list.siteId, list.listId)}
                            />
                          </label>
                        </th>
                        <td className="font-medium">{list.listName}</td>
                        <td className="text-base-content/60">{list.siteName}</td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {/* Action Buttons */}
        <div className="flex items-center justify-between mt-8 pt-6 border-t border-base-300">
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

          <div className="flex items-center gap-3">
            {hasChanges && (
              <span className="text-sm text-warning">Unsaved changes</span>
            )}
            <button
              onClick={handleCancel}
              disabled={!hasChanges || saving}
              className="btn btn-ghost"
            >
              Cancel
            </button>
            <button
              onClick={handleSave}
              disabled={!hasChanges || saving}
              className="btn btn-primary"
            >
              {saving ? (
                <>
                  <span className="loading loading-spinner loading-sm" />
                  Saving...
                </>
              ) : (
                'Save'
              )}
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}

export default DataPage;
