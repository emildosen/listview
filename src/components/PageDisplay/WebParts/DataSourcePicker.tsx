import { useState, useEffect, useCallback } from 'react';
import {
  makeStyles,
  tokens,
  Dropdown,
  Option,
  Field,
  Spinner,
  Text,
} from '@fluentui/react-components';
import { useMsal } from '@azure/msal-react';
import type { WebPartDataSource } from '../../../types/page';
import { getAllSites, getSiteLists, type GraphSite, type GraphList } from '../../../auth/graphClient';
import { SYSTEM_LIST_NAMES } from '../../../services/sharepoint';

const useStyles = makeStyles({
  container: {
    display: 'flex',
    flexDirection: 'column',
    gap: '12px',
  },
  loadingRow: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase200,
  },
  listInfo: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground3,
    marginTop: '4px',
  },
});

interface DataSourcePickerProps {
  value: WebPartDataSource | undefined;
  onChange: (source: WebPartDataSource) => void;
}

export default function DataSourcePicker({ value, onChange }: DataSourcePickerProps) {
  const styles = useStyles();
  const { instance, accounts } = useMsal();
  const account = accounts[0];

  const [sites, setSites] = useState<GraphSite[]>([]);
  const [lists, setLists] = useState<GraphList[]>([]);
  const [loadingSites, setLoadingSites] = useState(false);
  const [loadingLists, setLoadingLists] = useState(false);
  const [error, setError] = useState<string | null>(null);

  // Load sites on mount
  useEffect(() => {
    async function loadSites() {
      if (!account) return;

      setLoadingSites(true);
      setError(null);
      try {
        const fetchedSites = await getAllSites(instance, account);
        setSites(fetchedSites);
      } catch (err) {
        console.error('Failed to load sites:', err);
        setError('Failed to load sites');
      } finally {
        setLoadingSites(false);
      }
    }

    loadSites();
  }, [instance, account]);

  // Load lists when site changes
  useEffect(() => {
    async function loadLists() {
      if (!value?.siteId || !account) {
        setLists([]);
        return;
      }

      setLoadingLists(true);
      try {
        const fetchedLists = await getSiteLists(instance, account, value.siteId);
        // Filter out system lists
        const filteredLists = fetchedLists.filter(
          (list) => !SYSTEM_LIST_NAMES.includes(list.name as typeof SYSTEM_LIST_NAMES[number])
        );
        setLists(filteredLists);
      } catch (err) {
        console.error('Failed to load lists:', err);
        setLists([]);
      } finally {
        setLoadingLists(false);
      }
    }

    loadLists();
  }, [value?.siteId, instance, account]);

  const handleSiteChange = useCallback(
    (_: unknown, data: { optionValue?: string; optionText?: string }) => {
      if (!data.optionValue) return;

      const selectedSite = sites.find((s) => s.id === data.optionValue);
      if (selectedSite) {
        onChange({
          siteId: selectedSite.id,
          siteUrl: selectedSite.webUrl,
          listId: '',
          listName: '',
        });
      }
    },
    [sites, onChange]
  );

  const handleListChange = useCallback(
    (_: unknown, data: { optionValue?: string; optionText?: string }) => {
      if (!data.optionValue || !value) return;

      const selectedList = lists.find((l) => l.id === data.optionValue);
      if (selectedList) {
        onChange({
          ...value,
          listId: selectedList.id,
          listName: selectedList.displayName,
        });
      }
    },
    [lists, value, onChange]
  );

  const selectedSite = sites.find((s) => s.id === value?.siteId);
  const selectedList = lists.find((l) => l.id === value?.listId);

  if (error) {
    return <Text style={{ color: tokens.colorPaletteRedForeground1 }}>{error}</Text>;
  }

  return (
    <div className={styles.container}>
      {/* Site Selector */}
      <Field label="Site">
        {loadingSites ? (
          <div className={styles.loadingRow}>
            <Spinner size="tiny" />
            <span>Loading sites...</span>
          </div>
        ) : (
          <Dropdown
            placeholder="Select a site"
            value={selectedSite?.displayName || ''}
            selectedOptions={value?.siteId ? [value.siteId] : []}
            onOptionSelect={handleSiteChange}
          >
            {sites.map((site) => (
              <Option key={site.id} value={site.id}>
                {site.displayName}
              </Option>
            ))}
          </Dropdown>
        )}
      </Field>

      {/* List Selector */}
      {value?.siteId && (
        <Field label="List">
          {loadingLists ? (
            <div className={styles.loadingRow}>
              <Spinner size="tiny" />
              <span>Loading lists...</span>
            </div>
          ) : (
            <>
              <Dropdown
                placeholder="Select a list"
                value={selectedList?.displayName || ''}
                selectedOptions={value?.listId ? [value.listId] : []}
                onOptionSelect={handleListChange}
                disabled={lists.length === 0}
              >
                {lists.map((list) => (
                  <Option key={list.id} value={list.id}>
                    {list.displayName}
                  </Option>
                ))}
              </Dropdown>
              {lists.length === 0 && !loadingLists && (
                <Text className={styles.listInfo}>No lists found in this site</Text>
              )}
            </>
          )}
        </Field>
      )}

      {/* Selection Summary */}
      {value?.listId && selectedList && (
        <Text className={styles.listInfo}>
          Selected: {selectedSite?.displayName} / {selectedList.displayName}
        </Text>
      )}
    </div>
  );
}
