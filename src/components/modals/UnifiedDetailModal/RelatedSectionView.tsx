import { useState, useEffect, useCallback, useRef } from 'react';
import { useMsal } from '@azure/msal-react';
import type { SPFI } from '@pnp/sp';
import {
  makeStyles,
  tokens,
  Text,
  Button,
  Spinner,
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
  mergeClasses,
} from '@fluentui/react-components';
import type { TableColumnDefinition } from '@fluentui/react-components';
import { AddRegular } from '@fluentui/react-icons';
import { getListItems, type GraphListItem } from '../../../auth/graphClient';
import { createListItem, createSPClient } from '../../../services/sharepoint';
import type { RelatedSection } from '../../../types/page';
import { useModalNavigation, type NavigationEntry } from './ModalNavigationContext';
import { useTheme } from '../../../contexts/ThemeContext';
import ItemFormModal from '../ItemFormModal';

const useStyles = makeStyles({
  container: {
    marginTop: '24px',
  },
  header: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    marginBottom: '12px',
  },
  sectionTitle: {
    fontWeight: tokens.fontWeightSemibold,
    fontSize: tokens.fontSizeBase200,
    textTransform: 'uppercase',
    letterSpacing: '0.02em',
    color: tokens.colorNeutralForeground2,
  },
  card: {
    backgroundColor: tokens.colorNeutralBackground2,
    borderRadius: tokens.borderRadiusMedium,
    padding: '16px',
  },
  cardDark: {
    backgroundColor: '#1a1a1a',
    border: '1px solid #333333',
  },
  loadingContainer: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    padding: '24px',
  },
  emptyState: {
    textAlign: 'center',
    padding: '24px',
    color: tokens.colorNeutralForeground3,
  },
  dataGridWrapper: {
    width: '100%',
    '& [role="row"]': {
      display: 'flex',
      width: '100%',
      cursor: 'pointer',
    },
    '& [role="row"]:hover': {
      backgroundColor: tokens.colorNeutralBackground3,
    },
    '& [role="columnheader"], & [role="gridcell"]': {
      overflow: 'hidden',
      flex: '1 1 0',
      minWidth: 0,
    },
  },
  cellText: {
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap',
    maxWidth: '100%',
  },
});

interface RelatedSectionViewProps {
  section: RelatedSection;
  parentItem: GraphListItem;
}

export function RelatedSectionView({ section, parentItem }: RelatedSectionViewProps) {
  const styles = useStyles();
  const { theme } = useTheme();
  const { instance, accounts } = useMsal();
  const account = accounts[0];
  const { navigateToItem } = useModalNavigation();

  // Data state
  const [items, setItems] = useState<GraphListItem[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  // SP client for CRUD
  const spClientRef = useRef<SPFI | null>(null);

  // Form modal state
  const [formModalOpen, setFormModalOpen] = useState(false);
  const [saving, setSaving] = useState(false);

  // Initialize SP client
  useEffect(() => {
    if (!section.source.siteUrl || !account) return;
    createSPClient(instance, account, section.source.siteUrl)
      .then(client => {
        spClientRef.current = client;
      })
      .catch(console.error);
  }, [instance, account, section.source.siteUrl]);

  // Load related items
  const loadItems = useCallback(async () => {
    setLoading(true);
    setError(null);

    try {
      const result = await getListItems(instance, account, section.source.siteId, section.source.listId);

      // Filter items by lookup column match
      const parentId = parentItem.id;
      const filtered = result.items.filter(item => {
        const lookupValue = item.fields[`${section.lookupColumn}LookupId`];
        return String(lookupValue) === parentId;
      });

      // Sort if configured
      if (section.defaultSort) {
        filtered.sort((a, b) => {
          const aVal = a.fields[section.defaultSort!.column];
          const bVal = b.fields[section.defaultSort!.column];
          const comparison = String(aVal ?? '').localeCompare(String(bVal ?? ''));
          return section.defaultSort!.direction === 'desc' ? -comparison : comparison;
        });
      }

      setItems(filtered);
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Failed to load items');
    } finally {
      setLoading(false);
    }
  }, [instance, account, section, parentItem.id]);

  useEffect(() => {
    loadItems();
  }, [loadItems]);

  // Handle row click - navigate to related item
  const handleRowClick = (item: GraphListItem) => {
    const entry: NavigationEntry = {
      listId: section.source.listId,
      siteId: section.source.siteId,
      siteUrl: section.source.siteUrl,
      itemId: item.id,
      listName: section.source.listName,
    };
    navigateToItem(entry);
  };

  // Handle add
  const handleAdd = () => {
    setFormModalOpen(true);
  };

  // Handle form save
  const handleFormSave = async (fields: Record<string, unknown>) => {
    if (!spClientRef.current) return;

    setSaving(true);
    try {
      // Add parent lookup ID
      const saveFields = {
        ...fields,
        [`${section.lookupColumn}Id`]: parseInt(parentItem.id, 10),
      };

      await createListItem(spClientRef.current, section.source.listId, saveFields);

      setFormModalOpen(false);
      await loadItems();
    } finally {
      setSaving(false);
    }
  };

  // Build table columns
  const tableColumns: TableColumnDefinition<GraphListItem>[] = section.displayColumns.map(col =>
    createTableColumn<GraphListItem>({
      columnId: col.internalName,
      renderHeaderCell: () => col.displayName,
      renderCell: (item) => {
        const value = item.fields[col.internalName];
        let displayValue = '-';

        if (value !== null && value !== undefined) {
          if (typeof value === 'object' && 'LookupValue' in value) {
            displayValue = (value as { LookupValue: string }).LookupValue;
          } else if (Array.isArray(value)) {
            displayValue = value
              .map(v => (typeof v === 'object' && 'LookupValue' in v ? v.LookupValue : String(v)))
              .join(', ');
          } else {
            displayValue = String(value);
          }
        }

        return (
          <TableCellLayout>
            <span className={styles.cellText}>{displayValue}</span>
          </TableCellLayout>
        );
      },
    })
  );

  return (
    <div className={styles.container}>
      <div className={styles.header}>
        <Text className={styles.sectionTitle}>{section.title}</Text>
        <Button
          appearance="subtle"
          size="small"
          icon={<AddRegular />}
          onClick={handleAdd}
        >
          Add
        </Button>
      </div>

      <div className={mergeClasses(styles.card, theme === 'dark' && styles.cardDark)}>
        {loading ? (
          <div className={styles.loadingContainer}>
            <Spinner size="small" />
          </div>
        ) : error ? (
          <MessageBar intent="error">
            <MessageBarBody>{error}</MessageBarBody>
          </MessageBar>
        ) : items.length === 0 ? (
          <div className={styles.emptyState}>
            <Text>No related items</Text>
          </div>
        ) : (
          <div className={styles.dataGridWrapper}>
            <DataGrid
              items={items}
              columns={tableColumns}
              getRowId={(item) => item.id}
            >
              <DataGridHeader>
                <DataGridRow>
                  {({ renderHeaderCell }) => (
                    <DataGridHeaderCell>{renderHeaderCell()}</DataGridHeaderCell>
                  )}
                </DataGridRow>
              </DataGridHeader>
              <DataGridBody<GraphListItem>>
                {({ item, rowId }) => (
                  <DataGridRow<GraphListItem>
                    key={rowId}
                    onClick={() => handleRowClick(item)}
                  >
                    {({ renderCell }) => (
                      <DataGridCell>{renderCell(item)}</DataGridCell>
                    )}
                  </DataGridRow>
                )}
              </DataGridBody>
            </DataGrid>
          </div>
        )}
      </div>

      {/* Form modal for create */}
      {formModalOpen && (
        <ItemFormModal
          mode="create"
          siteId={section.source.siteId}
          listId={section.source.listId}
          initialValues={{}}
          saving={saving}
          onSave={handleFormSave}
          onClose={() => setFormModalOpen(false)}
          prefillLookupField={section.lookupColumn}
        />
      )}
    </div>
  );
}
