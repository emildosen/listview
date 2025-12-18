import { useState, useEffect, useCallback, useRef, useMemo } from 'react';
import { useMsal } from '@azure/msal-react';
import type { SPFI } from '@pnp/sp';
import {
  makeStyles,
  tokens,
  Card,
  Text,
  Button,
  Badge,
  Spinner,
  MessageBar,
  MessageBarBody,
  Link,
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
  DocumentTextRegular,
  OpenRegular,
  EditRegular,
  DeleteRegular,
  AddRegular,
} from '@fluentui/react-icons';
import { getListItems, type GraphListColumn, type GraphListItem } from '../../auth/graphClient';
import { createListItem, updateListItem, deleteListItem, createSPClient, getColumnFormatting, parseColumnFormattingForLink, getListColumnOrder } from '../../services/sharepoint';
import type { PageDefinition, RelatedSection } from '../../types/page';
import { useSettings } from '../../contexts/SettingsContext';
import ItemFormModal from './ItemFormModal';
import { SharePointLink } from '../common/SharePointLink';
import { isSharePointUrl } from '../../auth/graphClient';

const MAX_TEXT_LENGTH = 200;

const useStyles = makeStyles({
  container: {
    display: 'flex',
    flexDirection: 'column',
    width: '100%',
    minWidth: 0,
  },
  emptyState: {
    height: '100%',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    flexDirection: 'column',
    color: tokens.colorNeutralForeground2,
  },
  emptyIcon: {
    opacity: 0.3,
    marginBottom: '12px',
  },
  header: {
    paddingBottom: '16px',
    borderBottom: `1px solid ${tokens.colorNeutralStroke1}`,
  },
  headerTitle: {
    fontSize: tokens.fontSizeBase500,
    fontWeight: tokens.fontWeightBold,
  },
  content: {
    paddingTop: '16px',
    display: 'flex',
    flexDirection: 'column',
    gap: '24px',
    width: '100%',
    minWidth: 0,
  },
  cardBody: {
    padding: '16px',
    width: '100%',
    boxSizing: 'border-box',
    minWidth: 0,
    overflowX: 'auto',
  },
  cardTitle: {
    fontWeight: tokens.fontWeightMedium,
    marginBottom: '12px',
  },
  detailsGrid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(2, 1fr)',
    gap: '16px',
    width: '100%',
  },
  detailItem: {
    display: 'flex',
    flexDirection: 'column',
    gap: '4px',
    width: '100%',
  },
  detailLabel: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground2,
  },
  detailValue: {
    fontWeight: tokens.fontWeightMedium,
    wordBreak: 'break-word',
  },
  showMoreButton: {
    color: tokens.colorBrandForeground1,
    cursor: 'pointer',
    marginLeft: '4px',
    fontSize: tokens.fontSizeBase200,
    ':hover': {
      textDecoration: 'underline',
    },
  },
  sectionHeader: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    marginBottom: '16px',
  },
  sectionTitle: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
    fontWeight: tokens.fontWeightMedium,
  },
  emptySection: {
    textAlign: 'center',
    padding: '16px',
    color: tokens.colorNeutralForeground2,
    fontSize: tokens.fontSizeBase200,
  },
  loadingSection: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    padding: '32px',
  },
  actionsCell: {
    display: 'flex',
    alignItems: 'center',
    gap: '4px',
  },
  dataGridWrapper: {
    width: '100%',
    overflowX: 'auto',
  },
  addButton: {
    marginBottom: '12px',
  },
});

function TruncatedText({ text, maxLength = MAX_TEXT_LENGTH }: { text: string; maxLength?: number }) {
  const styles = useStyles();
  const [expanded, setExpanded] = useState(false);

  if (!text || typeof text !== 'string' || text.length <= maxLength) {
    return <>{text}</>;
  }

  const displayText = expanded ? text : text.slice(0, maxLength) + '...';

  return (
    <div>
      {displayText}
      <span
        className={styles.showMoreButton}
        onClick={() => setExpanded(!expanded)}
      >
        {expanded ? 'Show less' : 'Show more'}
      </span>
    </div>
  );
}

interface DetailPanelProps {
  page: PageDefinition;
  columns: GraphListColumn[];
  item: GraphListItem | null;
  spClient: SPFI | null;
}

function DetailPanel({ page, columns, item, spClient }: DetailPanelProps) {
  const styles = useStyles();

  // Track which columns have link formatting (from custom formatter JSON)
  const [linkFormattedColumns, setLinkFormattedColumns] = useState<Set<string>>(new Set());

  // Fetch column formatting when spClient and list are available
  useEffect(() => {
    if (!spClient || !page.primarySource?.listId) return;

    const fetchFormatting = async () => {
      try {
        const formatting = await getColumnFormatting(spClient, page.primarySource.listId);
        const linkCols = new Set<string>();
        for (const col of formatting) {
          if (parseColumnFormattingForLink(col.customFormatter)) {
            linkCols.add(col.internalName);
          }
        }
        setLinkFormattedColumns(linkCols);
      } catch (err) {
        console.error('Failed to fetch column formatting:', err);
      }
    };

    fetchFormatting();
  }, [spClient, page.primarySource?.listId]);

  // Check if column should render as link (either hyperlinkOrPicture type or custom link formatting)
  const isLinkColumn = useCallback((columnName: string): boolean => {
    const col = columns.find((c) => c.name === columnName);
    if (col?.hyperlinkOrPicture) return true;
    return linkFormattedColumns.has(columnName);
  }, [columns, linkFormattedColumns]);

  const getDisplayValue = useCallback((columnName: string): string => {
    if (!item) return '-';
    const value = item.fields[columnName];
    if (value === null || value === undefined) return '-';
    if (typeof value === 'object') {
      if ('LookupValue' in value) {
        return (value as { LookupValue: string }).LookupValue;
      }
      return JSON.stringify(value);
    }
    if (typeof value === 'boolean') {
      return value ? 'Yes' : 'No';
    }
    if (value instanceof Date) {
      return value.toLocaleDateString();
    }
    return String(value);
  }, [item]);

  const renderValue = useCallback((columnName: string) => {
    const value = getDisplayValue(columnName);
    if (value === '-') return value;

    // Check if value starts with SharePoint URL - treat entire value as URL
    if (isSharePointUrl(value)) {
      return <SharePointLink url={value} stopPropagation={false} />;
    }

    // Render link columns as clickable links
    if (isLinkColumn(columnName)) {
      return (
        <Link
          href={value}
          target="_blank"
          rel="noopener noreferrer"
          style={{ wordBreak: 'break-all' }}
        >
          {value}
        </Link>
      );
    }

    return <TruncatedText text={value} />;
  }, [getDisplayValue, isLinkColumn]);

  // Early return for empty state (after all hooks)
  if (!item) {
    return (
      <div className={styles.emptyState}>
        <DocumentTextRegular fontSize={48} className={styles.emptyIcon} />
        <Text>Select an item to view details</Text>
      </div>
    );
  }

  const titleColumn = page.searchConfig?.titleColumn || 'Title';
  const titleValue = getDisplayValue(titleColumn);

  return (
    <div className={styles.container}>
      {/* Header */}
      <div className={styles.header}>
        <Text className={styles.headerTitle}>{titleValue}</Text>
      </div>

      {/* Content */}
      <div className={styles.content}>
        {/* Details Section */}
        <Card className={styles.cardBody}>
          <Text className={styles.cardTitle}>Details</Text>
          <div className={styles.detailsGrid}>
            {page.displayColumns
              .filter((col) => col.internalName !== titleColumn)
              .map((col) => (
                <div key={col.internalName} className={styles.detailItem}>
                  <Text className={styles.detailLabel}>{col.displayName}</Text>
                  <Text className={styles.detailValue}>
                    {renderValue(col.internalName)}
                  </Text>
                </div>
              ))}
          </div>
        </Card>

        {/* Related Sections */}
        {page.relatedSections?.map((section) => (
          <RelatedSectionComponent
            key={section.id}
            section={section}
            parentItem={item}
          />
        ))}
      </div>
    </div>
  );
}

interface RowData {
  id: string;
  _item: GraphListItem;
  [key: string]: unknown;
}

interface RelatedSectionComponentProps {
  section: RelatedSection;
  parentItem: GraphListItem;
}

function RelatedSectionComponent({
  section,
  parentItem,
}: RelatedSectionComponentProps) {
  const styles = useStyles();
  const { instance, accounts } = useMsal();
  const { enabledLists } = useSettings();
  const account = accounts[0];

  const [items, setItems] = useState<GraphListItem[]>([]);
  const [columns, setColumns] = useState<GraphListColumn[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  // Site-specific SP client for this related list
  const spClientRef = useRef<SPFI | null>(null);
  const [spClientReady, setSpClientReady] = useState(false);

  // Modal state
  const [modalOpen, setModalOpen] = useState(false);
  const [modalMode, setModalMode] = useState<'create' | 'edit'>('create');
  const [editingItem, setEditingItem] = useState<GraphListItem | null>(null);
  const [saving, setSaving] = useState(false);
  const [deleting, setDeleting] = useState<string | null>(null);

  // Column order from SharePoint default view
  const [columnOrder, setColumnOrder] = useState<string[]>([]);

  // Get siteUrl - from section source or look up from enabledLists for backwards compat
  const siteUrl = useMemo(() => {
    if (section.source.siteUrl) {
      return section.source.siteUrl;
    }
    // Backwards compatibility: look up from enabledLists
    const list = enabledLists.find(
      (l) => l.siteId === section.source.siteId && l.listId === section.source.listId
    );
    return list?.siteUrl;
  }, [section.source.siteUrl, section.source.siteId, section.source.listId, enabledLists]);

  // Create site-specific SP client
  useEffect(() => {
    if (!account || !siteUrl) {
      setSpClientReady(false);
      return;
    }

    const initClient = async () => {
      try {
        const client = await createSPClient(instance, account, siteUrl);
        spClientRef.current = client;
        setSpClientReady(true);
      } catch (err) {
        console.error('Failed to create SP client for related list:', err);
        setSpClientReady(false);
      }
    };

    initClient();
  }, [instance, account, siteUrl]);

  // Fetch column order from SharePoint default view
  useEffect(() => {
    if (!spClientReady || !spClientRef.current || !section.source.listId) return;

    const fetchColumnOrder = async () => {
      try {
        const order = await getListColumnOrder(spClientRef.current!, section.source.listId);
        setColumnOrder(order);
      } catch (err) {
        console.error('Failed to fetch column order:', err);
      }
    };

    fetchColumnOrder();
  }, [spClientReady, section.source.listId]);

  // Load related items
  const loadRelatedItems = useCallback(async () => {
    if (!account || !section.source.siteId || !section.source.listId) {
      setLoading(false);
      return;
    }

    setLoading(true);
    setError(null);

    try {
      const result = await getListItems(
        instance,
        account,
        section.source.siteId,
        section.source.listId
      );

      // Filter items by lookup column matching parent item ID
      const parentId = parentItem.id;
      let filteredItems = result.items.filter((item) => {
        const lookupValue = item.fields[`${section.lookupColumn}LookupId`];
        return String(lookupValue) === parentId;
      });

      // Sort items if defaultSort is configured
      if (section.defaultSort?.column) {
        const { column, direction } = section.defaultSort;
        filteredItems = [...filteredItems].sort((a, b) => {
          const aVal = a.fields[column];
          const bVal = b.fields[column];

          // Handle null/undefined
          if (aVal == null && bVal == null) return 0;
          if (aVal == null) return direction === 'asc' ? 1 : -1;
          if (bVal == null) return direction === 'asc' ? -1 : 1;

          // Compare values
          let comparison = 0;
          if (typeof aVal === 'string' && typeof bVal === 'string') {
            comparison = aVal.localeCompare(bVal);
          } else if (typeof aVal === 'number' && typeof bVal === 'number') {
            comparison = aVal - bVal;
          } else {
            comparison = String(aVal).localeCompare(String(bVal));
          }

          return direction === 'desc' ? -comparison : comparison;
        });
      }

      setColumns(result.columns);
      setItems(filteredItems);
    } catch (err) {
      console.error('Failed to load related items:', err);
      setError('Failed to load related items');
    } finally {
      setLoading(false);
    }
  }, [instance, account, section, parentItem.id]);

  useEffect(() => {
    loadRelatedItems();
  }, [loadRelatedItems]);

  const handleCreate = () => {
    setModalMode('create');
    setEditingItem(null);
    setModalOpen(true);
  };

  const handleEdit = useCallback((item: GraphListItem) => {
    setModalMode('edit');
    setEditingItem(item);
    setModalOpen(true);
  }, []);

  const handleDelete = useCallback(async (itemId: string) => {
    if (!spClientRef.current || !confirm('Are you sure you want to delete this item?')) {
      return;
    }

    setDeleting(itemId);
    try {
      await deleteListItem(spClientRef.current, section.source.listId, parseInt(itemId, 10));
      await loadRelatedItems();
    } catch (err) {
      console.error('Failed to delete item:', err);
    } finally {
      setDeleting(null);
    }
  }, [section.source.listId, loadRelatedItems]);

  const handleSave = async (fields: Record<string, unknown>) => {
    if (!spClientRef.current) return;

    setSaving(true);
    try {
      // Add lookup field for parent relationship
      // SharePoint REST API uses {FieldName}Id for setting lookup values
      const saveFields = {
        ...fields,
        [`${section.lookupColumn}Id`]: parseInt(parentItem.id, 10),
      };

      if (modalMode === 'create') {
        await createListItem(spClientRef.current, section.source.listId, saveFields);
      } else if (editingItem) {
        await updateListItem(
          spClientRef.current,
          section.source.listId,
          parseInt(editingItem.id, 10),
          fields
        );
      }

      setModalOpen(false);
      await loadRelatedItems();
    } catch (err) {
      console.error('Failed to save item:', err);
      throw err;
    } finally {
      setSaving(false);
    }
  };

  // Format cell value for display
  const formatCellValue = useCallback((value: unknown): string => {
    if (value === null || value === undefined) return '-';
    if (typeof value === 'object') {
      if ('LookupValue' in value) {
        return (value as { LookupValue: string }).LookupValue;
      }
      return JSON.stringify(value);
    }
    if (typeof value === 'boolean') return value ? 'Yes' : 'No';
    if (typeof value === 'string' && /^\d{4}-\d{2}-\d{2}/.test(value)) {
      return new Date(value).toLocaleDateString();
    }
    return String(value);
  }, []);

  // Convert items to row data
  const rowData = useMemo((): RowData[] => {
    return items.map((item) => ({
      id: item.id,
      _item: item,
      ...item.fields,
    }));
  }, [items]);

  // Generate Fluent UI DataGrid column definitions
  const columnDefs = useMemo((): TableColumnDefinition<RowData>[] => {
    const cols: TableColumnDefinition<RowData>[] = section.displayColumns.map((col) =>
      createTableColumn<RowData>({
        columnId: col.internalName,
        compare: (a, b) => {
          const aVal = String(a[col.internalName] ?? '');
          const bVal = String(b[col.internalName] ?? '');
          return aVal.localeCompare(bVal);
        },
        renderHeaderCell: () => col.displayName,
        renderCell: (item) => (
          <TableCellLayout truncate>
            {formatCellValue(item[col.internalName])}
          </TableCellLayout>
        ),
      })
    );

    // Add actions column
    cols.push(
      createTableColumn<RowData>({
        columnId: '_actions',
        renderHeaderCell: () => 'Actions',
        renderCell: (item) => {
          const sharePointUrl = siteUrl
            ? `${siteUrl}/_layouts/15/listform.aspx?PageType=4&ListId=${encodeURIComponent(section.source.listId)}&ID=${item.id}`
            : null;
          return (
            <div className={styles.actionsCell}>
              {sharePointUrl && (
                <Button
                  as="a"
                  href={sharePointUrl}
                  target="_blank"
                  rel="noopener noreferrer"
                  appearance="subtle"
                  size="small"
                  icon={<OpenRegular />}
                  title="Open in SharePoint"
                />
              )}
              {section.allowEdit && (
                <Button
                  appearance="subtle"
                  size="small"
                  icon={<EditRegular />}
                  onClick={() => handleEdit(item._item)}
                  title="Edit"
                />
              )}
              {section.allowDelete && (
                <Button
                  appearance="subtle"
                  size="small"
                  icon={deleting === item.id ? <Spinner size="tiny" /> : <DeleteRegular />}
                  onClick={() => handleDelete(item.id)}
                  disabled={deleting === item.id}
                  title="Delete"
                  style={{ color: tokens.colorPaletteRedForeground1 }}
                />
              )}
            </div>
          );
        },
      })
    );

    return cols;
  }, [section.displayColumns, section.allowEdit, section.allowDelete, section.source.listId, siteUrl, deleting, handleEdit, handleDelete, formatCellValue, styles.actionsCell]);

  return (
    <Card className={styles.cardBody}>
      <div className={styles.sectionHeader}>
        <div className={styles.sectionTitle}>
          <Text weight="medium">{section.title}</Text>
          <Badge appearance="outline" size="small">{items.length}</Badge>
        </div>
        {section.allowCreate && (
          <Button
            appearance="primary"
            size="small"
            icon={<AddRegular />}
            onClick={handleCreate}
            disabled={!spClientReady}
          >
            Add
          </Button>
        )}
      </div>

      {loading ? (
        <div className={styles.loadingSection}>
          <Spinner size="small" />
        </div>
      ) : error ? (
        <MessageBar intent="error">
          <MessageBarBody>{error}</MessageBarBody>
        </MessageBar>
      ) : items.length === 0 ? (
        <div className={styles.emptySection}>
          <Text>No {section.title.toLowerCase()} yet</Text>
        </div>
      ) : (
        <div className={styles.dataGridWrapper}>
          <DataGrid
            items={rowData}
            columns={columnDefs}
            sortable
            getRowId={(item) => item.id}
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
      )}

      {/* Item Form Modal */}
      {modalOpen && (
        <ItemFormModal
          mode={modalMode}
          columns={columns
            .filter((c) => section.displayColumns.some((dc) => dc.internalName === c.name))
            .sort((a, b) => {
              const aIndex = columnOrder.indexOf(a.name);
              const bIndex = columnOrder.indexOf(b.name);
              // Columns not in columnOrder go to the end
              if (aIndex === -1 && bIndex === -1) return 0;
              if (aIndex === -1) return 1;
              if (bIndex === -1) return -1;
              return aIndex - bIndex;
            })}
          initialValues={editingItem?.fields || {}}
          saving={saving}
          onSave={handleSave}
          onClose={() => setModalOpen(false)}
        />
      )}
    </Card>
  );
}

export default DetailPanel;
