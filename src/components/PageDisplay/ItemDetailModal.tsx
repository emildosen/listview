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
  Dialog,
  DialogSurface,
  DialogBody,
  DialogTitle,
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
  DismissRegular,
  SettingsRegular,
  EditRegular,
  DeleteRegular,
  AddRegular,
  OpenRegular,
} from '@fluentui/react-icons';
import { getListItems, type GraphListColumn, type GraphListItem } from '../../auth/graphClient';
import { createListItem, updateListItem, deleteListItem, createSPClient, getColumnFormatting, parseColumnFormattingForLink } from '../../services/sharepoint';
import type { PageDefinition, PageColumn, RelatedSection, DetailLayoutConfig, ListDetailConfig } from '../../types/page';
import { useSettings } from '../../contexts/SettingsContext';
import ItemFormModal from './ItemFormModal';
import StatBox from './StatBox';
import DetailCustomizeDrawer from './DetailCustomizeDrawer';
import { SharePointLink } from '../common/SharePointLink';
import { isSharePointUrl } from '../../auth/graphClient';

const MAX_TEXT_LENGTH = 200;

const useStyles = makeStyles({
  surface: {
    maxWidth: '1000px',
    width: '95vw',
    maxHeight: '90vh',
  },
  dialogTitle: {
    display: 'flex',
    flexDirection: 'row',
    alignItems: 'center',
    justifyContent: 'space-between',
    gap: '16px'
  },
  titleText: {
    fontSize: tokens.fontSizeBase500,
    fontWeight: tokens.fontWeightSemibold,
    lineHeight: tokens.lineHeightBase500,
    flex: 1,
    minWidth: 0,
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap',
  },
  headerActions: {
    display: 'flex',
    alignItems: 'center',
    gap: '4px',
    flexShrink: 0,
  },
  statBoxContainer: {
    display: 'flex',
    flexWrap: 'wrap',
    gap: '8px',
    marginTop: '10px'
  },
  body: {
    display: 'block',
    overflowY: 'auto',
    maxHeight: 'calc(90vh - 80px)',
    '& > *': {
      marginBottom: '24px',
    },
    '& > *:last-child': {
      marginBottom: 0,
    },
  },
  cardBody: {
    padding: '16px',
    width: '100%',
    boxSizing: 'border-box',
  },
  cardTitle: {
    fontWeight: tokens.fontWeightMedium,
    marginBottom: '6px',
  },
  detailsGrid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(2, 1fr)',
    gap: '10px',
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
    justifyContent: 'flex-end',
    gap: '4px',
    width: '100%',
  },
  actionsColumn: {
    width: '100px',
    minWidth: '100px',
    maxWidth: '100px',
  },
  dataGridWrapper: {
    width: '100%',
    '& [role="row"]': {
      display: 'flex',
      width: '100%',
    },
    // All cells get overflow handling
    '& [role="columnheader"], & [role="gridcell"]': {
      overflow: 'hidden',
    },
    // Data columns share available space
    '& [role="columnheader"]:not(:last-child), & [role="gridcell"]:not(:last-child)': {
      flex: '1 1 0',
      minWidth: 0,
    },
    // Actions column fixed width
    '& [role="columnheader"]:last-child, & [role="gridcell"]:last-child': {
      flex: '0 0 100px',
    },
  },
  cellText: {
    display: '-webkit-box',
    WebkitLineClamp: 2,
    WebkitBoxOrient: 'vertical',
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'pre-wrap',
    wordBreak: 'break-word',
    maxWidth: '100%',
  },
  cellTextExpanded: {
    display: 'block',
    whiteSpace: 'pre-wrap',
    wordBreak: 'break-word',
  },
  cellShowMore: {
    color: tokens.colorBrandForeground1,
    cursor: 'pointer',
    fontSize: tokens.fontSizeBase200,
    marginTop: '4px',
    ':hover': {
      textDecoration: 'underline',
    },
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

// Expandable cell text for DataGrid - shows 2 lines with "Show more" toggle
function ExpandableCellText({ text }: { text: string }) {
  const styles = useStyles();
  const [expanded, setExpanded] = useState(false);
  const textRef = useRef<HTMLDivElement>(null);
  const [isTruncated, setIsTruncated] = useState(false);

  useEffect(() => {
    const el = textRef.current;
    if (el) {
      // Check if text is actually truncated (scrollHeight > clientHeight)
      setIsTruncated(el.scrollHeight > el.clientHeight + 2);
    }
  }, [text]);

  if (!text || typeof text !== 'string' || text.length < 50) {
    return <>{text}</>;
  }

  return (
    <div>
      <div
        ref={textRef}
        className={expanded ? styles.cellTextExpanded : styles.cellText}
      >
        {text}
      </div>
      {(isTruncated || expanded) && (
        <span
          className={styles.cellShowMore}
          onClick={(e) => {
            e.stopPropagation();
            setExpanded(!expanded);
          }}
        >
          {expanded ? 'Show less' : 'Show more'}
        </span>
      )}
    </div>
  );
}

// Helper to get default layout config from display columns
function getDefaultLayoutConfig(displayColumns: PageColumn[], relatedSections: RelatedSection[]): DetailLayoutConfig {
  return {
    columnSettings: displayColumns.map(col => ({
      internalName: col.internalName,
      visible: true,
      displayStyle: 'list' as const,
    })),
    relatedSectionOrder: relatedSections.map(s => s.id),
  };
}

// Helper to merge existing config with defaults for new columns/sections
function getEffectiveLayoutConfig(
  existingLayout: DetailLayoutConfig | undefined,
  displayColumns: PageColumn[],
  relatedSections: RelatedSection[]
): DetailLayoutConfig {
  const defaults = getDefaultLayoutConfig(displayColumns, relatedSections);

  if (!existingLayout) {
    return defaults;
  }

  // Merge column settings - preserve order from existing settings
  const validColumnNames = new Set(displayColumns.map(c => c.internalName));
  const existingNames = new Set(existingLayout.columnSettings.map(s => s.internalName));

  // Keep existing settings in order, filtering out any that no longer exist
  const existingSettings = existingLayout.columnSettings.filter(s =>
    validColumnNames.has(s.internalName)
  );

  // Add new columns at the end with default settings
  const newColumns = displayColumns
    .filter(col => !existingNames.has(col.internalName))
    .map(col => ({
      internalName: col.internalName,
      visible: true,
      displayStyle: 'list' as const,
    }));

  const columnSettings = [...existingSettings, ...newColumns];

  // Merge section order
  let relatedSectionOrder: string[];
  if (existingLayout.relatedSectionOrder) {
    const existingOrder = existingLayout.relatedSectionOrder;
    const allIds = new Set(relatedSections.map(s => s.id));
    // Keep existing order for sections that still exist, add new ones at end
    const orderedIds = existingOrder.filter(id => allIds.has(id));
    const newIds = relatedSections
      .map(s => s.id)
      .filter(id => !existingOrder.includes(id));
    relatedSectionOrder = [...orderedIds, ...newIds];
  } else {
    relatedSectionOrder = defaults.relatedSectionOrder!;
  }

  return { columnSettings, relatedSectionOrder };
}

// Helper to create default ListDetailConfig from columns
function createDefaultListDetailConfigFromColumns(
  listId: string,
  listName: string,
  siteId: string,
  siteUrl: string | undefined,
  columns: GraphListColumn[]
): ListDetailConfig {
  // Use all non-system columns as display columns
  const displayColumns: PageColumn[] = columns
    .filter(c => !c.readOnly && c.name !== 'ContentType' && c.name !== 'Attachments')
    .map(c => ({
      internalName: c.name,
      displayName: c.displayName,
      editable: !c.readOnly,
    }));

  return {
    listId,
    listName,
    siteId,
    siteUrl,
    displayColumns,
    detailLayout: getDefaultLayoutConfig(displayColumns, []),
    relatedSections: [],
  };
}

interface ItemDetailModalProps {
  // List identification - required
  listId: string;
  listName: string;
  siteId: string;
  siteUrl?: string;
  // Data
  columns: GraphListColumn[];
  item: GraphListItem;
  spClient: SPFI | null;
  // Optional page for initial defaults (used by lookup pages)
  page?: PageDefinition;
  // Optional title column override
  titleColumnOverride?: string;
  // Callbacks
  onClose: () => void;
}

function ItemDetailModal({
  listId,
  listName,
  siteId,
  siteUrl,
  columns,
  item,
  spClient,
  page,
  titleColumnOverride,
  onClose
}: ItemDetailModalProps) {
  const styles = useStyles();
  const { getListDetailConfig, saveListDetailConfig } = useSettings();
  const [customizeOpen, setCustomizeOpen] = useState(false);

  // Track which columns have link formatting
  const [linkFormattedColumns, setLinkFormattedColumns] = useState<Set<string>>(new Set());

  // Get or create list detail config - always uses list-level config, never page-specific
  const listDetailConfig = useMemo((): ListDetailConfig => {
    // First check if we have a saved config for this list
    const savedConfig = getListDetailConfig(listId);
    if (savedConfig) {
      return savedConfig;
    }

    // No saved config - create default from columns (ignoring any page-specific config)
    return createDefaultListDetailConfigFromColumns(listId, listName, siteId, siteUrl, columns);
  }, [getListDetailConfig, listId, listName, siteId, siteUrl, columns]);

  // Fetch column formatting
  useEffect(() => {
    if (!spClient || !listId) return;

    const fetchFormatting = async () => {
      try {
        const formatting = await getColumnFormatting(spClient, listId);
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
  }, [spClient, listId]);

  // Get effective layout configuration
  const layoutConfig = useMemo(() =>
    getEffectiveLayoutConfig(
      listDetailConfig.detailLayout,
      listDetailConfig.displayColumns,
      listDetailConfig.relatedSections
    ),
    [listDetailConfig]
  );

  // Get the title column - priority: override, first table column (from page), first display column, or Title
  const titleColumn = useMemo(() => {
    if (titleColumnOverride) return titleColumnOverride;
    if (page?.searchConfig?.tableColumns?.[0]?.internalName) {
      return page.searchConfig.tableColumns[0].internalName;
    }
    return listDetailConfig.displayColumns[0]?.internalName || 'Title';
  }, [titleColumnOverride, page, listDetailConfig.displayColumns]);

  // Separate visible columns by display style
  const { statColumns, listColumns } = useMemo(() => {
    const visible = layoutConfig.columnSettings.filter(c => c.visible);
    // Exclude title column from both lists (it's shown in the header)
    return {
      statColumns: visible.filter(c => c.displayStyle === 'stat' && c.internalName !== titleColumn),
      listColumns: visible.filter(c => c.displayStyle === 'list' && c.internalName !== titleColumn),
    };
  }, [layoutConfig, titleColumn]);

  // Order related sections
  const orderedSections = useMemo(() => {
    if (!layoutConfig.relatedSectionOrder) return listDetailConfig.relatedSections;
    const sectionMap = new Map(listDetailConfig.relatedSections.map(s => [s.id, s]));
    return layoutConfig.relatedSectionOrder
      .map(id => sectionMap.get(id))
      .filter((s): s is RelatedSection => s !== undefined);
  }, [listDetailConfig.relatedSections, layoutConfig.relatedSectionOrder]);

  // Check if column should render as link
  const isLinkColumn = useCallback((columnName: string): boolean => {
    const col = columns.find((c) => c.name === columnName);
    if (col?.hyperlinkOrPicture) return true;
    return linkFormattedColumns.has(columnName);
  }, [columns, linkFormattedColumns]);

  const getDisplayValue = useCallback((columnName: string): string => {
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
  }, [item.fields]);

  const getColumnDisplayName = useCallback((internalName: string): string => {
    const col = listDetailConfig.displayColumns.find(c => c.internalName === internalName);
    return col?.displayName || internalName;
  }, [listDetailConfig.displayColumns]);

  const renderValue = useCallback((columnName: string) => {
    const value = getDisplayValue(columnName);
    if (value === '-') return value;

    // Check if value starts with SharePoint URL - treat entire value as URL
    if (isSharePointUrl(value)) {
      return <SharePointLink url={value} stopPropagation={false} />;
    }

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

  const handleSaveConfig = async (config: DetailLayoutConfig, relatedSections?: RelatedSection[]) => {
    // Save to list-level config
    const updatedConfig: ListDetailConfig = {
      ...listDetailConfig,
      detailLayout: config,
      relatedSections: relatedSections ?? listDetailConfig.relatedSections,
    };
    await saveListDetailConfig(updatedConfig);
    setCustomizeOpen(false);
  };

  const titleValue = getDisplayValue(titleColumn);

  return (
    <>
      <Dialog open onOpenChange={(_, data) => !data.open && onClose()}>
        <DialogSurface className={styles.surface}>
          <DialogTitle className={styles.dialogTitle}>
            <Text className={styles.titleText}>{titleValue}</Text>
            <div className={styles.headerActions}>
              <Button
                appearance="subtle"
                icon={<SettingsRegular />}
                onClick={() => setCustomizeOpen(true)}
              >
                Customize
              </Button>
              <Button
                appearance="subtle"
                icon={<DismissRegular />}
                onClick={onClose}
                aria-label="Close"
              />
            </div>
          </DialogTitle>

          <DialogBody className={styles.body}>
            {/* Stat Boxes */}
            {statColumns.length > 0 && (
              <div className={styles.statBoxContainer}>
                {statColumns.map(col => (
                  <StatBox
                    key={col.internalName}
                    label={getColumnDisplayName(col.internalName)}
                    value={getDisplayValue(col.internalName)}
                  />
                ))}
              </div>
            )}
            {/* Details Card */}
            {listColumns.length > 0 && (
              <Card className={styles.cardBody}>
                <Text className={styles.cardTitle}>Details</Text>
                <div className={styles.detailsGrid}>
                  {listColumns.map(col => (
                    <div key={col.internalName} className={styles.detailItem}>
                      <Text className={styles.detailLabel}>
                        {getColumnDisplayName(col.internalName)}
                      </Text>
                      <Text className={styles.detailValue}>
                        {renderValue(col.internalName)}
                      </Text>
                    </div>
                  ))}
                </div>
              </Card>
            )}

            {/* Related Sections */}
            {orderedSections.map(section => (
              <RelatedSectionComponent
                key={section.id}
                section={section}
                parentItem={item}
              />
            ))}
          </DialogBody>
        </DialogSurface>
      </Dialog>

      {/* Customize Drawer */}
      <DetailCustomizeDrawer
        listDetailConfig={listDetailConfig}
        titleColumn={titleColumn}
        open={customizeOpen}
        onClose={() => setCustomizeOpen(false)}
        onSave={handleSaveConfig}
      />
    </>
  );
}

// Related Section Component (extracted from DetailPanel)
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

  const spClientRef = useRef<SPFI | null>(null);
  const [spClientReady, setSpClientReady] = useState(false);

  const [modalOpen, setModalOpen] = useState(false);
  const [modalMode, setModalMode] = useState<'create' | 'edit'>('create');
  const [editingItem, setEditingItem] = useState<GraphListItem | null>(null);
  const [saving, setSaving] = useState(false);
  const [deleting, setDeleting] = useState<string | null>(null);

  const siteUrl = useMemo(() => {
    if (section.source.siteUrl) {
      return section.source.siteUrl;
    }
    const list = enabledLists.find(
      (l) => l.siteId === section.source.siteId && l.listId === section.source.listId
    );
    return list?.siteUrl;
  }, [section.source.siteUrl, section.source.siteId, section.source.listId, enabledLists]);

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

      const parentId = parentItem.id;
      let filteredItems = result.items.filter((item) => {
        const lookupValue = item.fields[`${section.lookupColumn}LookupId`];
        return String(lookupValue) === parentId;
      });

      if (section.defaultSort?.column) {
        const { column, direction } = section.defaultSort;
        filteredItems = [...filteredItems].sort((a, b) => {
          const aVal = a.fields[column];
          const bVal = b.fields[column];

          if (aVal == null && bVal == null) return 0;
          if (aVal == null) return direction === 'asc' ? 1 : -1;
          if (bVal == null) return direction === 'asc' ? -1 : 1;

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

  const rowData = useMemo((): RowData[] => {
    return items.map((item) => ({
      id: item.id,
      _item: item,
      ...item.fields,
    }));
  }, [items]);

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
          <TableCellLayout>
            <ExpandableCellText text={formatCellValue(item[col.internalName])} />
          </TableCellLayout>
        ),
      })
    );

    cols.push(
      createTableColumn<RowData>({
        columnId: '_actions',
        renderHeaderCell: () => (
          <span style={{ display: 'block', textAlign: 'right', width: '100%' }}>Actions</span>
        ),
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

      {modalOpen && (
        <ItemFormModal
          mode={modalMode}
          columns={columns.filter((c) =>
            section.displayColumns.some((dc) => dc.internalName === c.name)
          )}
          initialValues={editingItem?.fields || {}}
          saving={saving}
          onSave={handleSave}
          onClose={() => setModalOpen(false)}
        />
      )}
    </Card>
  );
}

export default ItemDetailModal;
