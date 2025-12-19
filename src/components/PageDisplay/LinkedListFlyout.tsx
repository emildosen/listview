import { useState, useEffect, useCallback, useMemo } from 'react';
import { useMsal } from '@azure/msal-react';
import {
  makeStyles,
  tokens,
  Text,
  Button,
  Input,
  Field,
  Dropdown,
  Option,
  Badge,
  Spinner,
  DrawerBody,
  DrawerHeader,
  DrawerHeaderTitle,
  OverlayDrawer,
} from '@fluentui/react-components';
import { DismissRegular, ArrowLeftRegular, ArrowRightRegular } from '@fluentui/react-icons';
import { getListColumns, getSiteLists, type GraphListColumn, type GraphList } from '../../auth/graphClient';
import { SYSTEM_LIST_NAMES } from '../../services/sharepoint';
import type { RelatedSection, PageColumn } from '../../types/page';

const useStyles = makeStyles({
  drawerSurface: {
    // Ensure flyout appears on top of the parent drawer
    zIndex: 1001,
  },
  section: {
    display: 'flex',
    flexDirection: 'column',
    gap: '16px',
    marginBottom: '16px',
  },
  loadingRow: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
    padding: '8px 0',
  },
  badgeWrap: {
    display: 'flex',
    flexWrap: 'wrap',
    gap: '8px',
  },
  badgeItem: {
    cursor: 'pointer',
  },
  sortRow: {
    display: 'flex',
    gap: '8px',
  },
  sortColumn: {
    flex: 1,
  },
  sortDirection: {
    width: '120px',
  },
  footer: {
    display: 'flex',
    justifyContent: 'flex-end',
    gap: '8px',
    padding: '16px 0',
    borderTop: `1px solid ${tokens.colorNeutralStroke1}`,
    marginTop: 'auto',
  },
  warningText: {
    color: tokens.colorPaletteYellowForeground1,
  },
  listOption: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
  },
  directionBadge: {
    display: 'inline-flex',
    alignItems: 'center',
    gap: '4px',
    fontSize: tokens.fontSizeBase100,
  },
});

// Type for lists with lookup relationship info
interface LinkedListInfo {
  list: GraphList;
  direction: 'incoming' | 'outgoing';
  lookupColumn?: string; // The lookup column name (for incoming, it's in this list; for outgoing, it's in primary list)
  lookupColumnDisplayName?: string;
}

interface LinkedListFlyoutProps {
  open: boolean;
  section: RelatedSection | null; // null = creating new
  primaryListId: string;
  primarySiteId: string;
  primarySiteUrl: string;
  columnMetadata?: GraphListColumn[]; // Primary list columns for outgoing lookups
  onClose: () => void;
  onSave: (section: RelatedSection) => void;
}

function LinkedListFlyout({
  open,
  section,
  primaryListId,
  primarySiteId,
  primarySiteUrl,
  columnMetadata,
  onClose,
  onSave,
}: LinkedListFlyoutProps) {
  const styles = useStyles();
  const { instance, accounts } = useMsal();
  const account = accounts[0];

  const isEditing = section !== null;

  // Form state
  const [title, setTitle] = useState('');
  const [selectedListId, setSelectedListId] = useState('');
  const [lookupColumn, setLookupColumn] = useState('');
  const [displayColumns, setDisplayColumns] = useState<PageColumn[]>([]);
  const [defaultSort, setDefaultSort] = useState<{ column: string; direction: 'asc' | 'desc' } | undefined>();

  // Lists loading state
  const [allLists, setAllLists] = useState<GraphList[]>([]);
  const [loadingLists, setLoadingLists] = useState(false);

  // Column data for all lists (keyed by listId)
  const [listColumnsCache, setListColumnsCache] = useState<Record<string, GraphListColumn[]>>({});

  // Selected list columns
  const [columns, setColumns] = useState<GraphListColumn[]>([]);
  const [loadingColumns, setLoadingColumns] = useState(false);

  // Get linked lists with lookup relationships
  const linkedLists = useMemo((): LinkedListInfo[] => {
    const result: LinkedListInfo[] = [];

    for (const list of allLists) {
      if (list.id === primaryListId) continue;

      const listColumns = listColumnsCache[list.id];
      if (!listColumns) continue;

      // Check for incoming lookups (other list has lookup to primary)
      for (const col of listColumns) {
        if (col.lookup?.listId === primaryListId) {
          result.push({
            list,
            direction: 'incoming',
            lookupColumn: col.name,
            lookupColumnDisplayName: col.displayName,
          });
          break; // Only add once per list for incoming
        }
      }

      // Check for outgoing lookups (primary list has lookup to other list)
      if (columnMetadata) {
        for (const col of columnMetadata) {
          if (col.lookup?.listId === list.id) {
            // Check if not already added as incoming
            const alreadyAdded = result.some(r => r.list.id === list.id);
            if (!alreadyAdded) {
              result.push({
                list,
                direction: 'outgoing',
                lookupColumn: col.name,
                lookupColumnDisplayName: col.displayName,
              });
            }
            break;
          }
        }
      }
    }

    return result;
  }, [allLists, listColumnsCache, primaryListId, columnMetadata]);

  // Load all lists from the site when drawer opens
  useEffect(() => {
    if (!open || !account || !primarySiteId) {
      return;
    }

    const loadLists = async () => {
      setLoadingLists(true);
      try {
        const lists = await getSiteLists(instance, account, primarySiteId);
        // Filter out system lists
        const filtered = lists.filter(
          (list) => !SYSTEM_LIST_NAMES.includes(list.name as typeof SYSTEM_LIST_NAMES[number])
        );
        setAllLists(filtered);

        // Load columns for all lists to determine lookup relationships
        const columnsMap: Record<string, GraphListColumn[]> = {};
        await Promise.all(
          filtered.map(async (list) => {
            try {
              const cols = await getListColumns(instance, account, primarySiteId, list.id);
              columnsMap[list.id] = cols;
            } catch (err) {
              console.error(`Failed to load columns for list ${list.displayName}:`, err);
              columnsMap[list.id] = [];
            }
          })
        );
        setListColumnsCache(columnsMap);
      } catch (err) {
        console.error('Failed to load lists:', err);
        setAllLists([]);
      } finally {
        setLoadingLists(false);
      }
    };

    loadLists();
  }, [open, instance, account, primarySiteId]);

  // Initialize form when section changes
  useEffect(() => {
    if (section) {
      setTitle(section.title);
      setSelectedListId(section.source.listId || '');
      setLookupColumn(section.lookupColumn);
      setDisplayColumns(section.displayColumns);
      setDefaultSort(section.defaultSort);
    } else {
      // Reset for new section
      setTitle('');
      setSelectedListId('');
      setLookupColumn('');
      setDisplayColumns([]);
      setDefaultSort(undefined);
    }
  }, [section, open]);

  // Load columns when source list changes
  useEffect(() => {
    if (!selectedListId) {
      setColumns([]);
      return;
    }

    // Use cached columns if available
    const cached = listColumnsCache[selectedListId];
    if (cached) {
      setColumns(cached);
      return;
    }

    // Otherwise load them
    if (!account || !primarySiteId) {
      setColumns([]);
      return;
    }

    const loadColumns = async () => {
      setLoadingColumns(true);
      try {
        const cols = await getListColumns(instance, account, primarySiteId, selectedListId);
        setColumns(cols);
        setListColumnsCache(prev => ({ ...prev, [selectedListId]: cols }));
      } catch (err) {
        console.error('Failed to load columns:', err);
        setColumns([]);
      } finally {
        setLoadingColumns(false);
      }
    };

    loadColumns();
  }, [instance, account, selectedListId, primarySiteId, listColumnsCache]);

  // Get the selected linked list info
  const selectedLinkedListInfo = useMemo(() => {
    return linkedLists.find(ll => ll.list.id === selectedListId);
  }, [linkedLists, selectedListId]);

  // Get lookup columns based on direction
  const lookupColumns = useMemo(() => {
    if (!selectedLinkedListInfo) return [];

    if (selectedLinkedListInfo.direction === 'incoming') {
      // Incoming: lookup is in the selected list, pointing to primary list
      return columns.filter(col => col.lookup?.listId === primaryListId);
    } else {
      // Outgoing: lookup is in primary list, pointing to selected list
      // For outgoing, we need to use the Title column or ID of the selected list to match
      // Return a synthetic option for the lookup
      if (columnMetadata) {
        const outgoingLookup = columnMetadata.find(col => col.lookup?.listId === selectedListId);
        if (outgoingLookup) {
          // Create a synthetic column representing this relationship
          return [{
            id: `outgoing-${outgoingLookup.name}`,
            name: outgoingLookup.name,
            displayName: `${outgoingLookup.displayName} (from current list)`,
            lookup: { listId: selectedListId },
          } as GraphListColumn];
        }
      }
      return [];
    }
  }, [selectedLinkedListInfo, columns, primaryListId, columnMetadata, selectedListId]);

  const handleSourceChange = useCallback((listId: string) => {
    setSelectedListId(listId);
    // Auto-set lookup column if there's only one option
    const linkedInfo = linkedLists.find(ll => ll.list.id === listId);
    if (linkedInfo?.lookupColumn) {
      setLookupColumn(linkedInfo.lookupColumn);
    } else {
      setLookupColumn('');
    }
    // Auto-set title from list name
    const list = linkedLists.find(ll => ll.list.id === listId);
    if (list && !title) {
      setTitle(list.list.displayName);
    }
    // Reset dependent fields
    setDisplayColumns([]);
    setDefaultSort(undefined);
  }, [linkedLists, title]);

  const handleDisplayColumnToggle = useCallback((col: GraphListColumn) => {
    setDisplayColumns((prev) => {
      const exists = prev.some((dc) => dc.internalName === col.name);
      if (exists) {
        return prev.filter((dc) => dc.internalName !== col.name);
      }
      return [
        ...prev,
        {
          internalName: col.name,
          displayName: col.displayName,
          editable: !col.readOnly,
        },
      ];
    });
  }, []);

  const handleSave = () => {
    const selectedList = linkedLists.find(ll => ll.list.id === selectedListId);
    if (!selectedList) return;

    const updatedSection: RelatedSection = {
      id: section?.id || `section-${Date.now()}`,
      title,
      source: {
        siteId: primarySiteId,
        siteUrl: primarySiteUrl,
        listId: selectedList.list.id,
        listName: selectedList.list.displayName,
      },
      lookupColumn,
      displayColumns,
      defaultSort,
    };

    onSave(updatedSection);
  };

  const canSave = title.trim() && selectedListId && lookupColumn && displayColumns.length > 0;

  return (
    <OverlayDrawer
      open={open}
      onOpenChange={(_, data) => !data.open && onClose()}
      position="end"
      size="medium"
      className={styles.drawerSurface}
    >
      <DrawerHeader>
        <DrawerHeaderTitle
          action={
            <Button
              appearance="subtle"
              icon={<DismissRegular />}
              onClick={onClose}
              aria-label="Close"
            />
          }
        >
          {isEditing ? 'Edit Linked List' : 'Add Linked List'}
        </DrawerHeaderTitle>
      </DrawerHeader>

      <DrawerBody>
        <div className={styles.section}>
          <Field label="Section Title" required>
            <Input
              value={title}
              onChange={(_e, data) => setTitle(data.value)}
              placeholder="e.g., Correspondence"
            />
          </Field>

          <Field label="Linked List" required>
            {loadingLists ? (
              <div className={styles.loadingRow}>
                <Spinner size="tiny" />
                <Text size={200} style={{ color: tokens.colorNeutralForeground2 }}>
                  Loading lists and relationships...
                </Text>
              </div>
            ) : (
              <>
                <Dropdown
                  value={linkedLists.find(ll => ll.list.id === selectedListId)?.list.displayName || ''}
                  selectedOptions={selectedListId ? [selectedListId] : []}
                  onOptionSelect={(_e, data) => handleSourceChange(data.optionValue as string)}
                  placeholder="Select a list"
                >
                  {linkedLists.map((linkedList) => (
                    <Option key={linkedList.list.id} value={linkedList.list.id} text={linkedList.list.displayName}>
                      <div className={styles.listOption}>
                        <span>{linkedList.list.displayName}</span>
                        <Badge
                          size="small"
                          appearance="outline"
                          color={linkedList.direction === 'incoming' ? 'informative' : 'success'}
                        >
                          <span className={styles.directionBadge}>
                            {linkedList.direction === 'incoming' ? (
                              <>
                                <ArrowRightRegular fontSize={10} />
                                links here
                              </>
                            ) : (
                              <>
                                <ArrowLeftRegular fontSize={10} />
                                linked from here
                              </>
                            )}
                          </span>
                        </Badge>
                      </div>
                    </Option>
                  ))}
                </Dropdown>
                {linkedLists.length === 0 && !loadingLists && (
                  <Text size={200} className={styles.warningText} style={{ marginTop: '4px' }}>
                    No lists with lookup relationships found. Lists must have a lookup column that references this list, or this list must have a lookup to another list.
                  </Text>
                )}
              </>
            )}
          </Field>

          {selectedListId && (
            <>
              {loadingColumns ? (
                <div className={styles.loadingRow}>
                  <Spinner size="tiny" />
                  <Text size={200} style={{ color: tokens.colorNeutralForeground2 }}>
                    Loading columns...
                  </Text>
                </div>
              ) : (
                <>
                  <Field label="Link Column" required>
                    <Text size={200} style={{ color: tokens.colorNeutralForeground2, marginBottom: '8px', display: 'block' }}>
                      Column that links the lists together
                    </Text>
                    {lookupColumns.length === 0 ? (
                      <Text size={200} className={styles.warningText}>
                        No valid lookup column found for this relationship.
                      </Text>
                    ) : lookupColumns.length === 1 ? (
                      <Input
                        value={lookupColumns[0].displayName}
                        disabled
                      />
                    ) : (
                      <Dropdown
                        value={lookupColumns.find((c) => c.name === lookupColumn)?.displayName || ''}
                        selectedOptions={lookupColumn ? [lookupColumn] : []}
                        onOptionSelect={(_e, data) => setLookupColumn(data.optionValue as string)}
                        placeholder="Select link column"
                      >
                        {lookupColumns.map((col) => (
                          <Option key={col.id} value={col.name}>
                            {col.displayName}
                          </Option>
                        ))}
                      </Dropdown>
                    )}
                  </Field>

                  <Field label="Display Columns" required>
                    <Text size={200} style={{ color: tokens.colorNeutralForeground2, marginBottom: '8px', display: 'block' }}>
                      Columns to show in the linked items table
                    </Text>
                    <div className={styles.badgeWrap}>
                      {columns
                        .filter((col) => !col.hidden && col.name !== lookupColumn)
                        .slice(0, 15)
                        .map((col) => (
                          <Badge
                            key={col.id}
                            className={styles.badgeItem}
                            appearance={displayColumns.some((dc) => dc.internalName === col.name) ? 'filled' : 'outline'}
                            color={displayColumns.some((dc) => dc.internalName === col.name) ? 'brand' : 'informative'}
                            onClick={() => handleDisplayColumnToggle(col)}
                          >
                            {col.displayName}
                          </Badge>
                        ))}
                    </div>
                    {displayColumns.length === 0 && (
                      <Text size={200} className={styles.warningText} style={{ marginTop: '4px' }}>
                        Select at least one column to display.
                      </Text>
                    )}
                  </Field>

                  {displayColumns.length > 0 && (
                    <Field label="Order By">
                      <div className={styles.sortRow}>
                        <Dropdown
                          className={styles.sortColumn}
                          value={displayColumns.find((c) => c.internalName === defaultSort?.column)?.displayName || ''}
                          selectedOptions={defaultSort?.column ? [defaultSort.column] : []}
                          onOptionSelect={(_e, data) => {
                            if (data.optionValue) {
                              setDefaultSort({
                                column: data.optionValue as string,
                                direction: defaultSort?.direction || 'asc',
                              });
                            } else {
                              setDefaultSort(undefined);
                            }
                          }}
                          placeholder="None"
                        >
                          <Option value="">None</Option>
                          {displayColumns.map((col) => (
                            <Option key={col.internalName} value={col.internalName}>
                              {col.displayName}
                            </Option>
                          ))}
                        </Dropdown>
                        {defaultSort?.column && (
                          <Dropdown
                            className={styles.sortDirection}
                            value={defaultSort.direction === 'asc' ? 'Ascending' : 'Descending'}
                            selectedOptions={[defaultSort.direction]}
                            onOptionSelect={(_e, data) =>
                              setDefaultSort({
                                column: defaultSort.column,
                                direction: data.optionValue as 'asc' | 'desc',
                              })
                            }
                          >
                            <Option value="asc">Ascending</Option>
                            <Option value="desc">Descending</Option>
                          </Dropdown>
                        )}
                      </div>
                    </Field>
                  )}

                </>
              )}
            </>
          )}
        </div>

        {/* Footer */}
        <div className={styles.footer}>
          <Button appearance="secondary" onClick={onClose}>
            Cancel
          </Button>
          <Button appearance="primary" onClick={handleSave} disabled={!canSave}>
            {isEditing ? 'Save Changes' : 'Add Linked List'}
          </Button>
        </div>
      </DrawerBody>
    </OverlayDrawer>
  );
}

export default LinkedListFlyout;
