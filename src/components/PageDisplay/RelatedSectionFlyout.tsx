import { useState, useEffect, useCallback } from 'react';
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
  Checkbox,
  Divider,
  Spinner,
  DrawerBody,
  DrawerHeader,
  DrawerHeaderTitle,
  OverlayDrawer,
} from '@fluentui/react-components';
import { DismissRegular } from '@fluentui/react-icons';
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
  permissionsRow: {
    display: 'flex',
    flexWrap: 'wrap',
    gap: '16px',
    marginTop: '8px',
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
});

interface RelatedSectionFlyoutProps {
  open: boolean;
  section: RelatedSection | null; // null = creating new
  primaryListId: string;
  primarySiteId: string;
  primarySiteUrl: string;
  onClose: () => void;
  onSave: (section: RelatedSection) => void;
}

function RelatedSectionFlyout({
  open,
  section,
  primaryListId,
  primarySiteId,
  primarySiteUrl,
  onClose,
  onSave,
}: RelatedSectionFlyoutProps) {
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
  const [allowCreate, setAllowCreate] = useState(true);
  const [allowEdit, setAllowEdit] = useState(true);
  const [allowDelete, setAllowDelete] = useState(true);

  // Lists loading state (fetched from same site as primary list)
  const [availableLists, setAvailableLists] = useState<GraphList[]>([]);
  const [loadingLists, setLoadingLists] = useState(false);

  // Column loading state
  const [columns, setColumns] = useState<GraphListColumn[]>([]);
  const [loadingColumns, setLoadingColumns] = useState(false);

  // Derived values
  const selectedList = availableLists.find((l) => l.id === selectedListId);

  // Load lists from the same site when drawer opens
  useEffect(() => {
    if (!open || !account || !primarySiteId) {
      return;
    }

    const loadLists = async () => {
      setLoadingLists(true);
      try {
        const lists = await getSiteLists(instance, account, primarySiteId);
        // Filter out system lists and the primary list
        const filtered = lists.filter(
          (list) =>
            !SYSTEM_LIST_NAMES.includes(list.name as typeof SYSTEM_LIST_NAMES[number]) &&
            list.id !== primaryListId
        );
        setAvailableLists(filtered);
      } catch (err) {
        console.error('Failed to load lists:', err);
        setAvailableLists([]);
      } finally {
        setLoadingLists(false);
      }
    };

    loadLists();
  }, [open, instance, account, primarySiteId, primaryListId]);

  // Initialize form when section changes
  useEffect(() => {
    if (section) {
      setTitle(section.title);
      setSelectedListId(section.source.listId || '');
      setLookupColumn(section.lookupColumn);
      setDisplayColumns(section.displayColumns);
      setDefaultSort(section.defaultSort);
      setAllowCreate(section.allowCreate);
      setAllowEdit(section.allowEdit);
      setAllowDelete(section.allowDelete);
    } else {
      // Reset for new section
      setTitle('Related Items');
      setSelectedListId('');
      setLookupColumn('');
      setDisplayColumns([]);
      setDefaultSort(undefined);
      setAllowCreate(true);
      setAllowEdit(true);
      setAllowDelete(true);
    }
  }, [section, open]);

  // Load columns when source list changes
  useEffect(() => {
    if (!account || !selectedList || !primarySiteId) {
      setColumns([]);
      return;
    }

    const loadColumns = async () => {
      setLoadingColumns(true);
      try {
        const cols = await getListColumns(
          instance,
          account,
          primarySiteId,
          selectedList.id
        );
        setColumns(cols);
      } catch (err) {
        console.error('Failed to load columns:', err);
      } finally {
        setLoadingColumns(false);
      }
    };

    loadColumns();
  }, [instance, account, selectedList, primarySiteId]);

  // Get lookup columns that reference the primary list
  const lookupColumns = columns.filter(
    (col) => col.lookup?.listId === primaryListId
  );

  const handleSourceChange = useCallback((listId: string) => {
    setSelectedListId(listId);
    // Reset dependent fields
    setLookupColumn('');
    setDisplayColumns([]);
    setDefaultSort(undefined);
  }, []);

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
    if (!selectedList) return;

    const updatedSection: RelatedSection = {
      id: section?.id || `section-${Date.now()}`,
      title,
      source: {
        siteId: primarySiteId,
        siteUrl: primarySiteUrl,
        listId: selectedList.id,
        listName: selectedList.displayName,
      },
      lookupColumn,
      displayColumns,
      allowCreate,
      allowEdit,
      allowDelete,
      defaultSort,
    };

    onSave(updatedSection);
  };

  const canSave = title.trim() && selectedList && lookupColumn && displayColumns.length > 0;

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
          {isEditing ? 'Edit Related List' : 'Add Related List'}
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

          <Field label="Related List" required>
            {loadingLists ? (
              <div className={styles.loadingRow}>
                <Spinner size="tiny" />
                <Text size={200} style={{ color: tokens.colorNeutralForeground2 }}>
                  Loading lists...
                </Text>
              </div>
            ) : (
              <>
                <Dropdown
                  value={selectedList?.displayName || ''}
                  selectedOptions={selectedListId ? [selectedListId] : []}
                  onOptionSelect={(_e, data) => handleSourceChange(data.optionValue as string)}
                  placeholder="Select a list"
                >
                  {availableLists.map((list) => (
                    <Option key={list.id} value={list.id}>
                      {list.displayName}
                    </Option>
                  ))}
                </Dropdown>
                {availableLists.length === 0 && !loadingLists && (
                  <Text size={200} className={styles.warningText} style={{ marginTop: '4px' }}>
                    No other lists available in this site.
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
                      Column that links to the primary list
                    </Text>
                    {lookupColumns.length === 0 ? (
                      <Text size={200} className={styles.warningText}>
                        No lookup columns found that reference the primary list.
                      </Text>
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
                      Columns to show in the related items table
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

                  <Divider />

                  <Field label="Permissions">
                    <Text size={200} style={{ color: tokens.colorNeutralForeground2, marginBottom: '8px', display: 'block' }}>
                      What actions users can perform on related items
                    </Text>
                    <div className={styles.permissionsRow}>
                      <Checkbox
                        checked={allowCreate}
                        onChange={(_e, data) => setAllowCreate(data.checked === true)}
                        label="Allow Create"
                      />
                      <Checkbox
                        checked={allowEdit}
                        onChange={(_e, data) => setAllowEdit(data.checked === true)}
                        label="Allow Edit"
                      />
                      <Checkbox
                        checked={allowDelete}
                        onChange={(_e, data) => setAllowDelete(data.checked === true)}
                        label="Allow Delete"
                      />
                    </div>
                  </Field>
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
            {isEditing ? 'Save Changes' : 'Add Section'}
          </Button>
        </div>
      </DrawerBody>
    </OverlayDrawer>
  );
}

export default RelatedSectionFlyout;
