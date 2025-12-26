import { useState, useEffect, useCallback, useMemo } from 'react';
import { useMsal } from '@azure/msal-react';
import {
  makeStyles,
  tokens,
  Text,
  Button,
  Divider,
  DrawerBody,
  DrawerHeader,
  DrawerHeaderTitle,
  OverlayDrawer,
  Input,
  Textarea,
  Field,
  Checkbox,
  Spinner,
  Badge,
  Link,
  mergeClasses,
  Dropdown,
  Option,
} from '@fluentui/react-components';
import {
  DismissRegular,
  ReOrderDotsVerticalRegular,
} from '@fluentui/react-icons';
import type {
  PageDefinition,
  PageSource,
  PageColumn,
  SearchConfig,
  FilterColumn,
  WebPartDataSource,
} from '../../types/page';
import { getListColumns, type GraphListColumn } from '../../auth/graphClient';
import { useSettings } from '../../contexts/SettingsContext';
import { IconPicker } from '../PageEditor/IconPicker';
import { DEFAULT_PAGE_ICONS } from '../../utils/iconMap';
import DataSourcePicker from './WebParts/DataSourcePicker';

interface ColumnWithMeta extends GraphListColumn {
  sourceListId: string;
  sourceListName: string;
}

const useStyles = makeStyles({
  body: {
    display: 'flex',
    flexDirection: 'column',
    height: '100%',
  },
  content: {
    flex: 1,
    overflowY: 'auto',
    paddingBottom: '16px',
  },
  section: {
    display: 'flex',
    flexDirection: 'column',
    gap: '12px',
    marginBottom: '24px',
  },
  sectionHeader: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    marginBottom: '8px',
  },
  sectionTitle: {
    fontSize: tokens.fontSizeBase400,
    fontWeight: tokens.fontWeightSemibold,
  },
  sectionHint: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground3,
    marginBottom: '8px',
  },
  columnsGrid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(2, 1fr)',
    gap: '16px',
  },
  columnPanel: {
    display: 'flex',
    flexDirection: 'column',
    gap: '8px',
  },
  columnPanelHeader: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    marginBottom: '4px',
  },
  columnPanelTitle: {
    fontWeight: tokens.fontWeightMedium,
    fontSize: tokens.fontSizeBase200,
    textTransform: 'uppercase',
    letterSpacing: '0.02em',
    color: tokens.colorNeutralForeground2,
  },
  toggleLink: {
    fontSize: tokens.fontSizeBase100,
  },
  columnList: {
    display: 'flex',
    flexDirection: 'column',
    gap: '4px',
    maxHeight: '200px',
    overflowY: 'auto',
  },
  columnItem: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
    padding: '8px 10px',
    borderRadius: tokens.borderRadiusMedium,
    backgroundColor: tokens.colorNeutralBackground3,
    cursor: 'pointer',
    transition: 'background-color 0.1s ease',
    '&:hover': {
      backgroundColor: tokens.colorNeutralBackground4,
    },
  },
  columnItemSelected: {
    backgroundColor: tokens.colorBrandBackground2,
    cursor: 'move',
    '&:hover': {
      backgroundColor: tokens.colorBrandBackground2,
    },
  },
  columnItemDragging: {
    opacity: 0.5,
  },
  columnItemHidden: {
    opacity: 0.6,
    fontStyle: 'italic',
  },
  dragHandle: {
    color: tokens.colorNeutralForeground3,
  },
  checkboxList: {
    display: 'flex',
    flexDirection: 'column',
    gap: '8px',
  },
  checkboxItem: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
  },
  loadingContainer: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    padding: '24px',
  },
  footer: {
    display: 'flex',
    justifyContent: 'flex-end',
    gap: '8px',
    padding: '16px 0',
    borderTop: `1px solid ${tokens.colorNeutralStroke1}`,
    marginTop: 'auto',
  },
});

interface LookupCustomizeDrawerProps {
  page: PageDefinition;
  open: boolean;
  onClose: () => void;
  onSave: (page: PageDefinition) => Promise<void>;
}

export default function LookupCustomizeDrawer({
  page,
  open,
  onClose,
  onSave,
}: LookupCustomizeDrawerProps) {
  const styles = useStyles();
  const { instance, accounts } = useMsal();
  const account = accounts[0];
  const { sections } = useSettings();

  // Form state
  const [name, setName] = useState(page.name);
  const [description, setDescription] = useState(page.description || '');
  const [icon, setIcon] = useState(page.icon || DEFAULT_PAGE_ICONS.lookup);
  const [sectionId, setSectionId] = useState<string | null>(page.sectionId || null);
  const [primarySource, setPrimarySource] = useState<PageSource>(page.primarySource);
  const [displayColumns, setDisplayColumns] = useState<PageColumn[]>(page.displayColumns);
  const [searchConfig, setSearchConfig] = useState<SearchConfig>(page.searchConfig);

  // Sorted sections for dropdown
  const sortedSections = useMemo(() => {
    return Object.values(sections).sort((a, b) => a.order - b.order);
  }, [sections]);

  // Available columns from primary source
  const [availableColumns, setAvailableColumns] = useState<ColumnWithMeta[]>([]);
  const [loadingColumns, setLoadingColumns] = useState(false);
  const [showHiddenColumns, setShowHiddenColumns] = useState(false);

  // Drag and drop state
  const [draggedColIndex, setDraggedColIndex] = useState<number | null>(null);

  // Saving state
  const [saving, setSaving] = useState(false);

  // Sync state when drawer opens or page changes
  useEffect(() => {
    if (open) {
      setName(page.name);
      setDescription(page.description || '');
      setIcon(page.icon || DEFAULT_PAGE_ICONS.lookup);
      setSectionId(page.sectionId || null);
      setPrimarySource(page.primarySource);
      setDisplayColumns(page.displayColumns);
      setSearchConfig(page.searchConfig);
    }
  }, [open, page]);

  // Load columns when primary source changes
  useEffect(() => {
    if (!account || !primarySource?.siteId || !primarySource?.listId) {
      setAvailableColumns([]);
      return;
    }

    const loadColumns = async () => {
      setLoadingColumns(true);
      try {
        const cols = await getListColumns(
          instance,
          account,
          primarySource.siteId,
          primarySource.listId
        );
        setAvailableColumns(
          cols.map((col) => ({
            ...col,
            sourceListId: primarySource.listId,
            sourceListName: primarySource.listName,
          }))
        );
      } catch (err) {
        console.error('Failed to load columns:', err);
      } finally {
        setLoadingColumns(false);
      }
    };

    loadColumns();
  }, [instance, account, primarySource?.siteId, primarySource?.listId, primarySource?.listName]);

  const handlePrimarySourceChange = useCallback(
    (source: WebPartDataSource) => {
      // Only clear columns if the list actually changed
      if (source.listId !== primarySource?.listId) {
        setDisplayColumns([]);
        setSearchConfig({
          textSearchColumns: [],
          filterColumns: [],
        });
      }
      setPrimarySource({
        siteId: source.siteId,
        siteUrl: source.siteUrl,
        listId: source.listId,
        listName: source.listName,
      });
    },
    [primarySource?.listId]
  );

  const handleColumnToggle = useCallback((col: ColumnWithMeta) => {
    setDisplayColumns((prev) => {
      const exists = prev.some((c) => c.internalName === col.name);
      if (exists) {
        return prev.filter((c) => c.internalName !== col.name);
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

  const handleColumnReorder = useCallback((fromIndex: number, toIndex: number) => {
    setDisplayColumns((prev) => {
      const newCols = [...prev];
      const [moved] = newCols.splice(fromIndex, 1);
      newCols.splice(toIndex, 0, moved);
      return newCols;
    });
    setDraggedColIndex(toIndex);
  }, []);

  const handleSearchColumnToggle = useCallback((colName: string) => {
    setSearchConfig((prev) => {
      const exists = prev.textSearchColumns.includes(colName);
      return {
        ...prev,
        textSearchColumns: exists
          ? prev.textSearchColumns.filter((c) => c !== colName)
          : [...prev.textSearchColumns, colName],
      };
    });
  }, []);

  const handleFilterColumnToggle = useCallback((col: ColumnWithMeta) => {
    setSearchConfig((prev) => {
      const exists = prev.filterColumns.some((f) => f.internalName === col.name);
      if (exists) {
        return {
          ...prev,
          filterColumns: prev.filterColumns.filter((f) => f.internalName !== col.name),
        };
      }

      // Determine filter type based on column
      let type: FilterColumn['type'] = 'choice';
      if (col.lookup) {
        type = 'lookup';
      } else if (
        col.name === 'Boolean' ||
        col.displayName?.toLowerCase().includes('yes') ||
        col.displayName?.toLowerCase().includes('no')
      ) {
        type = 'boolean';
      }

      return {
        ...prev,
        filterColumns: [
          ...prev.filterColumns,
          {
            internalName: col.name,
            displayName: col.displayName,
            type,
          },
        ],
      };
    });
  }, []);

  // Save handler
  const handleSave = useCallback(async () => {
    setSaving(true);
    try {
      const updatedPage: PageDefinition = {
        ...page,
        name,
        description: description || undefined,
        icon,
        sectionId,
        primarySource,
        displayColumns,
        searchConfig: {
          ...searchConfig,
          tableColumns: displayColumns,
        },
      };
      await onSave(updatedPage);
      onClose();
    } finally {
      setSaving(false);
    }
  }, [page, name, description, icon, sectionId, primarySource, displayColumns, searchConfig, onSave, onClose]);

  // Get choice columns for filter dropdown
  const choiceColumns = availableColumns.filter(
    (col) => col.choice || col.lookup || col.name === 'Boolean'
  );

  // Filter available columns based on hidden toggle
  const visibleAvailableColumns = availableColumns.filter((col) => {
    if (showHiddenColumns) return true;
    return !col.hidden;
  });

  return (
    <OverlayDrawer
      position="end"
      size="medium"
      open={open}
      onOpenChange={(_, { open: isOpen }) => {
        if (!isOpen) onClose();
      }}
    >
      <DrawerHeader>
        <DrawerHeaderTitle
          action={
            <Button
              appearance="subtle"
              aria-label="Close"
              icon={<DismissRegular />}
              onClick={onClose}
            />
          }
        >
          Customize Page
        </DrawerHeaderTitle>
      </DrawerHeader>

      <DrawerBody className={styles.body}>
        <div className={styles.content}>
          {/* Basic Info Section */}
          <div className={styles.section}>
            <Text className={styles.sectionTitle}>Basic Information</Text>

            <Field label="Page Name" required>
              <Input
                value={name}
                onChange={(_e, data) => setName(data.value)}
              />
            </Field>

            <Field label="Description">
              <Textarea
                value={description}
                onChange={(_e, data) => setDescription(data.value)}
                rows={2}
              />
            </Field>

            <Field label="Icon">
              <IconPicker value={icon} onChange={setIcon} />
            </Field>

            <Field label="Sidebar Section">
              <Dropdown
                value={sectionId ? sections[sectionId]?.name : 'None'}
                selectedOptions={sectionId ? [sectionId] : []}
                onOptionSelect={(_, data) => {
                  setSectionId(data.optionValue === '' ? null : data.optionValue || null);
                }}
              >
                <Option value="">None (show at top)</Option>
                {sortedSections.map((section) => (
                  <Option key={section.id} value={section.id}>
                    {section.name}
                  </Option>
                ))}
              </Dropdown>
            </Field>
          </div>

          <Divider />

          {/* Data Source Section */}
          <div className={styles.section} style={{ marginTop: '24px' }}>
            <Text className={styles.sectionTitle}>Data Source</Text>
            <Text className={styles.sectionHint}>
              The SharePoint list this page displays data from.
            </Text>

            <DataSourcePicker
              value={
                primarySource
                  ? {
                      siteId: primarySource.siteId,
                      siteUrl: primarySource.siteUrl,
                      listId: primarySource.listId,
                      listName: primarySource.listName,
                    }
                  : undefined
              }
              onChange={handlePrimarySourceChange}
            />
          </div>

          <Divider />

          {/* Display Columns Section */}
          <div className={styles.section} style={{ marginTop: '24px' }}>
            <Text className={styles.sectionTitle}>Display Columns</Text>
            <Text className={styles.sectionHint}>
              Select which columns to show in the table. Drag to reorder.
            </Text>

            {loadingColumns ? (
              <div className={styles.loadingContainer}>
                <Spinner size="small" />
              </div>
            ) : availableColumns.length === 0 ? (
              <Text size={200} style={{ color: tokens.colorNeutralForeground3 }}>
                Select a data source to see available columns.
              </Text>
            ) : (
              <div className={styles.columnsGrid}>
                {/* Available Columns */}
                <div className={styles.columnPanel}>
                  <div className={styles.columnPanelHeader}>
                    <Text className={styles.columnPanelTitle}>Available</Text>
                    <Link
                      className={styles.toggleLink}
                      onClick={() => setShowHiddenColumns(!showHiddenColumns)}
                    >
                      {showHiddenColumns ? 'Hide hidden' : 'Show hidden'}
                    </Link>
                  </div>
                  <div className={styles.columnList}>
                    {visibleAvailableColumns
                      .filter((col) => !displayColumns.some((dc) => dc.internalName === col.name))
                      .map((col) => (
                        <div
                          key={col.id}
                          className={mergeClasses(
                            styles.columnItem,
                            col.hidden && styles.columnItemHidden
                          )}
                          onClick={() => handleColumnToggle(col)}
                        >
                          <Text size={200}>{col.displayName}</Text>
                          {col.readOnly && (
                            <Badge appearance="outline" size="small">
                              read-only
                            </Badge>
                          )}
                        </div>
                      ))}
                  </div>
                </div>

                {/* Selected Columns */}
                <div className={styles.columnPanel}>
                  <div className={styles.columnPanelHeader}>
                    <Text className={styles.columnPanelTitle}>
                      Selected ({displayColumns.length})
                    </Text>
                  </div>
                  <div className={styles.columnList}>
                    {displayColumns.map((col, index) => (
                      <div
                        key={col.internalName}
                        draggable
                        onDragStart={() => setDraggedColIndex(index)}
                        onDragEnd={() => setDraggedColIndex(null)}
                        onDragOver={(e) => {
                          e.preventDefault();
                          if (draggedColIndex !== null && draggedColIndex !== index) {
                            handleColumnReorder(draggedColIndex, index);
                          }
                        }}
                        className={mergeClasses(
                          styles.columnItem,
                          styles.columnItemSelected,
                          draggedColIndex === index && styles.columnItemDragging
                        )}
                      >
                        <ReOrderDotsVerticalRegular className={styles.dragHandle} />
                        <Text size={200} style={{ flex: 1 }}>
                          {col.displayName}
                        </Text>
                        <Button
                          appearance="subtle"
                          size="small"
                          icon={<DismissRegular />}
                          onClick={(e) => {
                            e.stopPropagation();
                            const fullCol = availableColumns.find(
                              (ac) => ac.name === col.internalName
                            );
                            if (fullCol) handleColumnToggle(fullCol);
                          }}
                        />
                      </div>
                    ))}
                    {displayColumns.length === 0 && (
                      <Text
                        size={200}
                        style={{ color: tokens.colorNeutralForeground3, padding: '8px' }}
                      >
                        Click columns on the left to add them
                      </Text>
                    )}
                  </div>
                </div>
              </div>
            )}
          </div>

          <Divider />

          {/* Search & Filters Section */}
          <div className={styles.section} style={{ marginTop: '24px' }}>
            <Text className={styles.sectionTitle}>Search & Filters</Text>

            {/* Text Search Columns */}
            <Field label="Text Search Columns">
              <Text
                size={200}
                style={{
                  color: tokens.colorNeutralForeground2,
                  marginBottom: '8px',
                  display: 'block',
                }}
              >
                Which columns should the search box look through?
              </Text>
              <div className={styles.checkboxList}>
                {displayColumns.map((col) => (
                  <Checkbox
                    key={col.internalName}
                    checked={searchConfig.textSearchColumns.includes(col.internalName)}
                    onChange={() => handleSearchColumnToggle(col.internalName)}
                    label={col.displayName}
                  />
                ))}
              </div>
              {displayColumns.length === 0 && (
                <Text size={200} style={{ color: tokens.colorNeutralForeground3 }}>
                  Select display columns first.
                </Text>
              )}
            </Field>

            {/* Dropdown Filters */}
            <Field label="Dropdown Filters" style={{ marginTop: '16px' }}>
              <Text
                size={200}
                style={{
                  color: tokens.colorNeutralForeground2,
                  marginBottom: '8px',
                  display: 'block',
                }}
              >
                Choice columns that appear as filter dropdowns.
              </Text>
              {choiceColumns.length === 0 ? (
                <Text size={200} style={{ color: tokens.colorNeutralForeground3 }}>
                  No choice or lookup columns available.
                </Text>
              ) : (
                <div className={styles.checkboxList}>
                  {choiceColumns.map((col) => {
                    const isChecked = searchConfig.filterColumns.some(
                      (f) => f.internalName === col.name
                    );
                    return (
                      <div key={col.id} className={styles.checkboxItem}>
                        <Checkbox
                          checked={isChecked}
                          onChange={() => handleFilterColumnToggle(col)}
                          label={col.displayName}
                        />
                        <Badge appearance="outline" size="small">
                          {col.lookup ? 'lookup' : col.choice ? 'choice' : 'boolean'}
                        </Badge>
                      </div>
                    );
                  })}
                </div>
              )}
            </Field>
          </div>
        </div>

        {/* Footer with Save/Cancel */}
        <div className={styles.footer}>
          <Button appearance="secondary" onClick={onClose} disabled={saving}>
            Cancel
          </Button>
          <Button
            appearance="primary"
            onClick={handleSave}
            disabled={saving || !name.trim()}
          >
            {saving ? 'Saving...' : 'Save'}
          </Button>
        </div>
      </DrawerBody>
    </OverlayDrawer>
  );
}
