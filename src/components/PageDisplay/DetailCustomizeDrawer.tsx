import { useState, useCallback, useEffect, useRef } from 'react';
import {
  makeStyles,
  tokens,
  Text,
  Button,
  Checkbox,
  Dropdown,
  Option,
  Divider,
  DrawerBody,
  DrawerHeader,
  DrawerHeaderTitle,
  OverlayDrawer,
  Spinner,
} from '@fluentui/react-components';
import {
  DismissRegular,
  ReOrderDotsVerticalRegular,
  AddRegular,
  EditRegular,
  DeleteRegular,
} from '@fluentui/react-icons';
import type { PageDefinition, PageColumn, DetailLayoutConfig, DetailColumnSetting, RelatedSection, ListDetailConfig } from '../../types/page';
import type { GraphListColumn } from '../../auth/graphClient';
import LinkedListFlyout from './LinkedListFlyout';

// Section IDs for built-in sections
const DETAILS_SECTION_ID = 'details';
const DESCRIPTION_SECTION_ID = 'description';

// Helper to check if a column is a multiline text column
function isMultilineColumn(internalName: string, columnMetadata?: GraphListColumn[]): boolean {
  if (!columnMetadata) return false;
  const col = columnMetadata.find(c => c.name === internalName);
  return col?.text?.allowMultipleLines === true;
}

// Helper to check if a column is rich text
function isRichTextColumn(internalName: string, columnMetadata?: GraphListColumn[]): boolean {
  if (!columnMetadata) return false;
  const col = columnMetadata.find(c => c.name === internalName);
  return col?.text?.textType === 'richText';
}

// Check if there's a column set to description display style
function hasDescriptionColumn(settings: DetailColumnSetting[]): boolean {
  return settings.some(s => s.visible && s.displayStyle === 'description');
}

const useStyles = makeStyles({
  section: {
    display: 'flex',
    flexDirection: 'column',
    gap: '12px',
    marginBottom: '24px',
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
  columnRow: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
    padding: '8px 12px',
    backgroundColor: tokens.colorNeutralBackground2,
    borderRadius: tokens.borderRadiusMedium,
    cursor: 'grab',
    transition: 'opacity 0.2s, border-color 0.2s',
    border: `1px solid transparent`,
    '&:hover': {
      backgroundColor: tokens.colorNeutralBackground3,
    },
  },
  columnRowDragging: {
    opacity: 0.5,
    border: `1px dashed ${tokens.colorBrandStroke1}`,
  },
  columnCheckbox: {
    flexShrink: 0,
  },
  columnName: {
    flex: 1,
    minWidth: 0,
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap',
  },
  displayStyleDropdown: {
    minWidth: '120px',
  },
  sectionItem: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
    padding: '10px 12px',
    backgroundColor: tokens.colorNeutralBackground2,
    borderRadius: tokens.borderRadiusMedium,
    cursor: 'grab',
    transition: 'opacity 0.2s, border-color 0.2s',
    border: `1px solid transparent`,
    '&:hover': {
      backgroundColor: tokens.colorNeutralBackground3,
    },
  },
  sectionItemDragging: {
    opacity: 0.5,
    border: `1px dashed ${tokens.colorBrandStroke1}`,
  },
  sectionItemNoDrag: {
    cursor: 'default',
    opacity: 0.6,
  },
  dragHandle: {
    color: tokens.colorNeutralForeground3,
    cursor: 'grab',
    flexShrink: 0,
  },
  dragHandleDisabled: {
    opacity: 0.3,
    cursor: 'not-allowed',
  },
  itemName: {
    flex: 1,
  },
  sectionHeader: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    marginBottom: '8px',
  },
  sectionActions: {
    display: 'flex',
    alignItems: 'center',
    gap: '4px',
    flexShrink: 0,
    marginLeft: '8px',
  },
  emptySection: {
    textAlign: 'center',
    padding: '16px',
    color: tokens.colorNeutralForeground2,
    fontSize: tokens.fontSizeBase200,
    backgroundColor: tokens.colorNeutralBackground3,
    borderRadius: tokens.borderRadiusMedium,
  },
  loadingContainer: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    padding: '48px',
  },
});

interface DetailCustomizeDrawerProps {
  // Either provide a page (legacy lookup pages) or listDetailConfig (new per-list config)
  page?: PageDefinition;
  listDetailConfig?: ListDetailConfig;
  // Column metadata for determining column types (multiline, rich text, etc.)
  columnMetadata?: GraphListColumn[];
  // Optional title column override
  titleColumn?: string;
  open: boolean;
  // Show loading spinner while columns are being refreshed
  loading?: boolean;
  onClose: () => void;
  // Called immediately on every change for real-time updates
  onChange: (config: DetailLayoutConfig, relatedSections?: RelatedSection[]) => void;
}

function DetailCustomizeDrawer({ page, listDetailConfig, columnMetadata, titleColumn: titleColumnProp, open, loading, onClose, onChange }: DetailCustomizeDrawerProps) {
  const styles = useStyles();

  // Get display columns from either source
  const displayColumns: PageColumn[] = listDetailConfig?.displayColumns ?? page?.displayColumns ?? [];

  // Get existing detail layout
  const existingDetailLayout: DetailLayoutConfig | undefined = listDetailConfig?.detailLayout ?? page?.detailLayout;

  // Get existing linked lists (formerly related sections)
  const existingLinkedLists: RelatedSection[] = listDetailConfig?.relatedSections ?? page?.relatedSections ?? [];

  // Get primary list/site info for linked list flyout
  const primaryListId = listDetailConfig?.listId ?? page?.primarySource?.listId ?? '';
  const primarySiteId = listDetailConfig?.siteId ?? page?.primarySource?.siteId ?? '';
  const primarySiteUrl = listDetailConfig?.siteUrl ?? page?.primarySource?.siteUrl ?? '';

  // Get the title column - first table column, first display column, or fallback to Title
  // This column is always shown in header, excluded from customization
  const titleColumn = titleColumnProp
    ?? page?.searchConfig?.tableColumns?.[0]?.internalName
    ?? displayColumns[0]?.internalName
    ?? 'Title';

  // Initialize column settings from existing config or defaults (excluding title column)
  const [columnSettings, setColumnSettings] = useState<DetailColumnSetting[]>(() => {
    // Filter out the title column from display columns
    const nonTitleColumns = displayColumns.filter(col => col.internalName !== titleColumn);

    // Build a map of existing settings for quick lookup (deduplicated)
    const existingSettingsMap = new Map<string, DetailColumnSetting>();
    if (existingDetailLayout?.columnSettings) {
      for (const setting of existingDetailLayout.columnSettings) {
        if (setting.internalName !== titleColumn && !existingSettingsMap.has(setting.internalName)) {
          existingSettingsMap.set(setting.internalName, setting);
        }
      }
    }

    // Build settings based on displayColumns, preserving existing settings where available
    return nonTitleColumns.map(col => {
      const existing = existingSettingsMap.get(col.internalName);
      if (existing) {
        return existing;
      }
      // New columns default to not visible
      return {
        internalName: col.internalName,
        visible: false,
        displayStyle: 'list' as const,
      };
    });
  });

  // Initialize linked lists (formerly related sections)
  const [linkedLists, setLinkedLists] = useState<RelatedSection[]>(() => {
    return [...existingLinkedLists];
  });

  // Initialize section order (details, description, linked list IDs)
  const [sectionOrder, setSectionOrder] = useState<string[]>(() => {
    // Check for new sectionOrder first, then legacy relatedSectionOrder
    if (existingDetailLayout?.sectionOrder) {
      // Validate and merge with current linked lists
      const existingOrder = existingDetailLayout.sectionOrder;
      const linkedListIds = existingLinkedLists.map(s => s.id);

      // Build valid order: keep valid IDs, add missing linked lists at end
      const validOrder = existingOrder.filter(id =>
        id === DETAILS_SECTION_ID ||
        id === DESCRIPTION_SECTION_ID ||
        linkedListIds.includes(id)
      );

      // Add any new linked lists not in the order
      const newListIds = linkedListIds.filter(id => !validOrder.includes(id));

      // Ensure details and description are present
      if (!validOrder.includes(DETAILS_SECTION_ID)) {
        validOrder.unshift(DETAILS_SECTION_ID);
      }
      if (!validOrder.includes(DESCRIPTION_SECTION_ID)) {
        const detailsIndex = validOrder.indexOf(DETAILS_SECTION_ID);
        validOrder.splice(detailsIndex + 1, 0, DESCRIPTION_SECTION_ID);
      }

      return [...validOrder, ...newListIds];
    }

    // Legacy: convert relatedSectionOrder to new format
    if (existingDetailLayout?.relatedSectionOrder) {
      const linkedListIds = existingDetailLayout.relatedSectionOrder.filter(id =>
        existingLinkedLists.some(s => s.id === id)
      );
      const newListIds = existingLinkedLists
        .filter(s => !linkedListIds.includes(s.id))
        .map(s => s.id);
      return [DETAILS_SECTION_ID, DESCRIPTION_SECTION_ID, ...linkedListIds, ...newListIds];
    }

    // Default: details, description, then all linked lists
    return [DETAILS_SECTION_ID, DESCRIPTION_SECTION_ID, ...existingLinkedLists.map(s => s.id)];
  });

  // Drag state for columns and sections
  const [draggedColIndex, setDraggedColIndex] = useState<number | null>(null);
  const [draggedSectionIndex, setDraggedSectionIndex] = useState<number | null>(null);

  // Flyout state for add/edit linked list
  const [flyoutOpen, setFlyoutOpen] = useState(false);
  const [editingLinkedList, setEditingLinkedList] = useState<RelatedSection | null>(null);

  // Track if this is the initial mount (to avoid triggering onChange on mount)
  const isInitialMount = useRef(true);
  const prevLoadingRef = useRef(loading);

  // Sync column settings when loading completes (columns were refreshed)
  useEffect(() => {
    const wasLoading = prevLoadingRef.current;
    prevLoadingRef.current = loading;

    // Only sync when loading transitions from true to false
    if (wasLoading && !loading) {
      const nonTitleColumns = displayColumns.filter(col => col.internalName !== titleColumn);

      // Build a map of existing settings for quick lookup (deduplicated)
      const existingSettingsMap = new Map<string, DetailColumnSetting>();
      if (existingDetailLayout?.columnSettings) {
        for (const setting of existingDetailLayout.columnSettings) {
          if (setting.internalName !== titleColumn && !existingSettingsMap.has(setting.internalName)) {
            existingSettingsMap.set(setting.internalName, setting);
          }
        }
      }

      // Build settings based on displayColumns, preserving existing settings where available
      const newSettings = nonTitleColumns.map(col => {
        const existing = existingSettingsMap.get(col.internalName);
        if (existing) {
          return existing;
        }
        // New columns default to not visible
        return {
          internalName: col.internalName,
          visible: false,
          displayStyle: 'list' as const,
        };
      });

      setColumnSettings(newSettings);

      // Also sync linked lists
      setLinkedLists([...existingLinkedLists]);
    }
  }, [loading, displayColumns, existingDetailLayout, existingLinkedLists, titleColumn]);

  // Keep onChange in a ref to avoid triggering effect when callback identity changes
  const onChangeRef = useRef(onChange);
  useEffect(() => {
    onChangeRef.current = onChange;
  }, [onChange]);

  // Debounced auto-save effect: prevents rapid-fire saves that cause SharePoint conflicts
  useEffect(() => {
    if (isInitialMount.current) {
      isInitialMount.current = false;
      return;
    }

    // Debounce saves by 300ms to batch rapid changes (e.g., during drag reordering)
    const timeoutId = setTimeout(() => {
      onChangeRef.current(
        { columnSettings, sectionOrder },
        linkedLists
      );
    }, 300);

    return () => clearTimeout(timeoutId);
  }, [columnSettings, linkedLists, sectionOrder]);

  const handleColumnVisibilityChange = useCallback((internalName: string, visible: boolean) => {
    setColumnSettings(prev => prev.map(s =>
      s.internalName === internalName ? { ...s, visible } : s
    ));
  }, []);

  const handleColumnStyleChange = useCallback((internalName: string, displayStyle: 'stat' | 'list' | 'description') => {
    setColumnSettings(prev => {
      // If setting a column to 'description', reset any other column that was 'description' to 'list'
      if (displayStyle === 'description') {
        return prev.map(s => {
          if (s.internalName === internalName) {
            return { ...s, displayStyle };
          }
          if (s.displayStyle === 'description') {
            return { ...s, displayStyle: 'list' as const };
          }
          return s;
        });
      }
      return prev.map(s =>
        s.internalName === internalName ? { ...s, displayStyle } : s
      );
    });
  }, []);

  const handleColumnReorder = useCallback((fromIndex: number, toIndex: number) => {
    if (fromIndex === toIndex) return;
    setColumnSettings(prev => {
      const newSettings = [...prev];
      const [removed] = newSettings.splice(fromIndex, 1);
      newSettings.splice(toIndex, 0, removed);
      return newSettings;
    });
  }, []);

  const handleSectionReorder = useCallback((fromIndex: number, toIndex: number) => {
    if (fromIndex === toIndex) return;
    setSectionOrder(prev => {
      const newOrder = [...prev];
      const [removed] = newOrder.splice(fromIndex, 1);
      newOrder.splice(toIndex, 0, removed);
      return newOrder;
    });
  }, []);

  // Linked list management
  const handleAddLinkedList = useCallback(() => {
    setEditingLinkedList(null);
    setFlyoutOpen(true);
  }, []);

  const handleEditLinkedList = useCallback((section: RelatedSection) => {
    setEditingLinkedList(section);
    setFlyoutOpen(true);
  }, []);

  const handleRemoveLinkedList = useCallback((sectionId: string) => {
    setLinkedLists(prev => prev.filter(s => s.id !== sectionId));
    setSectionOrder(prev => prev.filter(id => id !== sectionId));
  }, []);

  const handleSaveLinkedList = useCallback((section: RelatedSection) => {
    setLinkedLists(prev => {
      const existingIndex = prev.findIndex(s => s.id === section.id);
      if (existingIndex >= 0) {
        // Update existing
        const newSections = [...prev];
        newSections[existingIndex] = section;
        return newSections;
      } else {
        // Add new
        return [...prev, section];
      }
    });
    // Add to section order if new
    setSectionOrder(prev => {
      if (!prev.includes(section.id)) {
        return [...prev, section.id];
      }
      return prev;
    });
    setFlyoutOpen(false);
  }, []);

  // Get display name for a column
  const getColumnDisplayName = (internalName: string): string => {
    const col = displayColumns.find(c => c.internalName === internalName);
    return col?.displayName || internalName;
  };

  // Get linked list by ID (from local state)
  const getLinkedListById = (id: string): RelatedSection | undefined => {
    return linkedLists.find(s => s.id === id);
  };

  // Get section display name
  const getSectionDisplayName = (sectionId: string): string => {
    if (sectionId === DETAILS_SECTION_ID) return 'Details';
    if (sectionId === DESCRIPTION_SECTION_ID) return 'Description';
    const linkedList = getLinkedListById(sectionId);
    return linkedList?.title || 'Unknown Section';
  };

  // Check if section is a linked list
  const isLinkedListSection = (sectionId: string): boolean => {
    return sectionId !== DETAILS_SECTION_ID && sectionId !== DESCRIPTION_SECTION_ID;
  };

  // Check if description section should be shown (has a description column)
  const showDescriptionSection = hasDescriptionColumn(columnSettings);

  // Check if a section should be visible in the reorder UI
  const isSectionVisible = (sectionId: string): boolean => {
    if (sectionId === DETAILS_SECTION_ID) return true;
    if (sectionId === DESCRIPTION_SECTION_ID) return showDescriptionSection;
    return linkedLists.some(s => s.id === sectionId);
  };

  return (
  <>
    <OverlayDrawer
      open={open}
      onOpenChange={(_, data) => !data.open && onClose()}
      position="end"
      size="medium"
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
          Customize Layout
        </DrawerHeaderTitle>
      </DrawerHeader>

      <DrawerBody>
        {loading ? (
          <div className={styles.loadingContainer}>
            <Spinner size="medium" label="Loading columns..." />
          </div>
        ) : (
        <>
        {/* Detail Columns Section */}
        <div className={styles.section} style={{ marginTop: '8px' }}>
          <Text className={styles.sectionTitle}>Detail Columns</Text>
          <Text className={styles.sectionHint}>
            Configure which columns appear in Details and their display style
          </Text>

          {columnSettings.map((setting, index) => {
            const displayName = getColumnDisplayName(setting.internalName);
            return (
              <div
                key={setting.internalName}
                draggable
                onDragStart={() => setDraggedColIndex(index)}
                onDragEnd={() => setDraggedColIndex(null)}
                onDragOver={(e) => {
                  e.preventDefault();
                  if (draggedColIndex !== null && draggedColIndex !== index) {
                    handleColumnReorder(draggedColIndex, index);
                    setDraggedColIndex(index);
                  }
                }}
                className={`${styles.columnRow} ${
                  draggedColIndex === index ? styles.columnRowDragging : ''
                }`}
              >
                <ReOrderDotsVerticalRegular className={styles.dragHandle} />
                <Checkbox
                  className={styles.columnCheckbox}
                  checked={setting.visible}
                  onChange={(_, data) =>
                    handleColumnVisibilityChange(setting.internalName, !!data.checked)
                  }
                />
                <Text className={styles.columnName}>{displayName}</Text>
                {setting.visible && (
                  <Dropdown
                    className={styles.displayStyleDropdown}
                    value={
                      setting.displayStyle === 'stat'
                        ? 'Stat Box'
                        : setting.displayStyle === 'description'
                        ? 'Description'
                        : 'Detail List'
                    }
                    selectedOptions={[setting.displayStyle]}
                    onOptionSelect={(_, data) =>
                      handleColumnStyleChange(
                        setting.internalName,
                        data.optionValue as 'stat' | 'list' | 'description'
                      )
                    }
                    size="small"
                  >
                    <Option value="stat">Stat Box</Option>
                    <Option value="list">Detail List</Option>
                    {isMultilineColumn(setting.internalName, columnMetadata) && (
                      <Option text={`Description${isRichTextColumn(setting.internalName, columnMetadata) ? ' (Rich)' : ''}`} value="description">
                        Description{isRichTextColumn(setting.internalName, columnMetadata) ? ' (Rich)' : ''}
                      </Option>
                    )}
                  </Dropdown>
                )}
              </div>
            );
          })}
        </div>

        <Divider />

        {/* Sections - Drag to reorder */}
        <div className={styles.section} style={{ marginTop: '24px' }}>
          <Text className={styles.sectionTitle}>Customize Layout</Text>
          <Text className={styles.sectionHint}>
            Drag to reorder sections. Add linked lists to show related data.
          </Text>

          {sectionOrder.map((sectionId, index) => {
            // Skip invisible sections but keep indices aligned with sectionOrder
            if (!isSectionVisible(sectionId)) return null;

            const isLinkedList = isLinkedListSection(sectionId);
            const linkedList = isLinkedList ? getLinkedListById(sectionId) : null;

            return (
              <div
                key={sectionId}
                draggable
                onDragStart={() => setDraggedSectionIndex(index)}
                onDragEnd={() => setDraggedSectionIndex(null)}
                onDragOver={(e) => {
                  e.preventDefault();
                  if (draggedSectionIndex !== null && draggedSectionIndex !== index) {
                    handleSectionReorder(draggedSectionIndex, index);
                    setDraggedSectionIndex(index);
                  }
                }}
                className={`${styles.sectionItem} ${
                  draggedSectionIndex === index ? styles.sectionItemDragging : ''
                }`}
              >
                <ReOrderDotsVerticalRegular className={styles.dragHandle} />
                <Text className={styles.itemName}>{getSectionDisplayName(sectionId)}</Text>
                {isLinkedList && linkedList && (
                  <div className={styles.sectionActions}>
                    <Button
                      appearance="subtle"
                      size="small"
                      icon={<EditRegular />}
                      onClick={(e) => {
                        e.stopPropagation();
                        handleEditLinkedList(linkedList);
                      }}
                      title="Edit"
                    />
                    <Button
                      appearance="subtle"
                      size="small"
                      icon={<DeleteRegular />}
                      onClick={(e) => {
                        e.stopPropagation();
                        handleRemoveLinkedList(sectionId);
                      }}
                      title="Remove"
                    />
                  </div>
                )}
              </div>
            );
          })}

          {/* Add Linked List button */}
          <Button
            appearance="subtle"
            icon={<AddRegular />}
            onClick={handleAddLinkedList}
            style={{ alignSelf: 'flex-start', marginTop: '4px' }}
          >
            Add Linked List
          </Button>
        </div>
        </>
        )}
      </DrawerBody>
    </OverlayDrawer>

    {/* Linked List Flyout */}
    <LinkedListFlyout
      open={flyoutOpen}
      section={editingLinkedList}
      primaryListId={primaryListId}
      primarySiteId={primarySiteId}
      primarySiteUrl={primarySiteUrl}
      columnMetadata={columnMetadata}
      onClose={() => setFlyoutOpen(false)}
      onSave={handleSaveLinkedList}
    />
  </>
  );
}

export default DetailCustomizeDrawer;
