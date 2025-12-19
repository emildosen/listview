import { useState, useCallback } from 'react';
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
import RelatedSectionFlyout from './RelatedSectionFlyout';

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
  dragHandle: {
    color: tokens.colorNeutralForeground3,
    cursor: 'grab',
    flexShrink: 0,
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
  footer: {
    display: 'flex',
    justifyContent: 'flex-end',
    gap: '8px',
    padding: '16px 0',
    borderTop: `1px solid ${tokens.colorNeutralStroke1}`,
    marginTop: 'auto',
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
  onClose: () => void;
  onSave: (config: DetailLayoutConfig, relatedSections?: RelatedSection[]) => Promise<void>;
}

function DetailCustomizeDrawer({ page, listDetailConfig, columnMetadata, titleColumn: titleColumnProp, open, onClose, onSave }: DetailCustomizeDrawerProps) {
  const styles = useStyles();

  // Get display columns from either source
  const displayColumns: PageColumn[] = listDetailConfig?.displayColumns ?? page?.displayColumns ?? [];

  // Get existing detail layout
  const existingDetailLayout: DetailLayoutConfig | undefined = listDetailConfig?.detailLayout ?? page?.detailLayout;

  // Get existing related sections
  const existingRelatedSections: RelatedSection[] = listDetailConfig?.relatedSections ?? page?.relatedSections ?? [];

  // Get primary list/site info for related section flyout
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

    if (existingDetailLayout?.columnSettings) {
      // Preserve existing order and merge with any new columns
      const existingSettings = existingDetailLayout.columnSettings.filter(
        s => s.internalName !== titleColumn
      );
      const existingNames = new Set(existingSettings.map(s => s.internalName));

      // Keep existing settings in their order, add new columns at the end
      const newColumns = nonTitleColumns
        .filter(col => !existingNames.has(col.internalName))
        .map(col => ({
          internalName: col.internalName,
          visible: true,
          displayStyle: 'list' as const,
        }));

      // Filter out any settings for columns that no longer exist
      const validExisting = existingSettings.filter(s =>
        nonTitleColumns.some(col => col.internalName === s.internalName)
      );

      return [...validExisting, ...newColumns];
    }

    return nonTitleColumns.map(col => ({
      internalName: col.internalName,
      visible: true,
      displayStyle: 'list' as const,
    }));
  });

  // Initialize related sections (for add/edit/remove)
  const [relatedSections, setRelatedSections] = useState<RelatedSection[]>(() => {
    return [...existingRelatedSections];
  });

  // Initialize section order
  const [sectionOrder, setSectionOrder] = useState<string[]>(() => {
    if (existingDetailLayout?.relatedSectionOrder) {
      // Include any new sections at the end
      const existingOrder = existingDetailLayout.relatedSectionOrder;
      const allIds = existingRelatedSections.map(s => s.id);
      const orderedIds = existingOrder.filter(id => allIds.includes(id));
      const newIds = allIds.filter(id => !existingOrder.includes(id));
      return [...orderedIds, ...newIds];
    }
    return existingRelatedSections.map(s => s.id);
  });

  // Drag state for columns and sections
  const [draggedColIndex, setDraggedColIndex] = useState<number | null>(null);
  const [draggedSectionIndex, setDraggedSectionIndex] = useState<number | null>(null);

  // Flyout state for add/edit related section
  const [flyoutOpen, setFlyoutOpen] = useState(false);
  const [editingSection, setEditingSection] = useState<RelatedSection | null>(null);

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

  // Related section management
  const handleAddSection = useCallback(() => {
    setEditingSection(null);
    setFlyoutOpen(true);
  }, []);

  const handleEditSection = useCallback((section: RelatedSection) => {
    setEditingSection(section);
    setFlyoutOpen(true);
  }, []);

  const handleRemoveSection = useCallback((sectionId: string) => {
    setRelatedSections(prev => prev.filter(s => s.id !== sectionId));
    setSectionOrder(prev => prev.filter(id => id !== sectionId));
  }, []);

  const handleSaveSection = useCallback((section: RelatedSection) => {
    setRelatedSections(prev => {
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

  const handleSave = async () => {
    // Check if related sections changed
    const sectionsChanged = JSON.stringify(relatedSections) !== JSON.stringify(existingRelatedSections);
    await onSave(
      {
        columnSettings,
        relatedSectionOrder: sectionOrder,
      },
      sectionsChanged ? relatedSections : undefined
    );
  };

  // Get display name for a column
  const getColumnDisplayName = (internalName: string): string => {
    const col = displayColumns.find(c => c.internalName === internalName);
    return col?.displayName || internalName;
  };

  // Get section by ID (from local state)
  const getSectionById = (id: string): RelatedSection | undefined => {
    return relatedSections.find(s => s.id === id);
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
        {/* Column Settings Section */}
        <div className={styles.section}>
          <Text className={styles.sectionTitle}>Columns</Text>
          <Text className={styles.sectionHint}>
            Drag to reorder, toggle visibility, and choose display style
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

        {/* Related Sections */}
        <Divider />
        <div className={styles.section} style={{ marginTop: '24px' }}>
          <div className={styles.sectionHeader}>
            <Text className={styles.sectionTitle}>Related Sections</Text>
            <Button
              appearance="subtle"
              size="small"
              icon={<AddRegular />}
              onClick={handleAddSection}
            >
              Add
            </Button>
          </div>
          <Text className={styles.sectionHint}>
            Drag to reorder, click edit to configure settings
          </Text>

          {sectionOrder.length === 0 ? (
            <div className={styles.emptySection}>
              <Text>No related sections configured.</Text>
            </div>
          ) : (
            sectionOrder.map((sectionId, index) => {
              const section = getSectionById(sectionId);
              if (!section) return null;

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
                  <Text className={styles.itemName}>{section.title}</Text>
                  <div className={styles.sectionActions}>
                    <Button
                      appearance="subtle"
                      size="small"
                      icon={<EditRegular />}
                      onClick={(e) => {
                        e.stopPropagation();
                        handleEditSection(section);
                      }}
                      title="Edit"
                    />
                    <Button
                      appearance="subtle"
                      size="small"
                      icon={<DeleteRegular />}
                      onClick={(e) => {
                        e.stopPropagation();
                        handleRemoveSection(sectionId);
                      }}
                      title="Remove"
                    />
                  </div>
                </div>
              );
            })
          )}
        </div>

        {/* Footer */}
        <div className={styles.footer}>
          <Button appearance="secondary" onClick={onClose}>
            Cancel
          </Button>
          <Button appearance="primary" onClick={handleSave}>
            Save
          </Button>
        </div>
      </DrawerBody>
    </OverlayDrawer>

    {/* Related Section Flyout */}
    <RelatedSectionFlyout
      open={flyoutOpen}
      section={editingSection}
      primaryListId={primaryListId}
      primarySiteId={primarySiteId}
      primarySiteUrl={primarySiteUrl}
      onClose={() => setFlyoutOpen(false)}
      onSave={handleSaveSection}
    />
  </>
  );
}

export default DetailCustomizeDrawer;
