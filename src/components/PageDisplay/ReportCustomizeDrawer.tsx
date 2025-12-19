import { useState, useCallback, useEffect } from 'react';
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
} from '@fluentui/react-components';
import {
  DismissRegular,
  ReOrderDotsVerticalRegular,
  AddRegular,
  DeleteRegular,
} from '@fluentui/react-icons';
import type {
  PageDefinition,
  ReportSection,
  ReportColumn,
  SectionLayout,
  SectionHeight,
  WebPartType,
  AnyWebPartConfig,
} from '../../types/page';
import LayoutPicker from './LayoutPicker';
import WebPartPicker from './WebPartPicker';

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
  sectionItem: {
    display: 'flex',
    flexDirection: 'column',
    gap: '12px',
    padding: '12px',
    backgroundColor: tokens.colorNeutralBackground2,
    borderRadius: tokens.borderRadiusMedium,
    border: `1px solid transparent`,
    transition: 'opacity 0.2s, border-color 0.2s',
  },
  sectionItemDragging: {
    opacity: 0.5,
    border: `1px dashed ${tokens.colorBrandStroke1}`,
  },
  sectionItemHeader: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
  },
  dragHandle: {
    color: tokens.colorNeutralForeground3,
    cursor: 'grab',
    flexShrink: 0,
  },
  sectionNumber: {
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground2,
    fontSize: tokens.fontSizeBase200,
    flex: 1,
  },
  sectionActions: {
    display: 'flex',
    alignItems: 'center',
    gap: '4px',
    flexShrink: 0,
  },
  layoutLabel: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground2,
    marginBottom: '4px',
  },
  columnsContainer: {
    display: 'flex',
    flexDirection: 'column',
    gap: '8px',
    marginTop: '8px',
  },
  columnItem: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
  },
  columnLabel: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground2,
    minWidth: '70px',
  },
  columnPicker: {
    flex: 1,
  },
  webPartTitle: {
    flex: 1,
  },
  heightContainer: {
    display: 'flex',
    gap: '8px',
    flexWrap: 'wrap',
  },
  heightButton: {
    minWidth: '70px',
    padding: '4px 8px',
    fontSize: tokens.fontSizeBase200,
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    borderRadius: tokens.borderRadiusMedium,
    backgroundColor: tokens.colorNeutralBackground1,
    color: tokens.colorNeutralForeground1,
    cursor: 'pointer',
    transition: 'all 0.15s',
  },
  heightButtonSelected: {
    backgroundColor: tokens.colorBrandBackground,
    color: tokens.colorNeutralForegroundOnBrand,
    border: `1px solid ${tokens.colorBrandBackground}`,
  },
  emptySection: {
    textAlign: 'center',
    padding: '24px',
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

/**
 * Generate a unique ID
 */
function generateId(): string {
  return `${Date.now()}-${Math.random().toString(36).substr(2, 9)}`;
}

/**
 * Create columns for a given layout
 */
function createColumnsForLayout(layout: SectionLayout): ReportColumn[] {
  const columnCounts: Record<SectionLayout, number> = {
    'one-column': 1,
    'two-column': 2,
    'three-column': 3,
    'one-third-left': 2,
    'one-third-right': 2,
  };

  const count = columnCounts[layout];
  return Array.from({ length: count }, () => ({
    id: generateId(),
    webPart: null,
  }));
}

/**
 * Create a new section with default layout
 */
function createSection(layout: SectionLayout = 'one-column'): ReportSection {
  return {
    id: generateId(),
    layout,
    columns: createColumnsForLayout(layout),
  };
}

interface ReportCustomizeDrawerProps {
  page: PageDefinition;
  open: boolean;
  onClose: () => void;
  onSave: (page: PageDefinition) => Promise<void>;
}

export default function ReportCustomizeDrawer({
  page,
  open,
  onClose,
  onSave,
}: ReportCustomizeDrawerProps) {
  const styles = useStyles();

  // Basic info state
  const [name, setName] = useState(page.name);
  const [description, setDescription] = useState(page.description || '');

  // Initialize sections from existing config or create default
  const [sections, setSections] = useState<ReportSection[]>(() => {
    if (page.reportLayout?.sections && page.reportLayout.sections.length > 0) {
      return [...page.reportLayout.sections];
    }
    // Default: one full-width section
    return [createSection('one-column')];
  });

  // Sync state when drawer opens or page changes
  useEffect(() => {
    if (open) {
      setName(page.name);
      setDescription(page.description || '');
      if (page.reportLayout?.sections && page.reportLayout.sections.length > 0) {
        setSections([...page.reportLayout.sections]);
      } else {
        setSections([createSection('one-column')]);
      }
    }
  }, [open, page]);

  // Drag state
  const [draggedIndex, setDraggedIndex] = useState<number | null>(null);
  const [saving, setSaving] = useState(false);

  // Add a new section
  const handleAddSection = useCallback(() => {
    setSections(prev => [...prev, createSection('one-column')]);
  }, []);

  // Delete a section
  const handleDeleteSection = useCallback((index: number) => {
    setSections(prev => prev.filter((_, i) => i !== index));
  }, []);

  // Update section layout
  const handleLayoutChange = useCallback((index: number, layout: SectionLayout) => {
    setSections(prev => {
      const updated = [...prev];
      const section = updated[index];
      const newColumns = createColumnsForLayout(layout);

      // Preserve existing webparts where possible
      section.columns.forEach((col, colIdx) => {
        if (colIdx < newColumns.length && col.webPart) {
          newColumns[colIdx].webPart = col.webPart;
        }
      });

      updated[index] = {
        ...section,
        layout,
        columns: newColumns,
      };
      return updated;
    });
  }, []);

  // Update section height
  const handleHeightChange = useCallback((index: number, height: SectionHeight) => {
    setSections(prev => {
      const updated = [...prev];
      updated[index] = {
        ...updated[index],
        height,
      };
      return updated;
    });
  }, []);

  // Update column webpart type
  const handleWebPartChange = useCallback(
    (sectionIndex: number, columnIndex: number, type: WebPartType | null) => {
      setSections(prev => {
        const updated = [...prev];
        const section = updated[sectionIndex];

        let webPart: AnyWebPartConfig | null = null;
        if (type) {
          webPart = {
            id: generateId(),
            type,
            title: type === 'list-items' ? 'List Items' : 'Chart',
          } as AnyWebPartConfig;
        }

        updated[sectionIndex] = {
          ...section,
          columns: section.columns.map((col, idx) =>
            idx === columnIndex ? { ...col, webPart } : col
          ),
        };
        return updated;
      });
    },
    []
  );

  // Update webpart title
  const handleWebPartTitleChange = useCallback(
    (sectionIndex: number, columnIndex: number, title: string) => {
      setSections(prev => {
        const updated = [...prev];
        const section = updated[sectionIndex];
        const column = section.columns[columnIndex];

        if (column.webPart) {
          updated[sectionIndex] = {
            ...section,
            columns: section.columns.map((col, idx) =>
              idx === columnIndex && col.webPart
                ? { ...col, webPart: { ...col.webPart, title } }
                : col
            ),
          };
        }
        return updated;
      });
    },
    []
  );

  // Drag handlers for reordering sections
  const handleDragStart = useCallback((index: number) => {
    setDraggedIndex(index);
  }, []);

  const handleDragOver = useCallback((e: React.DragEvent, index: number) => {
    e.preventDefault();
    if (draggedIndex === null || draggedIndex === index) return;

    setSections(prev => {
      const updated = [...prev];
      const [dragged] = updated.splice(draggedIndex, 1);
      updated.splice(index, 0, dragged);
      setDraggedIndex(index);
      return updated;
    });
  }, [draggedIndex]);

  const handleDragEnd = useCallback(() => {
    setDraggedIndex(null);
  }, []);

  // Save handler
  const handleSave = useCallback(async () => {
    setSaving(true);
    try {
      const updatedPage: PageDefinition = {
        ...page,
        name,
        description: description || undefined,
        reportLayout: { sections },
      };
      await onSave(updatedPage);
      onClose();
    } finally {
      setSaving(false);
    }
  }, [page, name, description, sections, onSave, onClose]);

  // Get column label based on position and layout
  const getColumnLabel = (layout: SectionLayout, index: number): string => {
    if (layout === 'one-column') return 'Content';
    if (layout === 'one-third-left') return index === 0 ? 'Left (1/3)' : 'Right (2/3)';
    if (layout === 'one-third-right') return index === 0 ? 'Left (2/3)' : 'Right (1/3)';
    return `Column ${index + 1}`;
  };

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
          Customize Page Layout
        </DrawerHeaderTitle>
      </DrawerHeader>

      <DrawerBody className={styles.body}>
        <div className={styles.content}>
          {/* Basic Information */}
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
          </div>

          <Divider />

          {/* Sections */}
          <div className={styles.section} style={{ marginTop: '24px' }}>
            <div className={styles.sectionHeader}>
              <Text className={styles.sectionTitle}>Sections</Text>
              <Button
                appearance="subtle"
                icon={<AddRegular />}
                onClick={handleAddSection}
              >
                Add Section
              </Button>
            </div>
            <Text className={styles.sectionHint}>
              Drag to reorder sections. Each section can have a different column layout.
            </Text>

            {sections.length === 0 ? (
              <div className={styles.emptySection}>
                No sections yet. Click "Add Section" to get started.
              </div>
            ) : (
              sections.map((section, sectionIndex) => (
                <div
                  key={section.id}
                  className={`${styles.sectionItem} ${
                    draggedIndex === sectionIndex ? styles.sectionItemDragging : ''
                  }`}
                  draggable
                  onDragStart={() => handleDragStart(sectionIndex)}
                  onDragOver={(e) => handleDragOver(e, sectionIndex)}
                  onDragEnd={handleDragEnd}
                >
                  <div className={styles.sectionItemHeader}>
                    <ReOrderDotsVerticalRegular className={styles.dragHandle} />
                    <Text className={styles.sectionNumber}>
                      Section {sectionIndex + 1}
                    </Text>
                    <div className={styles.sectionActions}>
                      <Button
                        appearance="subtle"
                        icon={<DeleteRegular />}
                        size="small"
                        onClick={() => handleDeleteSection(sectionIndex)}
                        disabled={sections.length === 1}
                        title={sections.length === 1 ? 'Cannot delete the last section' : 'Delete section'}
                      />
                    </div>
                  </div>

                  <Divider />

                  {/* Layout Picker */}
                  <div>
                    <Text className={styles.layoutLabel}>Layout</Text>
                    <LayoutPicker
                      value={section.layout}
                      onChange={(layout) => handleLayoutChange(sectionIndex, layout)}
                    />
                  </div>

                  {/* Height Picker */}
                  <div>
                    <Text className={styles.layoutLabel}>Height</Text>
                    <div className={styles.heightContainer}>
                      {([
                        { value: 'half', label: 'Half', percent: '50%' },
                        { value: 'medium', label: 'Medium', percent: '75%' },
                        { value: 'full', label: 'Full', percent: '100%' },
                        { value: 'big', label: 'Big', percent: '125%' },
                      ] as const).map((option) => (
                        <button
                          key={option.value}
                          type="button"
                          className={`${styles.heightButton} ${
                            (section.height || 'full') === option.value ? styles.heightButtonSelected : ''
                          }`}
                          onClick={() => handleHeightChange(sectionIndex, option.value)}
                        >
                          {option.percent}
                        </button>
                      ))}
                    </div>
                  </div>

                  {/* Column WebPart Pickers */}
                  <div className={styles.columnsContainer}>
                    <Text className={styles.layoutLabel}>Web Parts</Text>
                    {section.columns.map((column, colIndex) => (
                      <div key={column.id} className={styles.columnItem}>
                        <Text className={styles.columnLabel}>
                          {getColumnLabel(section.layout, colIndex)}:
                        </Text>
                        <div className={styles.columnPicker}>
                          <WebPartPicker
                            value={column.webPart?.type || null}
                            onChange={(type) =>
                              handleWebPartChange(sectionIndex, colIndex, type)
                            }
                          />
                        </div>
                      </div>
                    ))}
                  </div>

                  {/* WebPart Titles */}
                  {section.columns.some((col) => col.webPart) && (
                    <div className={styles.columnsContainer}>
                      <Text className={styles.layoutLabel}>Titles</Text>
                      {section.columns.map((column, colIndex) =>
                        column.webPart ? (
                          <div key={column.id} className={styles.columnItem}>
                            <Text className={styles.columnLabel}>
                              {getColumnLabel(section.layout, colIndex)}:
                            </Text>
                            <Input
                              size="small"
                              className={styles.webPartTitle}
                              value={column.webPart.title || ''}
                              onChange={(_, data) =>
                                handleWebPartTitleChange(
                                  sectionIndex,
                                  colIndex,
                                  data.value
                                )
                              }
                              placeholder="Enter title..."
                            />
                          </div>
                        ) : null
                      )}
                    </div>
                  )}
                </div>
              ))
            )}
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
