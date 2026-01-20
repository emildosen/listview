import { useState, useMemo } from 'react';
import { Link } from 'react-router-dom';
import {
  makeStyles,
  mergeClasses,
  tokens,
  Button,
  Card,
  Text,
  Title2,
  Breadcrumb,
  BreadcrumbItem,
  BreadcrumbDivider,
  Table,
  TableHeader,
  TableRow,
  TableHeaderCell,
  TableBody,
  TableCell,
  TableCellLayout,
  Input,
  Dialog,
  DialogTrigger,
  DialogSurface,
  DialogBody,
  DialogTitle,
  DialogContent,
  DialogActions,
  Field,
} from '@fluentui/react-components';
import {
  AddRegular,
  DeleteRegular,
  EditRegular,
  FolderRegular,
  ReOrderDotsVerticalRegular,
} from '@fluentui/react-icons';
import {
  DndContext,
  closestCenter,
  KeyboardSensor,
  PointerSensor,
  useSensor,
  useSensors,
  type DragEndEvent,
} from '@dnd-kit/core';
import {
  arrayMove,
  SortableContext,
  sortableKeyboardCoordinates,
  useSortable,
  verticalListSortingStrategy,
} from '@dnd-kit/sortable';
import { CSS } from '@dnd-kit/utilities';
import { useSettings } from '../contexts/SettingsContext';
import { useTheme } from '../contexts/ThemeContext';
import type { Section } from '../types/page';

const useStyles = makeStyles({
  container: {
    padding: '32px',
    flex: 1,
  },
  breadcrumb: {
    marginBottom: '24px',
  },
  breadcrumbLink: {
    textDecoration: 'none',
    color: 'inherit',
  },
  content: {
    maxWidth: '800px',
  },
  header: {
    display: 'flex',
    alignItems: 'flex-start',
    justifyContent: 'space-between',
    marginBottom: '24px',
  },
  description: {
    display: 'block',
    color: tokens.colorNeutralForeground2,
    marginTop: '4px',
  },
  tableCard: {
    boxShadow: '0 1px 2px rgba(0, 0, 0, 0.08), 0 2px 4px rgba(0, 0, 0, 0.04)',
    borderRadius: '2px',
    overflow: 'hidden',
    backgroundColor: tokens.colorNeutralBackground1,
    border: '1px solid transparent',
    borderImage: 'linear-gradient(135deg, rgba(0,0,0,0.05) 0%, rgba(0,0,0,0.15) 100%) 1',
  },
  tableCardDark: {
    backgroundColor: '#1a1a1a',
    borderImage: 'none',
    border: '1px solid #333333',
  },
  toolbar: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    padding: '12px 16px',
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  tableHeaderCell: {
    fontWeight: tokens.fontWeightSemibold,
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground2,
    textTransform: 'uppercase',
    letterSpacing: '0.02em',
  },
  tableRow: {
    transitionProperty: 'background-color',
    transitionDuration: tokens.durationNormal,
  },
  tableRowHover: {
    backgroundColor: tokens.colorNeutralBackground1Hover,
  },
  dragHandle: {
    cursor: 'grab',
    color: tokens.colorNeutralForeground3,
    display: 'flex',
    alignItems: 'center',
  },
  dragHandleActive: {
    cursor: 'grabbing',
  },
  sectionIcon: {
    color: tokens.colorNeutralForeground3,
  },
  actions: {
    display: 'flex',
    gap: '4px',
    opacity: 0,
    transitionProperty: 'opacity',
    transitionDuration: tokens.durationNormal,
  },
  actionsVisible: {
    opacity: 1,
  },
  emptyState: {
    padding: '48px',
    textAlign: 'center',
    color: tokens.colorNeutralForeground2,
  },
  countBadge: {
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase200,
  },
  dialogInput: {
    width: '100%',
    marginTop: '8px',
  },
});

// Sortable table row component
interface SortableRowProps {
  section: Section;
  pageCount: number;
  onEdit: () => void;
  onDelete: () => void;
}

function SortableRow({ section, pageCount, onEdit, onDelete }: SortableRowProps) {
  const styles = useStyles();
  const [isHovered, setIsHovered] = useState(false);

  const {
    attributes,
    listeners,
    setNodeRef,
    transform,
    transition,
    isDragging,
  } = useSortable({ id: section.id });

  const style = {
    transform: CSS.Transform.toString(transform),
    transition,
    opacity: isDragging ? 0.5 : 1,
  };

  return (
    <TableRow
      ref={setNodeRef}
      style={style}
      className={mergeClasses(
        styles.tableRow,
        isHovered && styles.tableRowHover
      )}
      onMouseEnter={() => setIsHovered(true)}
      onMouseLeave={() => setIsHovered(false)}
    >
      <TableCell style={{ width: '40px' }}>
        <div
          {...attributes}
          {...listeners}
          className={mergeClasses(
            styles.dragHandle,
            isDragging && styles.dragHandleActive
          )}
        >
          <ReOrderDotsVerticalRegular fontSize={20} />
        </div>
      </TableCell>
      <TableCell>
        <TableCellLayout media={<FolderRegular className={styles.sectionIcon} />}>
          {section.name}
        </TableCellLayout>
      </TableCell>
      <TableCell>
        <span className={styles.countBadge}>
          {pageCount} {pageCount === 1 ? 'page' : 'pages'}
        </span>
      </TableCell>
      <TableCell style={{ width: '100px' }}>
        <div
          className={mergeClasses(
            styles.actions,
            isHovered && styles.actionsVisible
          )}
        >
          <Button
            appearance="subtle"
            size="small"
            icon={<EditRegular />}
            onClick={onEdit}
          />
          <Button
            appearance="subtle"
            size="small"
            icon={<DeleteRegular />}
            onClick={onDelete}
          />
        </div>
      </TableCell>
    </TableRow>
  );
}

function SectionsPage() {
  const styles = useStyles();
  const { theme } = useTheme();
  const { sections, pages, saveSection, removeSection, reorderSections } = useSettings();

  const [newSectionOpen, setNewSectionOpen] = useState(false);
  const [newSectionName, setNewSectionName] = useState('');
  const [editSection, setEditSection] = useState<Section | null>(null);
  const [editName, setEditName] = useState('');
  const [deleteSection, setDeleteSection] = useState<Section | null>(null);

  // Sorted sections
  const sortedSections = useMemo(() => {
    return Object.values(sections).sort((a, b) => a.order - b.order);
  }, [sections]);

  // Page counts by section
  const pageCounts = useMemo(() => {
    const counts: Record<string, number> = {};
    pages.forEach((page) => {
      if (page.sectionId) {
        counts[page.sectionId] = (counts[page.sectionId] || 0) + 1;
      }
    });
    return counts;
  }, [pages]);

  // dnd-kit sensors
  const sensors = useSensors(
    useSensor(PointerSensor),
    useSensor(KeyboardSensor, {
      coordinateGetter: sortableKeyboardCoordinates,
    })
  );

  // Handle drag end
  const handleDragEnd = async (event: DragEndEvent) => {
    const { active, over } = event;
    if (over && active.id !== over.id) {
      const oldIndex = sortedSections.findIndex((s) => s.id === active.id);
      const newIndex = sortedSections.findIndex((s) => s.id === over.id);
      const newOrder = arrayMove(sortedSections, oldIndex, newIndex);
      await reorderSections(newOrder.map((s) => s.id));
    }
  };

  // Create new section
  const handleCreateSection = async () => {
    if (!newSectionName.trim()) return;

    const newSection: Section = {
      id: `section-${Date.now()}`,
      name: newSectionName.trim(),
      order: sortedSections.length,
    };

    await saveSection(newSection);
    setNewSectionName('');
    setNewSectionOpen(false);
  };

  // Edit section
  const handleEditSection = async () => {
    if (!editSection || !editName.trim()) return;

    await saveSection({ ...editSection, name: editName.trim() });
    setEditSection(null);
    setEditName('');
  };

  // Delete section
  const handleDeleteSection = async () => {
    if (!deleteSection) return;

    await removeSection(deleteSection.id);
    setDeleteSection(null);
  };

  return (
    <div className={styles.container}>
      <Breadcrumb className={styles.breadcrumb}>
        <BreadcrumbItem>
          <Link to="/app" className={styles.breadcrumbLink}>
            Home
          </Link>
        </BreadcrumbItem>
        <BreadcrumbDivider />
        <BreadcrumbItem>Sections</BreadcrumbItem>
      </Breadcrumb>

      <div className={styles.content}>
        <div className={styles.header}>
          <div>
            <Title2>Sections</Title2>
            <Text className={styles.description}>
              Organize pages into sections. Drag to reorder.
            </Text>
          </div>
          <Dialog open={newSectionOpen} onOpenChange={(_, data) => setNewSectionOpen(data.open)}>
            <DialogTrigger disableButtonEnhancement>
              <Button appearance="primary" icon={<AddRegular />}>
                New Section
              </Button>
            </DialogTrigger>
            <DialogSurface>
              <DialogBody>
                <DialogTitle>Create Section</DialogTitle>
                <DialogContent>
                  <Field label="Section name">
                    <Input
                      className={styles.dialogInput}
                      value={newSectionName}
                      onChange={(_, data) => setNewSectionName(data.value)}
                      placeholder="e.g., CRM, Reports, Admin"
                      onKeyDown={(e) => {
                        if (e.key === 'Enter') handleCreateSection();
                      }}
                    />
                  </Field>
                </DialogContent>
                <DialogActions>
                  <DialogTrigger disableButtonEnhancement>
                    <Button appearance="secondary">Cancel</Button>
                  </DialogTrigger>
                  <Button
                    appearance="primary"
                    onClick={handleCreateSection}
                    disabled={!newSectionName.trim()}
                  >
                    Create
                  </Button>
                </DialogActions>
              </DialogBody>
            </DialogSurface>
          </Dialog>
        </div>

        <Card
          className={mergeClasses(
            styles.tableCard,
            theme === 'dark' && styles.tableCardDark
          )}
        >
          {sortedSections.length === 0 ? (
            <div className={styles.emptyState}>
              <FolderRegular fontSize={48} style={{ marginBottom: '16px', opacity: 0.5 }} />
              <Text size={400} weight="semibold" block>
                No sections yet
              </Text>
              <Text size={300} style={{ marginTop: '8px' }}>
                Create sections to organize your pages.
              </Text>
            </div>
          ) : (
            <DndContext
              sensors={sensors}
              collisionDetection={closestCenter}
              onDragEnd={handleDragEnd}
            >
              <Table>
                <TableHeader>
                  <TableRow>
                    <TableHeaderCell style={{ width: '40px' }} />
                    <TableHeaderCell className={styles.tableHeaderCell}>
                      Name
                    </TableHeaderCell>
                    <TableHeaderCell className={styles.tableHeaderCell}>
                      Pages
                    </TableHeaderCell>
                    <TableHeaderCell style={{ width: '100px' }} />
                  </TableRow>
                </TableHeader>
                <TableBody>
                  <SortableContext
                    items={sortedSections.map((s) => s.id)}
                    strategy={verticalListSortingStrategy}
                  >
                    {sortedSections.map((section) => (
                      <SortableRow
                        key={section.id}
                        section={section}
                        pageCount={pageCounts[section.id] || 0}
                        onEdit={() => {
                          setEditSection(section);
                          setEditName(section.name);
                        }}
                        onDelete={() => setDeleteSection(section)}
                      />
                    ))}
                  </SortableContext>
                </TableBody>
              </Table>
            </DndContext>
          )}
        </Card>
      </div>

      {/* Edit Dialog */}
      <Dialog
        open={editSection !== null}
        onOpenChange={(_, data) => {
          if (!data.open) {
            setEditSection(null);
            setEditName('');
          }
        }}
      >
        <DialogSurface>
          <DialogBody>
            <DialogTitle>Rename Section</DialogTitle>
            <DialogContent>
              <Field label="Section name">
                <Input
                  className={styles.dialogInput}
                  value={editName}
                  onChange={(_, data) => setEditName(data.value)}
                  onKeyDown={(e) => {
                    if (e.key === 'Enter') handleEditSection();
                  }}
                />
              </Field>
            </DialogContent>
            <DialogActions>
              <Button appearance="secondary" onClick={() => setEditSection(null)}>
                Cancel
              </Button>
              <Button
                appearance="primary"
                onClick={handleEditSection}
                disabled={!editName.trim()}
              >
                Save
              </Button>
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>

      {/* Delete Confirmation Dialog */}
      <Dialog
        open={deleteSection !== null}
        onOpenChange={(_, data) => {
          if (!data.open) setDeleteSection(null);
        }}
      >
        <DialogSurface>
          <DialogBody>
            <DialogTitle>Delete Section</DialogTitle>
            <DialogContent>
              <Text>
                Are you sure you want to delete "{deleteSection?.name}"?
                {pageCounts[deleteSection?.id || ''] > 0 && (
                  <>
                    {' '}
                    The {pageCounts[deleteSection?.id || '']} page(s) in this section will become unsectioned.
                  </>
                )}
              </Text>
            </DialogContent>
            <DialogActions>
              <Button appearance="secondary" onClick={() => setDeleteSection(null)}>
                Cancel
              </Button>
              <Button appearance="primary" onClick={handleDeleteSection}>
                Delete
              </Button>
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>
    </div>
  );
}

export default SectionsPage;
