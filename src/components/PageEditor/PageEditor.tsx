import { useState, useEffect, useCallback } from 'react';
import { useMsal } from '@azure/msal-react';
import {
  makeStyles,
  tokens,
  Button,
  Card,
  Input,
  Textarea,
  Checkbox,
  Text,
  Spinner,
  MessageBar,
  MessageBarBody,
  Badge,
  Field,
  Link,
  mergeClasses,
} from '@fluentui/react-components';
import {
  DismissCircleRegular,
  ReOrderDotsVerticalRegular,
  DismissRegular,
  CheckmarkRegular,
  ArrowRightRegular,
  ArrowLeftRegular,
  DocumentRegular,
  DatabaseRegular,
  TableRegular,
  FilterRegular,
  CheckmarkCircleFilled,
} from '@fluentui/react-icons';
import { getListColumns, type GraphListColumn } from '../../auth/graphClient';
import DataSourcePicker from '../PageDisplay/WebParts/DataSourcePicker';
import type {
  PageDefinition,
  PageSource,
  PageColumn,
  SearchConfig,
  FilterColumn,
  PageType,
  WebPartDataSource,
} from '../../types/page';
import { useTheme } from '../../contexts/ThemeContext';

interface PageEditorProps {
  initialPage?: PageDefinition;
  onSave: (page: PageDefinition) => Promise<PageDefinition>;
  onCancel: () => void;
}

type Step = 'basic' | 'source' | 'columns' | 'search';

// Steps for Search pages
const LOOKUP_STEPS: { key: Step; label: string; icon: React.ReactNode }[] = [
  { key: 'basic', label: 'Basic Information', icon: <DocumentRegular /> },
  { key: 'source', label: 'Data Source', icon: <DatabaseRegular /> },
  { key: 'columns', label: 'Display Columns', icon: <TableRegular /> },
  { key: 'search', label: 'Search & Filters', icon: <FilterRegular /> },
];

// Simplified steps for Report pages (just basic info)
const REPORT_STEPS: { key: Step; label: string; icon: React.ReactNode }[] = [
  { key: 'basic', label: 'Basic Information', icon: <DocumentRegular /> },
];

interface ColumnWithMeta extends GraphListColumn {
  sourceListId: string;
  sourceListName: string;
}

const useStyles = makeStyles({
  container: {
    display: 'flex',
    gap: '32px',
    minHeight: '480px',
  },
  // Vertical stepper styles
  stepper: {
    display: 'flex',
    flexDirection: 'column',
    width: '220px',
    flexShrink: 0,
  },
  stepItem: {
    display: 'flex',
    alignItems: 'flex-start',
    gap: '12px',
    position: 'relative',
    cursor: 'pointer',
    padding: '12px 16px',
    borderRadius: tokens.borderRadiusMedium,
    transition: 'background-color 0.15s ease',
    '&:hover': {
      backgroundColor: tokens.colorNeutralBackground3,
    },
  },
  stepItemActive: {
    backgroundColor: tokens.colorBrandBackground2,
    '&:hover': {
      backgroundColor: tokens.colorBrandBackground2,
    },
  },
  stepItemDisabled: {
    cursor: 'not-allowed',
    opacity: 0.5,
    '&:hover': {
      backgroundColor: 'transparent',
    },
  },
  stepIconContainer: {
    width: '32px',
    height: '32px',
    borderRadius: '50%',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    backgroundColor: '#e0e0e0',
    color: tokens.colorNeutralForeground3,
    fontSize: '16px',
    flexShrink: 0,
    transition: 'all 0.15s ease',
    position: 'relative',
    zIndex: 1,
  },
  stepIconContainerDark: {
    backgroundColor: '#2a2a2a',
  },
  stepIconActive: {
    backgroundColor: tokens.colorBrandBackground,
    color: tokens.colorNeutralForegroundOnBrand,
  },
  stepIconCompleted: {
    backgroundColor: tokens.colorPaletteGreenBackground3,
    color: '#fff',
  },
  stepContent: {
    display: 'flex',
    flexDirection: 'column',
    gap: '2px',
    paddingTop: '4px',
  },
  stepLabel: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground2,
    lineHeight: '1.4',
  },
  stepLabelActive: {
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground1,
  },
  stepLine: {
    position: 'absolute',
    left: '31px',
    top: '56px',
    width: '2px',
    height: 'calc(100% - 56px)',
    backgroundColor: tokens.colorNeutralStroke1,
    zIndex: 0,
  },
  stepLineCompleted: {
    backgroundColor: tokens.colorPaletteGreenBackground3,
  },
  // Content area styles
  contentArea: {
    flex: 1,
    display: 'flex',
    flexDirection: 'column',
  },
  contentCard: {
    flex: 1,
    padding: '24px',
    marginBottom: '24px',
  },
  contentCardDark: {
    backgroundColor: '#1a1a1a',
    border: '1px solid #333',
  },
  formSection: {
    display: 'flex',
    flexDirection: 'column',
    gap: '20px',
  },
  sectionTitle: {
    fontSize: tokens.fontSizeBase500,
    fontWeight: tokens.fontWeightSemibold,
    marginBottom: '4px',
    display: 'block',
  },
  helperText: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground2,
    marginBottom: '16px',
  },
  // Page type selector
  pageTypeGrid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(2, 1fr)',
    gap: '16px',
  },
  pageTypeOption: {
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    gap: '12px',
    padding: '20px',
    borderRadius: tokens.borderRadiusMedium,
    border: `2px solid ${tokens.colorNeutralStroke1}`,
    cursor: 'pointer',
    transition: 'all 0.15s ease',
    ':hover': {
      backgroundColor: tokens.colorNeutralBackground3,
    },
  },
  pageTypeOptionSelected: {
    border: `2px solid ${tokens.colorBrandStroke1}`,
    backgroundColor: tokens.colorBrandBackground2,
  },
  pageTypeLabel: {
    textAlign: 'center',
  },
  // Columns grid
  columnsGrid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(2, 1fr)',
    gap: '20px',
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
    marginBottom: '8px',
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
    maxHeight: '280px',
    overflowY: 'auto',
  },
  columnItem: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
    padding: '10px 12px',
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
  // Search & filters
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
  // Navigation
  navigation: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
  },
  navRight: {
    display: 'flex',
    alignItems: 'center',
    gap: '12px',
  },
  loadingContainer: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    padding: '48px',
  },
});

function PageEditor({ initialPage, onSave, onCancel }: PageEditorProps) {
  const styles = useStyles();
  const { theme } = useTheme();
  const { instance, accounts } = useMsal();
  const account = accounts[0];

  // Current step
  const [currentStep, setCurrentStep] = useState<Step>('basic');

  // Page state
  const [name, setName] = useState(initialPage?.name || '');
  const [description, setDescription] = useState(initialPage?.description || '');
  const [pageType, setPageType] = useState<PageType>(initialPage?.pageType || 'lookup');
  const [primarySource, setPrimarySource] = useState<PageSource | null>(
    initialPage?.primarySource || null
  );
  const [displayColumns, setDisplayColumns] = useState<PageColumn[]>(
    initialPage?.displayColumns || []
  );
  const [searchConfig, setSearchConfig] = useState<SearchConfig>(
    initialPage?.searchConfig || {
      textSearchColumns: [],
      filterColumns: [],
    }
  );

  // Available columns from primary source
  const [availableColumns, setAvailableColumns] = useState<ColumnWithMeta[]>([]);
  const [loadingColumns, setLoadingColumns] = useState(false);
  const [showHiddenColumns, setShowHiddenColumns] = useState(false);

  // Saving state
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState<string | null>(null);

  // Drag and drop state for column reordering
  const [draggedColIndex, setDraggedColIndex] = useState<number | null>(null);

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
  }, [instance, account, primarySource]);

  // Auto-select default text search column and filter columns when display columns change
  useEffect(() => {
    if (displayColumns.length === 0) return;

    setSearchConfig((prev) => {
      let newTextSearchColumns = prev.textSearchColumns;

      // Auto-select Title if present in display columns, otherwise first column
      if (newTextSearchColumns.length === 0) {
        const titleCol = displayColumns.find(
          (col) => col.internalName === 'Title' || col.displayName === 'Title'
        );
        if (titleCol) {
          newTextSearchColumns = [titleCol.internalName];
        } else if (displayColumns.length > 0) {
          newTextSearchColumns = [displayColumns[0].internalName];
        }
      }

      return {
        ...prev,
        textSearchColumns: newTextSearchColumns,
      };
    });
  }, [displayColumns]);

  // Auto-select all filter columns when columns are loaded
  useEffect(() => {
    if (availableColumns.length === 0) return;

    const choiceCols = availableColumns.filter(
      (col) => col.choice || col.lookup || col.name === 'Boolean'
    );

    if (choiceCols.length > 0 && searchConfig.filterColumns.length === 0) {
      const filters: FilterColumn[] = choiceCols.map((col) => {
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
          internalName: col.name,
          displayName: col.displayName,
          type,
        };
      });

      setSearchConfig((prev) => ({
        ...prev,
        filterColumns: filters,
      }));
    }
  }, [availableColumns]);

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

  // Get steps based on page type
  const STEPS = pageType === 'report' ? REPORT_STEPS : LOOKUP_STEPS;
  const currentStepIndex = STEPS.findIndex((s) => s.key === currentStep);

  const isStepComplete = (stepKey: Step): boolean => {
    switch (stepKey) {
      case 'basic':
        return name.trim().length > 0;
      case 'source':
        return primarySource !== null;
      case 'columns':
        return displayColumns.length > 0;
      case 'search':
        return true; // Optional step
      default:
        return false;
    }
  };

  const canNavigateToStep = (stepIndex: number): boolean => {
    // Can always go back
    if (stepIndex < currentStepIndex) return true;
    // Can only go forward if all previous steps are complete
    for (let i = 0; i < stepIndex; i++) {
      if (!isStepComplete(STEPS[i].key)) return false;
    }
    return true;
  };

  const canProceed = (): boolean => {
    return isStepComplete(currentStep);
  };

  const isLastStep = currentStepIndex === STEPS.length - 1;

  const handleNext = () => {
    if (isLastStep) {
      handleSave();
    } else {
      const nextIndex = currentStepIndex + 1;
      if (nextIndex < STEPS.length) {
        setCurrentStep(STEPS[nextIndex].key);
      }
    }
  };

  const handleBack = () => {
    const prevIndex = currentStepIndex - 1;
    if (prevIndex >= 0) {
      setCurrentStep(STEPS[prevIndex].key);
    }
  };

  const handleStepClick = (stepKey: Step, stepIndex: number) => {
    if (canNavigateToStep(stepIndex)) {
      setCurrentStep(stepKey);
    }
  };

  const handleSave = async () => {
    // For lookup pages, primarySource is required
    if (pageType === 'lookup' && !primarySource) return;

    setSaving(true);
    setError(null);

    try {
      const page: PageDefinition = {
        id: initialPage?.id,
        name,
        description,
        pageType,
        primarySource: primarySource || { siteId: '', listId: '', listName: '' },
        displayColumns,
        searchConfig: {
          ...searchConfig,
          // Use displayColumns as table columns
          tableColumns: displayColumns,
        },
        relatedSections: initialPage?.relatedSections || [],
      };

      await onSave(page);
    } catch (err) {
      console.error('Failed to save page:', err);
      setError(err instanceof Error ? err.message : 'Failed to save page');
    } finally {
      setSaving(false);
    }
  };

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
    <div className={styles.container}>
      {/* Vertical Stepper */}
      <div className={styles.stepper}>
        {STEPS.map((step, index) => {
          const isActive = step.key === currentStep;
          const isPastCurrent = index < currentStepIndex;
          const canNavigate = canNavigateToStep(index);
          const isLast = index === STEPS.length - 1;

          return (
            <div
              key={step.key}
              className={mergeClasses(
                styles.stepItem,
                isActive && styles.stepItemActive,
                !canNavigate && styles.stepItemDisabled
              )}
              onClick={() => handleStepClick(step.key, index)}
            >
              {/* Connecting line - rendered first so it's behind the icon */}
              {!isLast && (
                <div
                  className={mergeClasses(
                    styles.stepLine,
                    isPastCurrent && styles.stepLineCompleted
                  )}
                />
              )}

              {/* Step icon */}
              <div
                className={mergeClasses(
                  styles.stepIconContainer,
                  theme === 'dark' && styles.stepIconContainerDark,
                  isActive && styles.stepIconActive,
                  isPastCurrent && styles.stepIconCompleted
                )}
              >
                {isPastCurrent ? <CheckmarkCircleFilled /> : step.icon}
              </div>

              {/* Step label */}
              <div className={styles.stepContent}>
                <Text
                  className={mergeClasses(
                    styles.stepLabel,
                    isActive && styles.stepLabelActive
                  )}
                >
                  {step.label}
                </Text>
              </div>
            </div>
          );
        })}
      </div>

      {/* Content Area */}
      <div className={styles.contentArea}>
        {/* Error Display */}
        {error && (
          <MessageBar intent="error" style={{ marginBottom: '16px' }}>
            <MessageBarBody>
              <DismissCircleRegular /> {error}
            </MessageBarBody>
          </MessageBar>
        )}

        <Card
          className={mergeClasses(
            styles.contentCard,
            theme === 'dark' && styles.contentCardDark
          )}
        >
          {/* Basic Info Step */}
          {currentStep === 'basic' && (
            <div className={styles.formSection}>
              <div>
                <Text className={styles.sectionTitle}>Basic Information</Text>
                <Text className={styles.helperText}>
                  Name your page and choose what type it should be.
                </Text>
              </div>

              <Field label="Page Name" required>
                <Input
                  placeholder="e.g., Student Details"
                  value={name}
                  onChange={(_e, data) => setName(data.value)}
                  size="large"
                />
              </Field>

              <Field label="Description">
                <Textarea
                  placeholder="Optional description of what this page shows"
                  rows={2}
                  value={description}
                  onChange={(_e, data) => setDescription(data.value)}
                />
              </Field>

              <Field label="Page Type">
                <div className={styles.pageTypeGrid}>
                  {/* Search Type */}
                  <div
                    className={mergeClasses(
                      styles.pageTypeOption,
                      pageType === 'lookup' && styles.pageTypeOptionSelected
                    )}
                    onClick={() => setPageType('lookup')}
                  >
                    <svg
                      viewBox="0 0 120 80"
                      style={{ width: '100%', height: '60px' }}
                      fill="none"
                    >
                      <rect
                        x="2"
                        y="2"
                        width="28"
                        height="76"
                        rx="2"
                        fill={tokens.colorNeutralBackground3}
                        stroke={tokens.colorNeutralStroke1}
                      />
                      <rect x="5" y="6" width="22" height="5" rx="1" fill={tokens.colorNeutralForeground3} />
                      <rect x="5" y="14" width="22" height="4" rx="1" fill={tokens.colorNeutralBackground4} />
                      <rect x="5" y="21" width="22" height="4" rx="1" fill={tokens.colorNeutralBackground4} />
                      <rect
                        x="34"
                        y="2"
                        width="84"
                        height="76"
                        rx="2"
                        fill={tokens.colorNeutralBackground3}
                        stroke={tokens.colorNeutralStroke1}
                      />
                      <rect x="38" y="6" width="76" height="8" rx="1" fill={tokens.colorNeutralForeground3} />
                      <line x1="38" y1="18" x2="114" y2="18" stroke={tokens.colorNeutralStroke1} />
                      <rect x="38" y="22" width="20" height="3" rx="1" fill={tokens.colorBrandBackground} />
                      <rect x="62" y="22" width="25" height="3" rx="1" fill={tokens.colorNeutralBackground4} />
                      <rect x="91" y="22" width="20" height="3" rx="1" fill={tokens.colorNeutralBackground4} />
                    </svg>
                    <div className={styles.pageTypeLabel}>
                      <Text weight="semibold" size={300}>
                        Search
                      </Text>
                      <Text
                        size={200}
                        style={{ color: tokens.colorNeutralForeground2, display: 'block' }}
                      >
                        Browse and filter data
                      </Text>
                    </div>
                  </div>

                  {/* Report Type */}
                  <div
                    className={mergeClasses(
                      styles.pageTypeOption,
                      pageType === 'report' && styles.pageTypeOptionSelected
                    )}
                    onClick={() => setPageType('report')}
                  >
                    <svg
                      viewBox="0 0 120 80"
                      style={{ width: '100%', height: '60px' }}
                      fill="none"
                    >
                      <rect
                        x="2"
                        y="2"
                        width="116"
                        height="76"
                        rx="2"
                        fill={tokens.colorNeutralBackground3}
                        stroke={tokens.colorNeutralStroke1}
                      />
                      <rect
                        x="6"
                        y="6"
                        width="52"
                        height="32"
                        rx="2"
                        fill={tokens.colorNeutralBackground4}
                      />
                      <polyline
                        points="10,30 18,24 26,26 34,18 42,20 50,12"
                        stroke={tokens.colorBrandBackground}
                        strokeWidth="2"
                        fill="none"
                        strokeLinecap="round"
                      />
                      <rect
                        x="62"
                        y="6"
                        width="52"
                        height="32"
                        rx="2"
                        fill={tokens.colorNeutralBackground4}
                      />
                      <polyline
                        points="66,22 74,28 82,20 90,24 98,16 106,18"
                        stroke={tokens.colorPaletteGreenBackground3}
                        strokeWidth="2"
                        fill="none"
                        strokeLinecap="round"
                      />
                      <rect x="6" y="42" width="108" height="6" rx="1" fill={tokens.colorNeutralForeground3} />
                    </svg>
                    <div className={styles.pageTypeLabel}>
                      <Text weight="semibold" size={300}>
                        Report
                      </Text>
                      <Text
                        size={200}
                        style={{ color: tokens.colorNeutralForeground2, display: 'block' }}
                      >
                        Charts and dashboards
                      </Text>
                    </div>
                  </div>
                </div>
              </Field>
            </div>
          )}

          {/* Data Source Step */}
          {currentStep === 'source' && (
            <div className={styles.formSection}>
              <div>
                <Text className={styles.sectionTitle}>Data Source</Text>
                <Text className={styles.helperText}>
                  Choose the SharePoint list this page will display data from.
                </Text>
              </div>

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
          )}

          {/* Display Columns Step */}
          {currentStep === 'columns' && (
            <div className={styles.formSection}>
              <div>
                <Text className={styles.sectionTitle}>Display Columns</Text>
                <Text className={styles.helperText}>
                  Select which columns to show in the table. Drag to reorder.
                </Text>
              </div>

              {loadingColumns ? (
                <div className={styles.loadingContainer}>
                  <Spinner size="large" />
                </div>
              ) : availableColumns.length === 0 ? (
                <MessageBar intent="info">
                  <MessageBarBody>Select a data source first to see available columns.</MessageBarBody>
                </MessageBar>
              ) : (
                <div className={styles.columnsGrid}>
                  {/* Available Columns */}
                  <div className={styles.columnPanel}>
                    <div className={styles.columnPanelHeader}>
                      <Text className={styles.columnPanelTitle}>Available Columns</Text>
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
                          style={{ color: tokens.colorNeutralForeground3, padding: '12px' }}
                        >
                          Click columns on the left to add them
                        </Text>
                      )}
                    </div>
                  </div>
                </div>
              )}
            </div>
          )}

          {/* Search & Filters Step */}
          {currentStep === 'search' && (
            <div className={styles.formSection}>
              <div>
                <Text className={styles.sectionTitle}>Search & Filters</Text>
                <Text className={styles.helperText}>
                  Configure how users can search and filter the data.
                </Text>
              </div>

              {/* Text Search Columns */}
              <Field label="Text Search Columns">
                <Text
                  size={200}
                  style={{
                    color: tokens.colorNeutralForeground2,
                    marginBottom: '12px',
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
              <Field label="Dropdown Filters">
                <Text
                  size={200}
                  style={{
                    color: tokens.colorNeutralForeground2,
                    marginBottom: '12px',
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
          )}
        </Card>

        {/* Navigation Buttons */}
        <div className={styles.navigation}>
          <div>
            {currentStepIndex > 0 ? (
              <Button appearance="subtle" icon={<ArrowLeftRegular />} onClick={handleBack}>
                Back
              </Button>
            ) : (
              <Button appearance="subtle" onClick={onCancel}>
                Cancel
              </Button>
            )}
          </div>

          <div className={styles.navRight}>
            {isLastStep ? (
              <Button
                appearance="primary"
                icon={saving ? <Spinner size="tiny" /> : <CheckmarkRegular />}
                onClick={handleNext}
                disabled={saving || !canProceed()}
              >
                {saving ? 'Creating...' : 'Create Page'}
              </Button>
            ) : (
              <Button
                appearance="primary"
                icon={<ArrowRightRegular />}
                iconPosition="after"
                onClick={handleNext}
                disabled={!canProceed()}
              >
                Next
              </Button>
            )}
          </div>
        </div>
      </div>
    </div>
  );
}

export default PageEditor;
