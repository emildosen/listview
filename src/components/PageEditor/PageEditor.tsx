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
  Radio,
  RadioGroup,
  Dropdown,
  Option,
  Text,
  Spinner,
  MessageBar,
  MessageBarBody,
  Badge,
  Field,
  Divider,
} from '@fluentui/react-components';
import {
  DismissCircleRegular,
  ReOrderDotsVerticalRegular,
  DismissRegular,
  AddRegular,
  ArrowLeftRegular,
  CheckmarkRegular,
  ArrowRightRegular,
} from '@fluentui/react-icons';
import { useSettings, type EnabledList } from '../../contexts/SettingsContext';
import { getListColumns, type GraphListColumn } from '../../auth/graphClient';
import type {
  PageDefinition,
  PageSource,
  PageColumn,
  SearchConfig,
  FilterColumn,
  RelatedSection,
  PageType,
} from '../../types/page';

interface PageEditorProps {
  initialPage?: PageDefinition;
  onSave: (page: PageDefinition) => Promise<PageDefinition>;
  onCancel: () => void;
}

type Step = 'basic' | 'type' | 'primary' | 'columns' | 'search' | 'related' | 'review';

// All steps for Lookup pages
const LOOKUP_STEPS: { key: Step; label: string }[] = [
  { key: 'basic', label: 'Basic Info' },
  { key: 'type', label: 'Page Type' },
  { key: 'primary', label: 'Primary List' },
  { key: 'columns', label: 'Display Columns' },
  { key: 'search', label: 'Search & Filters' },
  { key: 'related', label: 'Related Lists' },
  { key: 'review', label: 'Review' },
];

// Simplified steps for Report pages (just basic info, type, and review)
const REPORT_STEPS: { key: Step; label: string }[] = [
  { key: 'basic', label: 'Basic Info' },
  { key: 'type', label: 'Page Type' },
  { key: 'review', label: 'Review' },
];

interface ColumnWithMeta extends GraphListColumn {
  sourceListId: string;
  sourceListName: string;
}

const useStyles = makeStyles({
  stepsContainer: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    gap: '8px',
    marginBottom: '32px',
  },
  stepItem: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
    cursor: 'pointer',
  },
  stepLabel: {
    fontSize: tokens.fontSizeBase100,
  },
  stepLabelActive: {
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorBrandForeground1,
  },
  stepConnector: {
    flex: 1,
    height: '2px',
    backgroundColor: tokens.colorNeutralStroke1,
  },
  stepConnectorActive: {
    backgroundColor: tokens.colorBrandBackground,
  },
  cardBody: {
    marginBottom: '24px',
  },
  formSection: {
    display: 'flex',
    flexDirection: 'column',
    gap: '16px',
  },
  sectionTitle: {
    fontSize: tokens.fontSizeBase400,
    fontWeight: tokens.fontWeightSemibold,
    marginBottom: '8px',
  },
  helperText: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground2,
    marginBottom: '16px',
  },
  sourceList: {
    display: 'flex',
    flexDirection: 'column',
    gap: '8px',
  },
  sourceItem: {
    display: 'flex',
    alignItems: 'center',
    gap: '12px',
    padding: '12px',
    borderRadius: tokens.borderRadiusMedium,
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    backgroundColor: tokens.colorNeutralBackground3,
    cursor: 'pointer',
    '&:hover': {
      backgroundColor: tokens.colorNeutralBackground4,
    },
  },
  sourceItemSelected: {
    backgroundColor: tokens.colorBrandBackground2,
    border: `1px solid ${tokens.colorBrandStroke1}`,
  },
  sourceInfo: {
    display: 'flex',
    flexDirection: 'column',
    gap: '2px',
  },
  loadingContainer: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    padding: '32px',
  },
  columnsGrid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(2, 1fr)',
    gap: '16px',
  },
  columnPanel: {
    display: 'flex',
    flexDirection: 'column',
    gap: '4px',
  },
  columnPanelHeader: {
    fontWeight: tokens.fontWeightMedium,
    marginBottom: '8px',
  },
  columnList: {
    display: 'flex',
    flexDirection: 'column',
    gap: '4px',
    maxHeight: '256px',
    overflowY: 'auto',
  },
  columnItem: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
    padding: '8px',
    borderRadius: tokens.borderRadiusMedium,
    backgroundColor: tokens.colorNeutralBackground3,
    cursor: 'pointer',
    '&:hover': {
      backgroundColor: tokens.colorNeutralBackground4,
    },
  },
  columnItemSelected: {
    backgroundColor: tokens.colorBrandBackground2,
    cursor: 'move',
  },
  columnItemDragging: {
    opacity: 0.5,
  },
  dragHandle: {
    color: tokens.colorNeutralForeground3,
  },
  displayModeGrid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(2, 1fr)',
    gap: '16px',
  },
  displayModeOption: {
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    gap: '12px',
    padding: '16px',
    borderRadius: tokens.borderRadiusMedium,
    border: `2px solid ${tokens.colorNeutralStroke1}`,
    cursor: 'pointer',
    '&:hover': {
      border: `2px solid ${tokens.colorNeutralStroke2}`,
    },
  },
  displayModeOptionSelected: {
    border: `2px solid ${tokens.colorBrandStroke1}`,
    backgroundColor: tokens.colorBrandBackground2,
  },
  displayModeLabel: {
    textAlign: 'center',
  },
  badgeWrap: {
    display: 'flex',
    flexWrap: 'wrap',
    gap: '8px',
  },
  badgeItem: {
    cursor: 'pointer',
  },
  checkboxList: {
    display: 'flex',
    flexDirection: 'column',
    gap: '4px',
  },
  relatedHeader: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    marginBottom: '16px',
  },
  relatedEmpty: {
    padding: '32px',
    textAlign: 'center',
    backgroundColor: tokens.colorNeutralBackground3,
    borderRadius: tokens.borderRadiusMedium,
  },
  relatedSectionCard: {
    marginBottom: '16px',
    padding: '16px',
  },
  relatedSectionHeader: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    marginBottom: '12px',
  },
  permissionsRow: {
    display: 'flex',
    flexWrap: 'wrap',
    gap: '16px',
    marginTop: '8px',
  },
  reviewGrid: {
    display: 'flex',
    flexDirection: 'column',
    gap: '16px',
  },
  reviewCard: {
    padding: '16px',
  },
  reviewCardTitle: {
    fontWeight: tokens.fontWeightMedium,
    marginBottom: '8px',
  },
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
});

function PageEditor({ initialPage, onSave, onCancel }: PageEditorProps) {
  const styles = useStyles();
  const { instance, accounts } = useMsal();
  const { enabledLists } = useSettings();
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
      tableColumns: [],
      textSearchColumns: [],
      filterColumns: [],
    }
  );
  const [relatedSections, setRelatedSections] = useState<RelatedSection[]>(
    initialPage?.relatedSections || []
  );

  // Available columns from primary source
  const [availableColumns, setAvailableColumns] = useState<ColumnWithMeta[]>([]);
  const [loadingColumns, setLoadingColumns] = useState(false);

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

  const handlePrimarySourceSelect = useCallback((list: EnabledList) => {
    setPrimarySource({
      siteId: list.siteId,
      siteUrl: list.siteUrl,
      listId: list.listId,
      listName: list.listName,
    });
    // Clear columns when source changes
    setDisplayColumns([]);
    setSearchConfig({
      tableColumns: [],
      textSearchColumns: [],
      filterColumns: [],
    });
  }, []);

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

  const handleFilterColumnToggle = useCallback(
    (col: ColumnWithMeta) => {
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
        } else if (col.name === 'Boolean' || col.displayName?.toLowerCase().includes('yes') || col.displayName?.toLowerCase().includes('no')) {
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
    },
    []
  );

  const handleAddRelatedSection = useCallback(() => {
    const newSection: RelatedSection = {
      id: `section-${Date.now()}`,
      title: 'Related Items',
      source: { siteId: '', siteUrl: '', listId: '', listName: '' },
      lookupColumn: '',
      displayColumns: [],
      allowCreate: true,
      allowEdit: true,
      allowDelete: true,
    };
    setRelatedSections((prev) => [...prev, newSection]);
  }, []);

  const handleUpdateRelatedSection = useCallback((index: number, updates: Partial<RelatedSection>) => {
    setRelatedSections((prev) => {
      const newSections = [...prev];
      newSections[index] = { ...newSections[index], ...updates };
      return newSections;
    });
  }, []);

  const handleRemoveRelatedSection = useCallback((index: number) => {
    setRelatedSections((prev) => prev.filter((_, i) => i !== index));
  }, []);

  // Get steps based on page type
  const STEPS = pageType === 'report' ? REPORT_STEPS : LOOKUP_STEPS;
  const currentStepIndex = STEPS.findIndex((s) => s.key === currentStep);

  const canProceed = (): boolean => {
    switch (currentStep) {
      case 'basic':
        return name.trim().length > 0;
      case 'type':
        return true; // Page type is always valid (has default)
      case 'primary':
        return primarySource !== null;
      case 'columns':
        return displayColumns.length > 0;
      case 'search':
        return (searchConfig.tableColumns?.length || 0) > 0;
      case 'related':
        return true; // Related sections are optional
      case 'review':
        return true;
      default:
        return true;
    }
  };

  const handleNext = () => {
    const nextIndex = currentStepIndex + 1;
    if (nextIndex < STEPS.length) {
      setCurrentStep(STEPS[nextIndex].key);
    }
  };

  const handleBack = () => {
    const prevIndex = currentStepIndex - 1;
    if (prevIndex >= 0) {
      setCurrentStep(STEPS[prevIndex].key);
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
        searchConfig,
        relatedSections,
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

  return (
    <div>
      {/* Step Indicator */}
      <div className={styles.stepsContainer}>
        {STEPS.map((step, index) => {
          const isActive = index <= currentStepIndex;
          const isCurrent = step.key === currentStep;

          return (
            <div key={step.key} style={{ display: 'contents' }}>
              <div className={styles.stepItem} onClick={() => setCurrentStep(step.key)}>
                <Badge
                  appearance={isActive ? 'filled' : 'outline'}
                  color={isActive ? 'brand' : 'informative'}
                  size="small"
                >
                  {index + 1}
                </Badge>
                <Text className={`${styles.stepLabel} ${isCurrent ? styles.stepLabelActive : ''}`}>
                  {step.label}
                </Text>
              </div>
              {index < STEPS.length - 1 && (
                <div
                  className={`${styles.stepConnector} ${isActive ? styles.stepConnectorActive : ''}`}
                />
              )}
            </div>
          );
        })}
      </div>

      {/* Error Display */}
      {error && (
        <MessageBar intent="error" style={{ marginBottom: '24px' }}>
          <MessageBarBody>
            <DismissCircleRegular /> {error}
          </MessageBarBody>
        </MessageBar>
      )}

      {/* Step Content */}
      <Card className={styles.cardBody}>
        {/* Basic Info Step */}
        {currentStep === 'basic' && (
          <div className={styles.formSection}>
            <Text className={styles.sectionTitle}>Basic Information</Text>
            <Text className={styles.helperText}>
              Give your page a name and optional description.
            </Text>

            <Field label="Page Name *" required>
              <Input
                placeholder="e.g., Student Details"
                value={name}
                onChange={(_e, data) => setName(data.value)}
              />
            </Field>

            <Field label="Description">
              <Textarea
                placeholder="Optional description of what this page shows"
                rows={3}
                value={description}
                onChange={(_e, data) => setDescription(data.value)}
              />
            </Field>
          </div>
        )}

        {/* Page Type Step */}
        {currentStep === 'type' && (
          <div className={styles.formSection}>
            <Text className={styles.sectionTitle}>Choose Page Type</Text>
            <Text className={styles.helperText}>
              Select the type of page you want to create.
            </Text>

            <div className={styles.displayModeGrid}>
              {/* Lookup Type */}
              <div
                className={`${styles.displayModeOption} ${
                  pageType === 'lookup' ? styles.displayModeOptionSelected : ''
                }`}
                onClick={() => setPageType('lookup')}
              >
                <svg viewBox="0 0 120 80" style={{ width: '100%', height: '80px' }} fill="none" xmlns="http://www.w3.org/2000/svg">
                  {/* Filter panel on left */}
                  <rect x="2" y="2" width="28" height="76" rx="2" fill={tokens.colorNeutralBackground3} stroke={tokens.colorNeutralStroke1} strokeWidth="1"/>
                  <rect x="5" y="6" width="22" height="5" rx="1" fill={tokens.colorNeutralForeground3}/>
                  <rect x="5" y="14" width="22" height="4" rx="1" fill={tokens.colorNeutralBackground4}/>
                  <rect x="5" y="21" width="22" height="4" rx="1" fill={tokens.colorNeutralBackground4}/>
                  {/* Table on right */}
                  <rect x="34" y="2" width="84" height="76" rx="2" fill={tokens.colorNeutralBackground3} stroke={tokens.colorNeutralStroke1} strokeWidth="1"/>
                  <rect x="38" y="6" width="76" height="8" rx="1" fill={tokens.colorNeutralForeground3}/>
                  <line x1="38" y1="18" x2="114" y2="18" stroke={tokens.colorNeutralStroke1} strokeWidth="1"/>
                  <rect x="38" y="22" width="20" height="3" rx="1" fill={tokens.colorBrandBackground}/>
                  <rect x="62" y="22" width="25" height="3" rx="1" fill={tokens.colorNeutralBackground4}/>
                  <rect x="91" y="22" width="20" height="3" rx="1" fill={tokens.colorNeutralBackground4}/>
                  <line x1="38" y1="29" x2="114" y2="29" stroke={tokens.colorNeutralStroke2} strokeWidth="1"/>
                  <rect x="38" y="33" width="20" height="3" rx="1" fill={tokens.colorNeutralBackground4}/>
                  <rect x="62" y="33" width="25" height="3" rx="1" fill={tokens.colorNeutralBackground4}/>
                  <rect x="91" y="33" width="20" height="3" rx="1" fill={tokens.colorNeutralBackground4}/>
                  <line x1="38" y1="40" x2="114" y2="40" stroke={tokens.colorNeutralStroke2} strokeWidth="1"/>
                  <rect x="38" y="44" width="20" height="3" rx="1" fill={tokens.colorNeutralBackground4}/>
                  <rect x="62" y="44" width="25" height="3" rx="1" fill={tokens.colorNeutralBackground4}/>
                  <rect x="91" y="44" width="20" height="3" rx="1" fill={tokens.colorNeutralBackground4}/>
                </svg>
                <div className={styles.displayModeLabel}>
                  <Text weight="medium" size={200}>Lookup</Text>
                  <Text size={100} style={{ color: tokens.colorNeutralForeground2, display: 'block' }}>Browse and filter data</Text>
                </div>
              </div>

              {/* Report Type */}
              <div
                className={`${styles.displayModeOption} ${
                  pageType === 'report' ? styles.displayModeOptionSelected : ''
                }`}
                onClick={() => setPageType('report')}
              >
                <svg viewBox="0 0 120 80" style={{ width: '100%', height: '80px' }} fill="none" xmlns="http://www.w3.org/2000/svg">
                  {/* Main container */}
                  <rect x="2" y="2" width="116" height="76" rx="2" fill={tokens.colorNeutralBackground3} stroke={tokens.colorNeutralStroke1} strokeWidth="1"/>

                  {/* Two charts on top */}
                  {/* Left chart box */}
                  <rect x="6" y="6" width="52" height="32" rx="2" fill={tokens.colorNeutralBackground4} stroke={tokens.colorNeutralStroke2} strokeWidth="0.5"/>
                  {/* Left chart line */}
                  <polyline
                    points="10,30 18,24 26,26 34,18 42,20 50,12"
                    stroke={tokens.colorBrandBackground}
                    strokeWidth="2"
                    fill="none"
                    strokeLinecap="round"
                    strokeLinejoin="round"
                  />

                  {/* Right chart box */}
                  <rect x="62" y="6" width="52" height="32" rx="2" fill={tokens.colorNeutralBackground4} stroke={tokens.colorNeutralStroke2} strokeWidth="0.5"/>
                  {/* Right chart line */}
                  <polyline
                    points="66,22 74,28 82,20 90,24 98,16 106,18"
                    stroke={tokens.colorPaletteGreenBackground3}
                    strokeWidth="2"
                    fill="none"
                    strokeLinecap="round"
                    strokeLinejoin="round"
                  />

                  {/* List below */}
                  <rect x="6" y="42" width="108" height="6" rx="1" fill={tokens.colorNeutralForeground3}/>
                  <line x1="6" y1="52" x2="114" y2="52" stroke={tokens.colorNeutralStroke2} strokeWidth="0.5"/>
                  <rect x="6" y="56" width="30" height="3" rx="1" fill={tokens.colorNeutralBackground4}/>
                  <rect x="40" y="56" width="40" height="3" rx="1" fill={tokens.colorNeutralBackground4}/>
                  <rect x="84" y="56" width="24" height="3" rx="1" fill={tokens.colorNeutralBackground4}/>
                  <line x1="6" y1="63" x2="114" y2="63" stroke={tokens.colorNeutralStroke2} strokeWidth="0.5"/>
                  <rect x="6" y="67" width="30" height="3" rx="1" fill={tokens.colorNeutralBackground4}/>
                  <rect x="40" y="67" width="40" height="3" rx="1" fill={tokens.colorNeutralBackground4}/>
                  <rect x="84" y="67" width="24" height="3" rx="1" fill={tokens.colorNeutralBackground4}/>
                </svg>
                <div className={styles.displayModeLabel}>
                  <Text weight="medium" size={200}>Report</Text>
                  <Text size={100} style={{ color: tokens.colorNeutralForeground2, display: 'block' }}>Charts and dashboards</Text>
                </div>
              </div>
            </div>
          </div>
        )}

        {/* Primary List Step */}
        {currentStep === 'primary' && (
          <div className={styles.formSection}>
            <Text className={styles.sectionTitle}>Select Primary List</Text>
            <Text className={styles.helperText}>
              Choose the main data source for this page (e.g., Students, Customers).
            </Text>

            {enabledLists.length === 0 ? (
              <MessageBar intent="warning">
                <MessageBarBody>
                  No lists enabled. Enable lists in the Data page first.
                </MessageBarBody>
              </MessageBar>
            ) : (
              <RadioGroup
                value={primarySource?.listId || ''}
                onChange={(_e, data) => {
                  const list = enabledLists.find((l) => l.listId === data.value);
                  if (list) handlePrimarySourceSelect(list);
                }}
              >
                <div className={styles.sourceList}>
                  {enabledLists.map((list) => (
                    <div
                      key={`${list.siteId}-${list.listId}`}
                      className={`${styles.sourceItem} ${
                        primarySource?.listId === list.listId ? styles.sourceItemSelected : ''
                      }`}
                      onClick={() => handlePrimarySourceSelect(list)}
                    >
                      <Radio value={list.listId} />
                      <div className={styles.sourceInfo}>
                        <Text weight="medium">{list.listName}</Text>
                        <Text size={200} style={{ color: tokens.colorNeutralForeground2 }}>
                          {list.siteName}
                        </Text>
                      </div>
                    </div>
                  ))}
                </div>
              </RadioGroup>
            )}
          </div>
        )}

        {/* Display Columns Step */}
        {currentStep === 'columns' && (
          <div className={styles.formSection}>
            <Text className={styles.sectionTitle}>Select Display Columns</Text>
            <Text className={styles.helperText}>
              Choose which columns to show in the detail view. Drag to reorder.
            </Text>

            {loadingColumns ? (
              <div className={styles.loadingContainer}>
                <Spinner size="large" />
              </div>
            ) : availableColumns.length === 0 ? (
              <MessageBar intent="info">
                <MessageBarBody>
                  Select a primary list first to see available columns.
                </MessageBarBody>
              </MessageBar>
            ) : (
              <div className={styles.columnsGrid}>
                {/* Available Columns */}
                <div className={styles.columnPanel}>
                  <Text className={styles.columnPanelHeader}>Available Columns</Text>
                  <div className={styles.columnList}>
                    {availableColumns
                      .filter((col) => !displayColumns.some((dc) => dc.internalName === col.name))
                      .map((col) => (
                        <div
                          key={col.id}
                          className={styles.columnItem}
                          onClick={() => handleColumnToggle(col)}
                        >
                          <Text size={200}>{col.displayName}</Text>
                          {col.readOnly && (
                            <Badge appearance="outline" size="small">read-only</Badge>
                          )}
                        </div>
                      ))}
                  </div>
                </div>

                {/* Selected Columns */}
                <div className={styles.columnPanel}>
                  <Text className={styles.columnPanelHeader}>
                    Selected Columns ({displayColumns.length})
                  </Text>
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
                        className={`${styles.columnItem} ${styles.columnItemSelected} ${
                          draggedColIndex === index ? styles.columnItemDragging : ''
                        }`}
                      >
                        <ReOrderDotsVerticalRegular className={styles.dragHandle} />
                        <Text size={200} style={{ flex: 1 }}>{col.displayName}</Text>
                        <Button
                          appearance="subtle"
                          size="small"
                          icon={<DismissRegular />}
                          onClick={() =>
                            handleColumnToggle(
                              availableColumns.find((ac) => ac.name === col.internalName)!
                            )
                          }
                        />
                      </div>
                    ))}
                  </div>
                </div>
              </div>
            )}
          </div>
        )}

        {/* Search & Filters Step */}
        {currentStep === 'search' && (
          <div className={styles.formSection}>
            <Text className={styles.sectionTitle}>Configure Search & Filters</Text>
            <Text className={styles.helperText}>
              Configure how users can search and filter entities.
            </Text>

            {/* Table Columns */}
            <Field label="Table Columns *">
              <Text size={200} style={{ color: tokens.colorNeutralForeground2, marginBottom: '8px', display: 'block' }}>
                Columns to show in the search results table. Detail view will show all display columns.
              </Text>
              <div className={styles.badgeWrap}>
                {displayColumns.map((col) => (
                  <Badge
                    key={col.internalName}
                    className={styles.badgeItem}
                    appearance={searchConfig.tableColumns?.some((tc) => tc.internalName === col.internalName) ? 'filled' : 'outline'}
                    color={searchConfig.tableColumns?.some((tc) => tc.internalName === col.internalName) ? 'brand' : 'informative'}
                    onClick={() =>
                      setSearchConfig((prev) => {
                        const exists = prev.tableColumns?.some((tc) => tc.internalName === col.internalName);
                        return {
                          ...prev,
                          tableColumns: exists
                            ? prev.tableColumns?.filter((tc) => tc.internalName !== col.internalName)
                            : [...(prev.tableColumns || []), col],
                        };
                      })
                    }
                  >
                    {col.displayName}
                  </Badge>
                ))}
              </div>
              {(searchConfig.tableColumns?.length || 0) === 0 && (
                <Text size={200} style={{ color: tokens.colorPaletteYellowForeground1, marginTop: '8px' }}>
                  Select at least one column to display in the table.
                </Text>
              )}
            </Field>

            {/* Text Search Columns */}
            <Field label="Text Search Columns">
              <Text size={200} style={{ color: tokens.colorNeutralForeground2, marginBottom: '8px', display: 'block' }}>
                Columns to search when using the text search box.
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
            </Field>

            {/* Dropdown Filters */}
            <Field label="Dropdown Filters">
              <Text size={200} style={{ color: tokens.colorNeutralForeground2, marginBottom: '8px', display: 'block' }}>
                Choice columns that appear as dropdown filters.
              </Text>
              {choiceColumns.length === 0 ? (
                <Text size={200} style={{ color: tokens.colorNeutralForeground2 }}>
                  No choice or lookup columns available.
                </Text>
              ) : (
                <div className={styles.checkboxList}>
                  {choiceColumns.map((col) => (
                    <div key={col.id} style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                      <Checkbox
                        checked={searchConfig.filterColumns.some((f) => f.internalName === col.name)}
                        onChange={() => handleFilterColumnToggle(col)}
                        label={col.displayName}
                      />
                      <Badge appearance="outline" size="small">
                        {col.lookup ? 'lookup' : col.choice ? 'choice' : 'boolean'}
                      </Badge>
                    </div>
                  ))}
                </div>
              )}
            </Field>
          </div>
        )}

        {/* Related Lists Step */}
        {currentStep === 'related' && (
          <div className={styles.formSection}>
            <div className={styles.relatedHeader}>
              <div>
                <Text className={styles.sectionTitle}>Related Lists</Text>
                <Text className={styles.helperText} style={{ marginBottom: 0 }}>
                  Add lists that relate to the primary entity (e.g., Correspondence for Students).
                </Text>
              </div>
              <Button
                appearance="primary"
                size="small"
                icon={<AddRegular />}
                onClick={handleAddRelatedSection}
              >
                Add Related List
              </Button>
            </div>

            {relatedSections.length === 0 ? (
              <div className={styles.relatedEmpty}>
                <Text style={{ color: tokens.colorNeutralForeground2 }}>
                  No related lists configured yet.
                </Text>
                <Text size={200} style={{ color: tokens.colorNeutralForeground3, display: 'block', marginTop: '8px' }}>
                  Related lists let you show associated data like correspondence, activities, or orders.
                </Text>
              </div>
            ) : (
              <div>
                {relatedSections.map((section, index) => (
                  <RelatedSectionEditor
                    key={section.id}
                    section={section}
                    enabledLists={enabledLists}
                    primaryListId={primarySource?.listId || ''}
                    onUpdate={(updates) => handleUpdateRelatedSection(index, updates)}
                    onRemove={() => handleRemoveRelatedSection(index)}
                  />
                ))}
              </div>
            )}
          </div>
        )}

        {/* Review Step */}
        {currentStep === 'review' && (
          <div className={styles.formSection}>
            <Text className={styles.sectionTitle}>Review Configuration</Text>
            <Text className={styles.helperText}>
              Review your page configuration before saving.
            </Text>

            <div className={styles.reviewGrid}>
              <Card className={styles.reviewCard}>
                <Text className={styles.reviewCardTitle}>Basic Info</Text>
                <Text size={200}>
                  <strong>Name:</strong> {name}
                </Text>
                {description && (
                  <Text size={200} style={{ display: 'block' }}>
                    <strong>Description:</strong> {description}
                  </Text>
                )}
                <Text size={200} style={{ display: 'block' }}>
                  <strong>Type:</strong> {pageType === 'lookup' ? 'Lookup' : 'Report'}
                </Text>
              </Card>

              {/* Lookup-specific review cards */}
              {pageType === 'lookup' && (
                <>
                  <Card className={styles.reviewCard}>
                    <Text className={styles.reviewCardTitle}>Primary List</Text>
                    <Text size={200}>{primarySource?.listName || 'Not selected'}</Text>
                  </Card>

                  <Card className={styles.reviewCard}>
                    <Text className={styles.reviewCardTitle}>Display Columns</Text>
                    <Text size={200}>{displayColumns.length} columns selected</Text>
                    <div className={styles.badgeWrap} style={{ marginTop: '8px' }}>
                      {displayColumns.map((col) => (
                        <Badge key={col.internalName} appearance="outline" size="small">
                          {col.displayName}
                        </Badge>
                      ))}
                    </div>
                  </Card>

                  <Card className={styles.reviewCard}>
                    <Text className={styles.reviewCardTitle}>Search Configuration</Text>
                    <Text size={200} style={{ display: 'block' }}>
                      <strong>Table columns:</strong> {searchConfig.tableColumns?.length || 0}
                    </Text>
                    <Text size={200} style={{ display: 'block' }}>
                      <strong>Search columns:</strong> {searchConfig.textSearchColumns.length}
                    </Text>
                    <Text size={200} style={{ display: 'block' }}>
                      <strong>Filter columns:</strong> {searchConfig.filterColumns.length}
                    </Text>
                  </Card>

                  <Card className={styles.reviewCard}>
                    <Text className={styles.reviewCardTitle}>Related Sections</Text>
                    <Text size={200}>{relatedSections.length} related list(s)</Text>
                    {relatedSections.map((section) => (
                      <Text key={section.id} size={200} style={{ color: tokens.colorNeutralForeground2, display: 'block' }}>
                        - {section.title}: {section.source.listName || 'Not configured'}
                      </Text>
                    ))}
                  </Card>
                </>
              )}

              {/* Report-specific review - just shows that it's a report page */}
              {pageType === 'report' && (
                <Card className={styles.reviewCard}>
                  <Text className={styles.reviewCardTitle}>Page Type</Text>
                  <Text size={200}>Report pages display charts and dashboards.</Text>
                  <Text size={200} style={{ display: 'block', marginTop: '8px', color: tokens.colorNeutralForeground2 }}>
                    You can configure the report content after creating the page.
                  </Text>
                </Card>
              )}
            </div>
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
          {currentStep === 'review' ? (
            <Button
              appearance="primary"
              icon={saving ? <Spinner size="tiny" /> : <CheckmarkRegular />}
              onClick={handleSave}
              disabled={saving}
            >
              {saving ? 'Saving...' : 'Save Page'}
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
  );
}

// Sub-component for editing a related section
interface RelatedSectionEditorProps {
  section: RelatedSection;
  enabledLists: EnabledList[];
  primaryListId: string;
  onUpdate: (updates: Partial<RelatedSection>) => void;
  onRemove: () => void;
}

const useRelatedStyles = makeStyles({
  card: {
    marginBottom: '16px',
    padding: '16px',
  },
  header: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    marginBottom: '12px',
  },
  formSection: {
    display: 'flex',
    flexDirection: 'column',
    gap: '12px',
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
});

function RelatedSectionEditor({
  section,
  enabledLists,
  primaryListId,
  onUpdate,
  onRemove,
}: RelatedSectionEditorProps) {
  const styles = useRelatedStyles();
  const { instance, accounts } = useMsal();
  const [columns, setColumns] = useState<GraphListColumn[]>([]);
  const [loadingColumns, setLoadingColumns] = useState(false);
  const account = accounts[0];

  // Load columns when source changes
  useEffect(() => {
    if (!account || !section.source.siteId || !section.source.listId) {
      setColumns([]);
      return;
    }

    const loadColumns = async () => {
      setLoadingColumns(true);
      try {
        const cols = await getListColumns(
          instance,
          account,
          section.source.siteId,
          section.source.listId
        );
        setColumns(cols);
      } catch (err) {
        console.error('Failed to load columns:', err);
      } finally {
        setLoadingColumns(false);
      }
    };

    loadColumns();
  }, [instance, account, section.source.siteId, section.source.listId]);

  // Get lookup columns that reference the primary list
  const lookupColumns = columns.filter(
    (col) => col.lookup?.listId === primaryListId
  );

  return (
    <Card className={styles.card}>
      {/* Header with delete button */}
      <div className={styles.header}>
        <Text weight="medium">Related Section</Text>
        <Button
          appearance="subtle"
          size="small"
          icon={<DismissRegular />}
          onClick={onRemove}
          style={{ color: tokens.colorPaletteRedForeground1 }}
        />
      </div>

      <div className={styles.formSection}>
        <Field label="Section Title">
          <Input
            value={section.title}
            onChange={(_e, data) => onUpdate({ title: data.value })}
            placeholder="e.g., Correspondence"
            size="small"
          />
        </Field>

        <Field label="Related List">
          <Dropdown
            value={section.source.listName || ''}
            selectedOptions={section.source.listId ? [`${section.source.siteId}|${section.source.listId}`] : []}
            onOptionSelect={(_e, data) => {
              const [siteId, listId] = (data.optionValue as string).split('|');
              const list = enabledLists.find(
                (l) => l.siteId === siteId && l.listId === listId
              );
              if (list) {
                onUpdate({
                  source: {
                    siteId: list.siteId,
                    siteUrl: list.siteUrl,
                    listId: list.listId,
                    listName: list.listName,
                  },
                  lookupColumn: '',
                  displayColumns: [],
                });
              }
            }}
            placeholder="Select a list"
            size="small"
          >
            {enabledLists
              .filter((l) => l.listId !== primaryListId)
              .map((list) => (
                <Option
                  key={`${list.siteId}-${list.listId}`}
                  value={`${list.siteId}|${list.listId}`}
                >
                  {list.listName}
                </Option>
              ))}
          </Dropdown>
        </Field>

        {section.source.listId && (
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
                <Field label="Link Column">
                  <Text size={200} style={{ color: tokens.colorNeutralForeground2, marginBottom: '8px', display: 'block' }}>
                    Column that links to the primary list
                  </Text>
                  {lookupColumns.length === 0 ? (
                    <Text size={200} style={{ color: tokens.colorPaletteYellowForeground1 }}>
                      No lookup columns found that reference the primary list.
                    </Text>
                  ) : (
                    <Dropdown
                      value={lookupColumns.find(c => c.name === section.lookupColumn)?.displayName || ''}
                      selectedOptions={section.lookupColumn ? [section.lookupColumn] : []}
                      onOptionSelect={(_e, data) => onUpdate({ lookupColumn: data.optionValue as string })}
                      placeholder="Select link column"
                      size="small"
                    >
                      {lookupColumns.map((col) => (
                        <Option key={col.id} value={col.name}>
                          {col.displayName}
                        </Option>
                      ))}
                    </Dropdown>
                  )}
                </Field>

                <Field label="Display Columns">
                  <div className={styles.badgeWrap}>
                    {columns
                      .filter((col) => !col.hidden && col.name !== section.lookupColumn)
                      .slice(0, 10)
                      .map((col) => (
                        <Badge
                          key={col.id}
                          className={styles.badgeItem}
                          appearance={section.displayColumns.some((dc) => dc.internalName === col.name) ? 'filled' : 'outline'}
                          color={section.displayColumns.some((dc) => dc.internalName === col.name) ? 'brand' : 'informative'}
                          onClick={() => {
                            const exists = section.displayColumns.some(
                              (dc) => dc.internalName === col.name
                            );
                            onUpdate({
                              displayColumns: exists
                                ? section.displayColumns.filter(
                                    (dc) => dc.internalName !== col.name
                                  )
                                : [
                                    ...section.displayColumns,
                                    {
                                      internalName: col.name,
                                      displayName: col.displayName,
                                      editable: !col.readOnly,
                                    },
                                  ],
                            });
                          }}
                        >
                          {col.displayName}
                        </Badge>
                      ))}
                  </div>
                </Field>

                <Field label="Order By">
                  <div className={styles.sortRow}>
                    <Dropdown
                      className={styles.sortColumn}
                      value={section.displayColumns.find(c => c.internalName === section.defaultSort?.column)?.displayName || ''}
                      selectedOptions={section.defaultSort?.column ? [section.defaultSort.column] : []}
                      onOptionSelect={(_e, data) => {
                        if (data.optionValue) {
                          onUpdate({
                            defaultSort: {
                              column: data.optionValue as string,
                              direction: section.defaultSort?.direction || 'asc',
                            },
                          });
                        } else {
                          onUpdate({ defaultSort: undefined });
                        }
                      }}
                      placeholder="None"
                      size="small"
                    >
                      <Option value="">None</Option>
                      {section.displayColumns.map((col) => (
                        <Option key={col.internalName} value={col.internalName}>
                          {col.displayName}
                        </Option>
                      ))}
                    </Dropdown>
                    {section.defaultSort?.column && (
                      <Dropdown
                        className={styles.sortDirection}
                        value={section.defaultSort.direction === 'asc' ? 'Ascending' : 'Descending'}
                        selectedOptions={[section.defaultSort.direction]}
                        onOptionSelect={(_e, data) =>
                          onUpdate({
                            defaultSort: {
                              column: section.defaultSort!.column,
                              direction: data.optionValue as 'asc' | 'desc',
                            },
                          })
                        }
                        size="small"
                      >
                        <Option value="asc">Ascending</Option>
                        <Option value="desc">Descending</Option>
                      </Dropdown>
                    )}
                  </div>
                </Field>

                <Divider style={{ margin: '8px 0' }} />

                <div className={styles.permissionsRow}>
                  <Checkbox
                    checked={section.allowCreate}
                    onChange={(_e, data) => onUpdate({ allowCreate: data.checked === true })}
                    label="Allow Create"
                  />
                  <Checkbox
                    checked={section.allowEdit}
                    onChange={(_e, data) => onUpdate({ allowEdit: data.checked === true })}
                    label="Allow Edit"
                  />
                  <Checkbox
                    checked={section.allowDelete}
                    onChange={(_e, data) => onUpdate({ allowDelete: data.checked === true })}
                    label="Allow Delete"
                  />
                </div>
              </>
            )}
          </>
        )}
      </div>
    </Card>
  );
}

export default PageEditor;
