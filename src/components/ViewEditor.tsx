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
} from '@fluentui/react-components';
import {
  WarningRegular,
  ReOrderDotsVerticalRegular,
  DismissRegular,
  AddRegular,
  ArrowLeftRegular,
  ArrowRightRegular,
  CheckmarkRegular,
} from '@fluentui/react-icons';
import { useSettings, type EnabledList } from '../contexts/SettingsContext';
import { getListColumns, type GraphListColumn } from '../auth/graphClient';
import type {
  ViewDefinition,
  ViewSource,
  ViewColumn,
  ViewFilter,
  ViewSorting,
  ViewSortRule,
  AggregationType,
  FilterOperator,
} from '../types/view';

interface ViewEditorProps {
  initialView?: ViewDefinition;
  onSave: (view: ViewDefinition) => Promise<void>;
  onCancel: () => void;
}

type Step = 'basic' | 'sources' | 'columns' | 'groupby' | 'filters' | 'sorting';

const UNION_STEPS: { key: Step; label: string }[] = [
  { key: 'basic', label: 'Basic Info' },
  { key: 'sources', label: 'Data Sources' },
  { key: 'columns', label: 'Columns' },
  { key: 'filters', label: 'Filters' },
  { key: 'sorting', label: 'Sorting' },
];

const AGGREGATE_STEPS: { key: Step; label: string }[] = [
  { key: 'basic', label: 'Basic Info' },
  { key: 'sources', label: 'Data Sources' },
  { key: 'columns', label: 'Columns' },
  { key: 'groupby', label: 'Group By' },
  { key: 'filters', label: 'Filters' },
  { key: 'sorting', label: 'Sorting' },
];

const AGGREGATION_OPTIONS: { value: AggregationType; label: string }[] = [
  { value: 'count', label: 'Count' },
  { value: 'sum', label: 'Sum' },
  { value: 'avg', label: 'Average' },
  { value: 'min', label: 'Min' },
  { value: 'max', label: 'Max' },
];

const FILTER_OPERATORS: { value: FilterOperator; label: string }[] = [
  { value: 'eq', label: 'Equals' },
  { value: 'ne', label: 'Not equals' },
  { value: 'gt', label: 'Greater than' },
  { value: 'lt', label: 'Less than' },
  { value: 'contains', label: 'Contains' },
];

interface ColumnWithMeta extends GraphListColumn {
  sourceListId: string;
  sourceListName: string;
}

const useStyles = makeStyles({
  container: {
    display: 'flex',
    flexDirection: 'column',
    gap: '24px',
  },
  stepsContainer: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    gap: '8px',
  },
  stepItem: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
    cursor: 'pointer',
  },
  stepLabel: {
    fontSize: tokens.fontSizeBase200,
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
  formSection: {
    display: 'flex',
    flexDirection: 'column',
    gap: '16px',
  },
  modeOptions: {
    display: 'flex',
    gap: '24px',
  },
  modeOption: {
    display: 'flex',
    alignItems: 'flex-start',
    gap: '8px',
    cursor: 'pointer',
  },
  modeLabel: {
    display: 'flex',
    flexDirection: 'column',
    gap: '2px',
  },
  modeSublabel: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground2,
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
    backgroundColor: tokens.colorNeutralBackground2,
    cursor: 'pointer',
    ':hover': {
      backgroundColor: tokens.colorNeutralBackground3,
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
  sourceName: {
    fontWeight: tokens.fontWeightMedium,
  },
  sourceSite: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground2,
  },
  loadingContainer: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    padding: '32px',
  },
  sourceCard: {
    marginBottom: '12px',
  },
  sourceCardHeader: {
    fontWeight: tokens.fontWeightMedium,
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground2,
    marginBottom: '8px',
  },
  columnsGrid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(2, 1fr)',
    gap: '8px',
  },
  columnCheckbox: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
  },
  columnConfig: {
    marginTop: '16px',
  },
  columnConfigHeader: {
    fontWeight: tokens.fontWeightMedium,
    marginBottom: '8px',
  },
  columnConfigHint: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground2,
    marginBottom: '12px',
  },
  columnConfigList: {
    display: 'flex',
    flexDirection: 'column',
    gap: '8px',
  },
  columnItem: {
    display: 'flex',
    alignItems: 'center',
    gap: '12px',
    padding: '8px',
    borderRadius: tokens.borderRadiusMedium,
    backgroundColor: tokens.colorNeutralBackground1,
    border: `1px solid transparent`,
    cursor: 'move',
    '&:hover': {
      border: `1px solid ${tokens.colorNeutralStroke1}`,
    },
  },
  columnItemDragging: {
    opacity: 0.5,
    border: `1px solid ${tokens.colorBrandStroke1}`,
  },
  dragHandle: {
    color: tokens.colorNeutralForeground3,
    flexShrink: 0,
  },
  columnInput: {
    flex: 1,
  },
  groupByList: {
    display: 'flex',
    flexDirection: 'column',
    gap: '8px',
  },
  groupByItem: {
    display: 'flex',
    alignItems: 'center',
    gap: '12px',
    padding: '12px',
    borderRadius: tokens.borderRadiusMedium,
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    backgroundColor: tokens.colorNeutralBackground2,
    cursor: 'pointer',
    ':hover': {
      backgroundColor: tokens.colorNeutralBackground3,
    },
  },
  groupByItemSelected: {
    backgroundColor: tokens.colorBrandBackground2,
    border: `1px solid ${tokens.colorBrandStroke1}`,
  },
  groupByInfo: {
    flex: 1,
    display: 'flex',
    flexDirection: 'column',
    gap: '2px',
  },
  groupByHint: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground2,
  },
  filterList: {
    display: 'flex',
    flexDirection: 'column',
    gap: '12px',
  },
  filterRow: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
  },
  filterColumn: {
    flex: 1,
    minWidth: '120px',
  },
  filterOperator: {
    minWidth: '120px',
  },
  filterValue: {
    flex: 1,
    minWidth: '120px',
  },
  emptyFilters: {
    textAlign: 'center',
    padding: '32px',
    color: tokens.colorNeutralForeground2,
  },
  sortCard: {
    marginBottom: '16px',
  },
  sortCardHeader: {
    fontWeight: tokens.fontWeightMedium,
    marginBottom: '12px',
  },
  sortRow: {
    display: 'flex',
    gap: '12px',
    alignItems: 'flex-end',
  },
  sortColumn: {
    flex: 1,
  },
  sortDirection: {
    minWidth: '200px',
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
  helperText: {
    color: tokens.colorNeutralForeground2,
  },
});

function ViewEditor({ initialView, onSave, onCancel }: ViewEditorProps) {
  const styles = useStyles();
  const { instance, accounts } = useMsal();
  const { enabledLists } = useSettings();
  const account = accounts[0];

  // Current step
  const [currentStep, setCurrentStep] = useState<Step>('basic');

  // View state
  const [name, setName] = useState(initialView?.name || '');
  const [description, setDescription] = useState(initialView?.description || '');
  const [mode, setMode] = useState<'union' | 'aggregate'>(initialView?.mode || 'union');
  const [sources, setSources] = useState<ViewSource[]>(initialView?.sources || []);
  const [columns, setColumns] = useState<ViewColumn[]>(initialView?.columns || []);
  const [groupBy, setGroupBy] = useState<string[]>(initialView?.groupBy || []);
  const [filters, setFilters] = useState<ViewFilter[]>(initialView?.filters || []);
  const [sorting, setSorting] = useState<ViewSorting>(initialView?.sorting || []);

  // Get steps based on mode
  const STEPS = mode === 'aggregate' ? AGGREGATE_STEPS : UNION_STEPS;

  // Available columns from selected sources
  const [availableColumns, setAvailableColumns] = useState<ColumnWithMeta[]>([]);
  const [loadingColumns, setLoadingColumns] = useState(false);

  // Saving state
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState<string | null>(null);

  // Drag and drop state for column reordering
  const [draggedColIndex, setDraggedColIndex] = useState<number | null>(null);

  // Load columns when sources change
  useEffect(() => {
    if (!account || sources.length === 0) {
      setAvailableColumns([]);
      return;
    }

    const loadColumns = async () => {
      setLoadingColumns(true);
      try {
        const allColumns: ColumnWithMeta[] = [];

        for (const source of sources) {
          const cols = await getListColumns(instance, account, source.siteId, source.listId);
          allColumns.push(
            ...cols.map((col) => ({
              ...col,
              sourceListId: source.listId,
              sourceListName: source.listName,
            }))
          );
        }

        setAvailableColumns(allColumns);
      } catch (err) {
        console.error('Failed to load columns:', err);
      } finally {
        setLoadingColumns(false);
      }
    };

    loadColumns();
  }, [instance, account, sources]);

  const handleSourceToggle = useCallback(
    (list: EnabledList) => {
      setSources((prev) => {
        const exists = prev.some((s) => s.listId === list.listId && s.siteId === list.siteId);
        if (exists) {
          // Also remove columns from this source
          setColumns((cols) => cols.filter((c) => c.sourceListId !== list.listId));
          return prev.filter((s) => !(s.listId === list.listId && s.siteId === list.siteId));
        }
        return [
          ...prev,
          {
            siteId: list.siteId,
            listId: list.listId,
            listName: list.listName,
          },
        ];
      });
    },
    []
  );

  const handleColumnToggle = useCallback(
    (col: ColumnWithMeta) => {
      setColumns((prev) => {
        const exists = prev.some(
          (c) => c.internalName === col.name && c.sourceListId === col.sourceListId
        );
        if (exists) {
          return prev.filter(
            (c) => !(c.internalName === col.name && c.sourceListId === col.sourceListId)
          );
        }
        return [
          ...prev,
          {
            sourceListId: col.sourceListId,
            internalName: col.name,
            displayName: col.displayName,
          },
        ];
      });
    },
    []
  );

  const handleColumnDisplayNameChange = useCallback(
    (sourceListId: string, internalName: string, displayName: string) => {
      setColumns((prev) =>
        prev.map((c) =>
          c.sourceListId === sourceListId && c.internalName === internalName
            ? { ...c, displayName }
            : c
        )
      );
    },
    []
  );

  const handleColumnAggregationChange = useCallback(
    (sourceListId: string, internalName: string, aggregation: AggregationType | undefined) => {
      setColumns((prev) =>
        prev.map((c) =>
          c.sourceListId === sourceListId && c.internalName === internalName
            ? { ...c, aggregation }
            : c
        )
      );
    },
    []
  );

  const handleGroupByToggle = useCallback((internalName: string) => {
    setGroupBy((prev) => {
      if (prev.includes(internalName)) {
        return prev.filter((n) => n !== internalName);
      }
      return [...prev, internalName];
    });
  }, []);

  const handleColumnReorder = useCallback((fromIndex: number, toIndex: number) => {
    if (fromIndex === toIndex) return;
    setColumns((prev) => {
      const newColumns = [...prev];
      const [removed] = newColumns.splice(fromIndex, 1);
      newColumns.splice(toIndex, 0, removed);
      return newColumns;
    });
  }, []);

  const handleAddFilter = useCallback(() => {
    if (columns.length === 0) return;
    setFilters((prev) => [
      ...prev,
      {
        column: columns[0].internalName,
        operator: 'eq',
        value: '',
      },
    ]);
  }, [columns]);

  const handleRemoveFilter = useCallback((index: number) => {
    setFilters((prev) => prev.filter((_, i) => i !== index));
  }, []);

  const handleFilterChange = useCallback(
    (index: number, field: keyof ViewFilter, value: string) => {
      setFilters((prev) =>
        prev.map((f, i) =>
          i === index ? { ...f, [field]: value } : f
        )
      );
    },
    []
  );

  const handleSave = async () => {
    if (!name.trim()) {
      setError('Please enter a name for the view');
      setCurrentStep('basic');
      return;
    }

    if (sources.length === 0) {
      setError('Please select at least one data source');
      setCurrentStep('sources');
      return;
    }

    if (columns.length === 0) {
      setError('Please select at least one column');
      setCurrentStep('columns');
      return;
    }

    setSaving(true);
    setError(null);

    try {
      const view: ViewDefinition = {
        id: initialView?.id,
        name: name.trim(),
        description: description.trim() || undefined,
        mode,
        sources,
        columns,
        groupBy: mode === 'aggregate' && groupBy.length > 0 ? groupBy : undefined,
        filters: filters.length > 0 ? filters : undefined,
        sorting: sorting.length > 0 ? sorting : undefined,
      };

      await onSave(view);
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Failed to save view');
    } finally {
      setSaving(false);
    }
  };

  const canProceed = (step: Step): boolean => {
    switch (step) {
      case 'basic':
        return name.trim().length > 0;
      case 'sources':
        return sources.length > 0;
      case 'columns':
        return columns.length > 0;
      default:
        return true;
    }
  };

  const goToNextStep = () => {
    const currentIndex = STEPS.findIndex((s) => s.key === currentStep);
    if (currentIndex < STEPS.length - 1) {
      setCurrentStep(STEPS[currentIndex + 1].key);
    }
  };

  const goToPreviousStep = () => {
    const currentIndex = STEPS.findIndex((s) => s.key === currentStep);
    if (currentIndex > 0) {
      setCurrentStep(STEPS[currentIndex - 1].key);
    }
  };

  const renderStepContent = () => {
    switch (currentStep) {
      case 'basic':
        return (
          <div className={styles.formSection}>
            <Field label="View Name *" required>
              <Input
                value={name}
                onChange={(_e, data) => setName(data.value)}
                placeholder="Enter view name"
              />
            </Field>

            <Field label="Description">
              <Textarea
                value={description}
                onChange={(_e, data) => setDescription(data.value)}
                placeholder="Optional description"
                rows={3}
              />
            </Field>

            <Field label="View Mode *">
              <RadioGroup
                value={mode}
                onChange={(_e, data) => setMode(data.value as 'union' | 'aggregate')}
              >
                <div className={styles.modeOptions}>
                  <div className={styles.modeOption}>
                    <Radio value="union" />
                    <div className={styles.modeLabel}>
                      <Text weight="medium">Union</Text>
                      <Text className={styles.modeSublabel}>
                        Stack all rows from selected lists into one table
                      </Text>
                    </div>
                  </div>
                  <div className={styles.modeOption}>
                    <Radio value="aggregate" />
                    <div className={styles.modeLabel}>
                      <Text weight="medium">Aggregate</Text>
                      <Text className={styles.modeSublabel}>
                        Show counts, sums, or other aggregations
                      </Text>
                    </div>
                  </div>
                </div>
              </RadioGroup>
            </Field>
          </div>
        );

      case 'sources':
        return (
          <div className={styles.formSection}>
            <Text className={styles.helperText}>
              Select the SharePoint lists to include in this view.
            </Text>

            {enabledLists.length === 0 ? (
              <MessageBar intent="warning">
                <MessageBarBody>
                  <WarningRegular /> No lists enabled. Please enable some lists in the Data page first.
                </MessageBarBody>
              </MessageBar>
            ) : (
              <div className={styles.sourceList}>
                {enabledLists.map((list) => {
                  const isSelected = sources.some(
                    (s) => s.listId === list.listId && s.siteId === list.siteId
                  );
                  return (
                    <div
                      key={`${list.siteId}:${list.listId}`}
                      className={`${styles.sourceItem} ${isSelected ? styles.sourceItemSelected : ''}`}
                      onClick={() => handleSourceToggle(list)}
                    >
                      <Checkbox checked={isSelected} />
                      <div className={styles.sourceInfo}>
                        <Text className={styles.sourceName}>{list.listName}</Text>
                        <Text className={styles.sourceSite}>{list.siteName}</Text>
                      </div>
                    </div>
                  );
                })}
              </div>
            )}
          </div>
        );

      case 'columns':
        return (
          <div className={styles.formSection}>
            <Text className={styles.helperText}>
              Select the columns to display in this view.
            </Text>

            {loadingColumns ? (
              <div className={styles.loadingContainer}>
                <Spinner size="large" />
              </div>
            ) : availableColumns.length === 0 ? (
              <MessageBar intent="info">
                <MessageBarBody>
                  Select data sources first to see available columns.
                </MessageBarBody>
              </MessageBar>
            ) : (
              <div>
                {/* Available columns grouped by source */}
                {sources.map((source) => {
                  const sourceCols = availableColumns.filter(
                    (c) => c.sourceListId === source.listId
                  );
                  if (sourceCols.length === 0) return null;

                  return (
                    <Card key={source.listId} className={styles.sourceCard}>
                      <Text className={styles.sourceCardHeader}>{source.listName}</Text>
                      <div className={styles.columnsGrid}>
                        {sourceCols.map((col) => {
                          const isSelected = columns.some(
                            (c) =>
                              c.internalName === col.name && c.sourceListId === col.sourceListId
                          );
                          return (
                            <Checkbox
                              key={`${col.sourceListId}:${col.name}`}
                              checked={isSelected}
                              onChange={() => handleColumnToggle(col)}
                              label={col.displayName}
                            />
                          );
                        })}
                      </div>
                    </Card>
                  );
                })}

                {/* Selected columns configuration */}
                {columns.length > 0 && (
                  <Card className={styles.columnConfig}>
                    <Text className={styles.columnConfigHeader}>Column Configuration</Text>
                    <Text className={styles.columnConfigHint}>Drag to reorder columns</Text>
                    <div className={styles.columnConfigList}>
                      {columns.map((col, index) => (
                        <div
                          key={`${col.sourceListId}:${col.internalName}`}
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
                          className={`${styles.columnItem} ${
                            draggedColIndex === index ? styles.columnItemDragging : ''
                          }`}
                        >
                          <ReOrderDotsVerticalRegular className={styles.dragHandle} />
                          <Input
                            className={styles.columnInput}
                            value={col.displayName}
                            onChange={(_e, data) =>
                              handleColumnDisplayNameChange(
                                col.sourceListId,
                                col.internalName,
                                data.value
                              )
                            }
                            onClick={(e) => e.stopPropagation()}
                            placeholder="Display name"
                            size="small"
                          />
                          {mode === 'aggregate' && (
                            <Dropdown
                              value={col.aggregation || ''}
                              selectedOptions={col.aggregation ? [col.aggregation] : []}
                              onOptionSelect={(_e, data) =>
                                handleColumnAggregationChange(
                                  col.sourceListId,
                                  col.internalName,
                                  (data.optionValue as AggregationType) || undefined
                                )
                              }
                              onClick={(e) => e.stopPropagation()}
                              placeholder="No aggregation"
                              size="small"
                            >
                              <Option value="">No aggregation</Option>
                              {AGGREGATION_OPTIONS.map((opt) => (
                                <Option key={opt.value} value={opt.value}>
                                  {opt.label}
                                </Option>
                              ))}
                            </Dropdown>
                          )}
                        </div>
                      ))}
                    </div>
                  </Card>
                )}
              </div>
            )}
          </div>
        );

      case 'groupby':
        return (
          <div className={styles.formSection}>
            <Text className={styles.helperText}>
              Select columns to group by. Rows will be grouped by these columns, and aggregations
              will be computed for each group.
            </Text>

            {columns.length === 0 ? (
              <MessageBar intent="info">
                <MessageBarBody>
                  Select columns first to configure grouping.
                </MessageBarBody>
              </MessageBar>
            ) : (
              <div className={styles.groupByList}>
                {columns.map((col) => {
                  const isGroupBy = groupBy.includes(col.internalName);
                  return (
                    <div
                      key={`${col.sourceListId}:${col.internalName}`}
                      className={`${styles.groupByItem} ${isGroupBy ? styles.groupByItemSelected : ''}`}
                      onClick={() => handleGroupByToggle(col.internalName)}
                    >
                      <Checkbox checked={isGroupBy} />
                      <div className={styles.groupByInfo}>
                        <Text weight="medium">{col.displayName}</Text>
                        <Text className={styles.groupByHint}>
                          {isGroupBy ? 'Group by this column' : col.aggregation ? `Aggregate: ${col.aggregation}` : 'No aggregation'}
                        </Text>
                      </div>
                    </div>
                  );
                })}
              </div>
            )}

            {groupBy.length === 0 && columns.length > 0 && (
              <MessageBar intent="warning">
                <MessageBarBody>
                  <WarningRegular /> Without group by columns, the view will show a single row with aggregated totals.
                </MessageBarBody>
              </MessageBar>
            )}
          </div>
        );

      case 'filters':
        return (
          <div className={styles.formSection}>
            <Text className={styles.helperText}>
              Add filters to narrow down the data (optional).
            </Text>

            {filters.length === 0 ? (
              <div className={styles.emptyFilters}>
                <Text>No filters added</Text>
              </div>
            ) : (
              <div className={styles.filterList}>
                {filters.map((filter, index) => (
                  <div key={index} className={styles.filterRow}>
                    <Dropdown
                      className={styles.filterColumn}
                      value={columns.find(c => c.internalName === filter.column)?.displayName || filter.column}
                      selectedOptions={[filter.column]}
                      onOptionSelect={(_e, data) => handleFilterChange(index, 'column', data.optionValue as string)}
                      size="small"
                    >
                      {columns.map((col) => (
                        <Option key={`${col.sourceListId}:${col.internalName}`} value={col.internalName}>
                          {col.displayName}
                        </Option>
                      ))}
                    </Dropdown>
                    <Dropdown
                      className={styles.filterOperator}
                      value={FILTER_OPERATORS.find(op => op.value === filter.operator)?.label || filter.operator}
                      selectedOptions={[filter.operator]}
                      onOptionSelect={(_e, data) => handleFilterChange(index, 'operator', data.optionValue as string)}
                      size="small"
                    >
                      {FILTER_OPERATORS.map((op) => (
                        <Option key={op.value} value={op.value}>
                          {op.label}
                        </Option>
                      ))}
                    </Dropdown>
                    <Input
                      className={styles.filterValue}
                      value={filter.value}
                      onChange={(_e, data) => handleFilterChange(index, 'value', data.value)}
                      placeholder="Value"
                      size="small"
                    />
                    <Button
                      appearance="subtle"
                      icon={<DismissRegular />}
                      onClick={() => handleRemoveFilter(index)}
                      size="small"
                    />
                  </div>
                ))}
              </div>
            )}

            <Button
              appearance="outline"
              icon={<AddRegular />}
              onClick={handleAddFilter}
              disabled={columns.length === 0}
              size="small"
            >
              Add Filter
            </Button>
          </div>
        );

      case 'sorting': {
        const updateSortRule = (index: number, updates: Partial<ViewSortRule>) => {
          setSorting((prev) => {
            const newSorting = [...prev];
            if (newSorting[index]) {
              newSorting[index] = { ...newSorting[index], ...updates };
            }
            return newSorting;
          });
        };

        const setPrimarySortColumn = (column: string) => {
          if (!column) {
            setSorting([]);
          } else {
            setSorting((prev) => {
              if (prev.length === 0) {
                return [{ column, direction: 'asc' }];
              }
              return [{ ...prev[0], column }, ...prev.slice(1)];
            });
          }
        };

        const setSecondarySortColumn = (column: string) => {
          if (!column) {
            setSorting((prev) => prev.slice(0, 1));
          } else {
            setSorting((prev) => {
              if (prev.length === 0) return prev;
              if (prev.length === 1) {
                return [...prev, { column, direction: 'asc' }];
              }
              return [prev[0], { ...prev[1], column }];
            });
          }
        };

        return (
          <div className={styles.formSection}>
            <Text className={styles.helperText}>
              Choose how to sort the data. You can add a primary sort and an optional secondary sort.
            </Text>

            {/* Primary Sort */}
            <Card className={styles.sortCard}>
              <Text className={styles.sortCardHeader}>Primary Sort</Text>
              <div className={styles.sortRow}>
                <Field label="Column" className={styles.sortColumn}>
                  <Dropdown
                    value={sorting[0]?.column ? columns.find(c => c.internalName === sorting[0].column)?.displayName : ''}
                    selectedOptions={sorting[0]?.column ? [sorting[0].column] : []}
                    onOptionSelect={(_e, data) => setPrimarySortColumn(data.optionValue as string)}
                    placeholder="No sorting"
                  >
                    <Option value="">No sorting</Option>
                    {columns.map((col) => (
                      <Option key={`${col.sourceListId}:${col.internalName}`} value={col.internalName}>
                        {col.displayName}
                      </Option>
                    ))}
                  </Dropdown>
                </Field>
                {sorting[0] && (
                  <Field label="Direction" className={styles.sortDirection}>
                    <Dropdown
                      value={sorting[0].direction === 'asc' ? 'Ascending (A-Z, 0-9)' : 'Descending (Z-A, 9-0)'}
                      selectedOptions={[sorting[0].direction]}
                      onOptionSelect={(_e, data) => updateSortRule(0, { direction: data.optionValue as 'asc' | 'desc' })}
                    >
                      <Option value="asc">Ascending (A-Z, 0-9)</Option>
                      <Option value="desc">Descending (Z-A, 9-0)</Option>
                    </Dropdown>
                  </Field>
                )}
              </div>
            </Card>

            {/* Secondary Sort */}
            {sorting.length > 0 && (
              <Card className={styles.sortCard}>
                <Text className={styles.sortCardHeader}>Secondary Sort (Optional)</Text>
                <div className={styles.sortRow}>
                  <Field label="Column" className={styles.sortColumn}>
                    <Dropdown
                      value={sorting[1]?.column ? columns.find(c => c.internalName === sorting[1].column)?.displayName : ''}
                      selectedOptions={sorting[1]?.column ? [sorting[1].column] : []}
                      onOptionSelect={(_e, data) => setSecondarySortColumn(data.optionValue as string)}
                      placeholder="No secondary sort"
                    >
                      <Option value="">No secondary sort</Option>
                      {columns
                        .filter((col) => col.internalName !== sorting[0]?.column)
                        .map((col) => (
                          <Option key={`${col.sourceListId}:${col.internalName}`} value={col.internalName}>
                            {col.displayName}
                          </Option>
                        ))}
                    </Dropdown>
                  </Field>
                  {sorting[1] && (
                    <Field label="Direction" className={styles.sortDirection}>
                      <Dropdown
                        value={sorting[1].direction === 'asc' ? 'Ascending (A-Z, 0-9)' : 'Descending (Z-A, 9-0)'}
                        selectedOptions={[sorting[1].direction]}
                        onOptionSelect={(_e, data) => updateSortRule(1, { direction: data.optionValue as 'asc' | 'desc' })}
                      >
                        <Option value="asc">Ascending (A-Z, 0-9)</Option>
                        <Option value="desc">Descending (Z-A, 9-0)</Option>
                      </Dropdown>
                    </Field>
                  )}
                </div>
              </Card>
            )}
          </div>
        );
      }
    }
  };

  const isLastStep = currentStep === STEPS[STEPS.length - 1].key;

  return (
    <div className={styles.container}>
      {/* Steps indicator */}
      <div className={styles.stepsContainer}>
        {STEPS.map((step, index) => {
          const currentIndex = STEPS.findIndex((s) => s.key === currentStep);
          const isActive = index <= currentIndex;
          const isCurrent = step.key === currentStep;

          return (
            <div key={step.key} style={{ display: 'contents' }}>
              <div className={styles.stepItem} onClick={() => setCurrentStep(step.key)}>
                <Badge
                  appearance={isActive ? 'filled' : 'outline'}
                  color={isActive ? 'brand' : 'informative'}
                  size="medium"
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

      {/* Error message */}
      {error && (
        <MessageBar intent="error">
          <MessageBarBody>
            <WarningRegular /> {error}
          </MessageBarBody>
        </MessageBar>
      )}

      {/* Step content */}
      <Card>
        {renderStepContent()}
      </Card>

      {/* Navigation buttons */}
      <div className={styles.navigation}>
        <div>
          {currentStep !== STEPS[0].key && (
            <Button appearance="subtle" icon={<ArrowLeftRegular />} onClick={goToPreviousStep}>
              Previous
            </Button>
          )}
        </div>

        <div className={styles.navRight}>
          <Button appearance="subtle" onClick={onCancel}>
            Cancel
          </Button>
          {isLastStep ? (
            <Button
              appearance="primary"
              icon={saving ? <Spinner size="tiny" /> : <CheckmarkRegular />}
              onClick={handleSave}
              disabled={saving}
            >
              {saving ? 'Saving...' : 'Save View'}
            </Button>
          ) : (
            <Button
              appearance="primary"
              icon={<ArrowRightRegular />}
              iconPosition="after"
              onClick={goToNextStep}
              disabled={!canProceed(currentStep)}
            >
              Next
            </Button>
          )}
        </div>
      </div>
    </div>
  );
}

export default ViewEditor;
