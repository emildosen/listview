import { useState, useEffect, useCallback } from 'react';
import { useMsal } from '@azure/msal-react';
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

function ViewEditor({ initialView, onSave, onCancel }: ViewEditorProps) {
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
          <div className="space-y-4">
            <div className="form-control">
              <label className="label">
                <span className="label-text font-medium">View Name *</span>
              </label>
              <input
                type="text"
                value={name}
                onChange={(e) => setName(e.target.value)}
                placeholder="Enter view name"
                className="input input-bordered w-full"
              />
            </div>

            <div className="form-control">
              <label className="label">
                <span className="label-text font-medium">Description</span>
              </label>
              <textarea
                value={description}
                onChange={(e) => setDescription(e.target.value)}
                placeholder="Optional description"
                className="textarea textarea-bordered w-full"
                rows={3}
              />
            </div>

            <div className="form-control">
              <label className="label">
                <span className="label-text font-medium">View Mode *</span>
              </label>
              <div className="flex gap-4">
                <label className="flex items-center gap-2 cursor-pointer">
                  <input
                    type="radio"
                    name="mode"
                    value="union"
                    checked={mode === 'union'}
                    onChange={() => setMode('union')}
                    className="radio radio-primary"
                  />
                  <div>
                    <div className="font-medium">Union</div>
                    <div className="text-sm text-base-content/60">
                      Stack all rows from selected lists into one table
                    </div>
                  </div>
                </label>
                <label className="flex items-center gap-2 cursor-pointer">
                  <input
                    type="radio"
                    name="mode"
                    value="aggregate"
                    checked={mode === 'aggregate'}
                    onChange={() => setMode('aggregate')}
                    className="radio radio-primary"
                  />
                  <div>
                    <div className="font-medium">Aggregate</div>
                    <div className="text-sm text-base-content/60">
                      Show counts, sums, or other aggregations
                    </div>
                  </div>
                </label>
              </div>
            </div>
          </div>
        );

      case 'sources':
        return (
          <div className="space-y-4">
            <p className="text-base-content/60">
              Select the SharePoint lists to include in this view.
            </p>

            {enabledLists.length === 0 ? (
              <div className="alert alert-warning">
                <svg
                  xmlns="http://www.w3.org/2000/svg"
                  fill="none"
                  viewBox="0 0 24 24"
                  strokeWidth={1.5}
                  stroke="currentColor"
                  className="w-5 h-5"
                >
                  <path
                    strokeLinecap="round"
                    strokeLinejoin="round"
                    d="M12 9v3.75m9-.75a9 9 0 1 1-18 0 9 9 0 0 1 18 0Zm-9 3.75h.008v.008H12v-.008Z"
                  />
                </svg>
                <span>No lists enabled. Please enable some lists in the Data page first.</span>
              </div>
            ) : (
              <div className="space-y-2">
                {enabledLists.map((list) => {
                  const isSelected = sources.some(
                    (s) => s.listId === list.listId && s.siteId === list.siteId
                  );
                  return (
                    <label
                      key={`${list.siteId}:${list.listId}`}
                      className={`flex items-center gap-3 p-3 rounded-lg cursor-pointer transition-colors ${
                        isSelected ? 'bg-primary/10 border-primary' : 'bg-base-200 hover:bg-base-300'
                      } border`}
                    >
                      <input
                        type="checkbox"
                        checked={isSelected}
                        onChange={() => handleSourceToggle(list)}
                        className="checkbox checkbox-primary"
                      />
                      <div>
                        <div className="font-medium">{list.listName}</div>
                        <div className="text-sm text-base-content/60">{list.siteName}</div>
                      </div>
                    </label>
                  );
                })}
              </div>
            )}
          </div>
        );

      case 'columns':
        return (
          <div className="space-y-4">
            <p className="text-base-content/60">
              Select the columns to display in this view.
            </p>

            {loadingColumns ? (
              <div className="flex items-center justify-center py-8">
                <span className="loading loading-spinner loading-lg text-primary" />
              </div>
            ) : availableColumns.length === 0 ? (
              <div className="alert alert-info">
                <span>Select data sources first to see available columns.</span>
              </div>
            ) : (
              <div className="space-y-4">
                {/* Available columns grouped by source */}
                {sources.map((source) => {
                  const sourceCols = availableColumns.filter(
                    (c) => c.sourceListId === source.listId
                  );
                  if (sourceCols.length === 0) return null;

                  return (
                    <div key={source.listId} className="card bg-base-200">
                      <div className="card-body p-4">
                        <h3 className="font-medium text-sm text-base-content/60 mb-2">
                          {source.listName}
                        </h3>
                        <div className="grid grid-cols-2 gap-2">
                          {sourceCols.map((col) => {
                            const isSelected = columns.some(
                              (c) =>
                                c.internalName === col.name && c.sourceListId === col.sourceListId
                            );
                            return (
                              <label
                                key={`${col.sourceListId}:${col.name}`}
                                className="flex items-center gap-2 cursor-pointer"
                              >
                                <input
                                  type="checkbox"
                                  checked={isSelected}
                                  onChange={() => handleColumnToggle(col)}
                                  className="checkbox checkbox-sm checkbox-primary"
                                />
                                <span className="text-sm">{col.displayName}</span>
                              </label>
                            );
                          })}
                        </div>
                      </div>
                    </div>
                  );
                })}

                {/* Selected columns configuration */}
                {columns.length > 0 && (
                  <div className="card bg-base-200 mt-4">
                    <div className="card-body p-4">
                      <h3 className="font-medium mb-2">Column Configuration</h3>
                      <p className="text-sm text-base-content/60 mb-3">Drag to reorder columns</p>
                      <div className="space-y-2">
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
                            className={`flex items-center gap-3 p-2 rounded-lg bg-base-100 border cursor-move transition-all ${
                              draggedColIndex === index ? 'opacity-50 border-primary' : 'border-transparent hover:border-base-300'
                            }`}
                          >
                            <svg
                              xmlns="http://www.w3.org/2000/svg"
                              fill="none"
                              viewBox="0 0 24 24"
                              strokeWidth={1.5}
                              stroke="currentColor"
                              className="w-4 h-4 text-base-content/40 flex-shrink-0"
                            >
                              <path
                                strokeLinecap="round"
                                strokeLinejoin="round"
                                d="M3.75 6.75h16.5M3.75 12h16.5m-16.5 5.25h16.5"
                              />
                            </svg>
                            <input
                              type="text"
                              value={col.displayName}
                              onChange={(e) =>
                                handleColumnDisplayNameChange(
                                  col.sourceListId,
                                  col.internalName,
                                  e.target.value
                                )
                              }
                              onClick={(e) => e.stopPropagation()}
                              className="input input-sm input-bordered flex-1"
                              placeholder="Display name"
                            />
                            {mode === 'aggregate' && (
                              <select
                                value={col.aggregation || ''}
                                onChange={(e) =>
                                  handleColumnAggregationChange(
                                    col.sourceListId,
                                    col.internalName,
                                    (e.target.value as AggregationType) || undefined
                                  )
                                }
                                onClick={(e) => e.stopPropagation()}
                                className="select select-sm select-bordered"
                              >
                                <option value="">No aggregation</option>
                                {AGGREGATION_OPTIONS.map((opt) => (
                                  <option key={opt.value} value={opt.value}>
                                    {opt.label}
                                  </option>
                                ))}
                              </select>
                            )}
                          </div>
                        ))}
                      </div>
                    </div>
                  </div>
                )}
              </div>
            )}
          </div>
        );

      case 'groupby':
        return (
          <div className="space-y-4">
            <p className="text-base-content/60">
              Select columns to group by. Rows will be grouped by these columns, and aggregations
              will be computed for each group.
            </p>

            {columns.length === 0 ? (
              <div className="alert alert-info">
                <span>Select columns first to configure grouping.</span>
              </div>
            ) : (
              <div className="space-y-2">
                {columns.map((col) => {
                  const isGroupBy = groupBy.includes(col.internalName);
                  return (
                    <label
                      key={`${col.sourceListId}:${col.internalName}`}
                      className={`flex items-center gap-3 p-3 rounded-lg cursor-pointer transition-colors ${
                        isGroupBy ? 'bg-primary/10 border-primary' : 'bg-base-200 hover:bg-base-300'
                      } border`}
                    >
                      <input
                        type="checkbox"
                        checked={isGroupBy}
                        onChange={() => handleGroupByToggle(col.internalName)}
                        className="checkbox checkbox-primary"
                      />
                      <div className="flex-1">
                        <div className="font-medium">{col.displayName}</div>
                        <div className="text-sm text-base-content/60">
                          {isGroupBy ? 'Group by this column' : col.aggregation ? `Aggregate: ${col.aggregation}` : 'No aggregation'}
                        </div>
                      </div>
                    </label>
                  );
                })}
              </div>
            )}

            {groupBy.length === 0 && columns.length > 0 && (
              <div className="alert alert-warning">
                <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={1.5} stroke="currentColor" className="w-5 h-5">
                  <path strokeLinecap="round" strokeLinejoin="round" d="M12 9v3.75m9-.75a9 9 0 1 1-18 0 9 9 0 0 1 18 0Zm-9 3.75h.008v.008H12v-.008Z" />
                </svg>
                <span>Without group by columns, the view will show a single row with aggregated totals.</span>
              </div>
            )}
          </div>
        );

      case 'filters':
        return (
          <div className="space-y-4">
            <p className="text-base-content/60">
              Add filters to narrow down the data (optional).
            </p>

            {filters.length === 0 ? (
              <div className="text-center py-8 text-base-content/60">
                <p>No filters added</p>
              </div>
            ) : (
              <div className="space-y-3">
                {filters.map((filter, index) => (
                  <div key={index} className="flex items-center gap-2">
                    <select
                      value={filter.column}
                      onChange={(e) => handleFilterChange(index, 'column', e.target.value)}
                      className="select select-sm select-bordered flex-1"
                    >
                      {columns.map((col) => (
                        <option key={`${col.sourceListId}:${col.internalName}`} value={col.internalName}>
                          {col.displayName}
                        </option>
                      ))}
                    </select>
                    <select
                      value={filter.operator}
                      onChange={(e) => handleFilterChange(index, 'operator', e.target.value)}
                      className="select select-sm select-bordered"
                    >
                      {FILTER_OPERATORS.map((op) => (
                        <option key={op.value} value={op.value}>
                          {op.label}
                        </option>
                      ))}
                    </select>
                    <input
                      type="text"
                      value={filter.value}
                      onChange={(e) => handleFilterChange(index, 'value', e.target.value)}
                      placeholder="Value"
                      className="input input-sm input-bordered flex-1"
                    />
                    <button
                      onClick={() => handleRemoveFilter(index)}
                      className="btn btn-sm btn-ghost btn-square text-error"
                    >
                      <svg
                        xmlns="http://www.w3.org/2000/svg"
                        fill="none"
                        viewBox="0 0 24 24"
                        strokeWidth={1.5}
                        stroke="currentColor"
                        className="w-4 h-4"
                      >
                        <path
                          strokeLinecap="round"
                          strokeLinejoin="round"
                          d="M6 18 18 6M6 6l12 12"
                        />
                      </svg>
                    </button>
                  </div>
                ))}
              </div>
            )}

            <button
              onClick={handleAddFilter}
              disabled={columns.length === 0}
              className="btn btn-sm btn-outline"
            >
              <svg
                xmlns="http://www.w3.org/2000/svg"
                fill="none"
                viewBox="0 0 24 24"
                strokeWidth={1.5}
                stroke="currentColor"
                className="w-4 h-4"
              >
                <path strokeLinecap="round" strokeLinejoin="round" d="M12 4.5v15m7.5-7.5h-15" />
              </svg>
              Add Filter
            </button>
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
          <div className="space-y-6">
            <p className="text-base-content/60">
              Choose how to sort the data. You can add a primary sort and an optional secondary sort.
            </p>

            {/* Primary Sort */}
            <div className="card bg-base-200">
              <div className="card-body p-4">
                <h3 className="font-medium mb-3">Primary Sort</h3>
                <div className="flex gap-3 items-end">
                  <div className="form-control flex-1">
                    <label className="label">
                      <span className="label-text">Column</span>
                    </label>
                    <select
                      value={sorting[0]?.column || ''}
                      onChange={(e) => setPrimarySortColumn(e.target.value)}
                      className="select select-bordered w-full"
                    >
                      <option value="">No sorting</option>
                      {columns.map((col) => (
                        <option key={`${col.sourceListId}:${col.internalName}`} value={col.internalName}>
                          {col.displayName}
                        </option>
                      ))}
                    </select>
                  </div>
                  {sorting[0] && (
                    <div className="form-control">
                      <label className="label">
                        <span className="label-text">Direction</span>
                      </label>
                      <select
                        value={sorting[0].direction}
                        onChange={(e) => updateSortRule(0, { direction: e.target.value as 'asc' | 'desc' })}
                        className="select select-bordered"
                      >
                        <option value="asc">Ascending (A-Z, 0-9)</option>
                        <option value="desc">Descending (Z-A, 9-0)</option>
                      </select>
                    </div>
                  )}
                </div>
              </div>
            </div>

            {/* Secondary Sort */}
            {sorting.length > 0 && (
              <div className="card bg-base-200">
                <div className="card-body p-4">
                  <h3 className="font-medium mb-3">Secondary Sort (Optional)</h3>
                  <div className="flex gap-3 items-end">
                    <div className="form-control flex-1">
                      <label className="label">
                        <span className="label-text">Column</span>
                      </label>
                      <select
                        value={sorting[1]?.column || ''}
                        onChange={(e) => setSecondarySortColumn(e.target.value)}
                        className="select select-bordered w-full"
                      >
                        <option value="">No secondary sort</option>
                        {columns
                          .filter((col) => col.internalName !== sorting[0]?.column)
                          .map((col) => (
                            <option key={`${col.sourceListId}:${col.internalName}`} value={col.internalName}>
                              {col.displayName}
                            </option>
                          ))}
                      </select>
                    </div>
                    {sorting[1] && (
                      <div className="form-control">
                        <label className="label">
                          <span className="label-text">Direction</span>
                        </label>
                        <select
                          value={sorting[1].direction}
                          onChange={(e) => updateSortRule(1, { direction: e.target.value as 'asc' | 'desc' })}
                          className="select select-bordered"
                        >
                          <option value="asc">Ascending (A-Z, 0-9)</option>
                          <option value="desc">Descending (Z-A, 9-0)</option>
                        </select>
                      </div>
                    )}
                  </div>
                </div>
              </div>
            )}
          </div>
        );
      }
    }
  };

  const isLastStep = currentStep === STEPS[STEPS.length - 1].key;

  return (
    <div className="space-y-6">
      {/* Steps indicator */}
      <ul className="steps steps-horizontal w-full">
        {STEPS.map((step) => (
          <li
            key={step.key}
            className={`step cursor-pointer ${
              STEPS.findIndex((s) => s.key === currentStep) >=
              STEPS.findIndex((s) => s.key === step.key)
                ? 'step-primary'
                : ''
            }`}
            onClick={() => setCurrentStep(step.key)}
          >
            {step.label}
          </li>
        ))}
      </ul>

      {/* Error message */}
      {error && (
        <div className="alert alert-error">
          <svg
            xmlns="http://www.w3.org/2000/svg"
            fill="none"
            viewBox="0 0 24 24"
            strokeWidth={1.5}
            stroke="currentColor"
            className="w-5 h-5"
          >
            <path
              strokeLinecap="round"
              strokeLinejoin="round"
              d="M12 9v3.75m9-.75a9 9 0 1 1-18 0 9 9 0 0 1 18 0Zm-9 3.75h.008v.008H12v-.008Z"
            />
          </svg>
          <span>{error}</span>
        </div>
      )}

      {/* Step content */}
      <div className="card bg-base-200">
        <div className="card-body">{renderStepContent()}</div>
      </div>

      {/* Navigation buttons */}
      <div className="flex items-center justify-between">
        <div>
          {currentStep !== STEPS[0].key && (
            <button onClick={goToPreviousStep} className="btn btn-ghost">
              <svg
                xmlns="http://www.w3.org/2000/svg"
                fill="none"
                viewBox="0 0 24 24"
                strokeWidth={1.5}
                stroke="currentColor"
                className="w-4 h-4"
              >
                <path
                  strokeLinecap="round"
                  strokeLinejoin="round"
                  d="M10.5 19.5 3 12m0 0 7.5-7.5M3 12h18"
                />
              </svg>
              Previous
            </button>
          )}
        </div>

        <div className="flex items-center gap-3">
          <button onClick={onCancel} className="btn btn-ghost">
            Cancel
          </button>
          {isLastStep ? (
            <button onClick={handleSave} disabled={saving} className="btn btn-primary">
              {saving ? (
                <>
                  <span className="loading loading-spinner loading-sm" />
                  Saving...
                </>
              ) : (
                <>
                  <svg
                    xmlns="http://www.w3.org/2000/svg"
                    fill="none"
                    viewBox="0 0 24 24"
                    strokeWidth={1.5}
                    stroke="currentColor"
                    className="w-4 h-4"
                  >
                    <path strokeLinecap="round" strokeLinejoin="round" d="m4.5 12.75 6 6 9-13.5" />
                  </svg>
                  Save View
                </>
              )}
            </button>
          ) : (
            <button
              onClick={goToNextStep}
              disabled={!canProceed(currentStep)}
              className="btn btn-primary"
            >
              Next
              <svg
                xmlns="http://www.w3.org/2000/svg"
                fill="none"
                viewBox="0 0 24 24"
                strokeWidth={1.5}
                stroke="currentColor"
                className="w-4 h-4"
              >
                <path
                  strokeLinecap="round"
                  strokeLinejoin="round"
                  d="M13.5 4.5 21 12m0 0-7.5 7.5M21 12H3"
                />
              </svg>
            </button>
          )}
        </div>
      </div>
    </div>
  );
}

export default ViewEditor;
