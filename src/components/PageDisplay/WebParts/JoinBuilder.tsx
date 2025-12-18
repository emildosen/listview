import { useState, useCallback, useEffect, useMemo } from 'react';
import {
  makeStyles,
  tokens,
  Dropdown,
  Option,
  Field,
  Button,
  Text,
  Checkbox,
  Accordion,
  AccordionItem,
  AccordionHeader,
  AccordionPanel,
  Spinner,
  OptionGroup,
} from '@fluentui/react-components';
import { DeleteRegular, ArrowRightRegular, ArrowLeftRegular } from '@fluentui/react-icons';
import { useMsal } from '@azure/msal-react';
import type { WebPartJoin, WebPartDataSource, JoinColumnConfig, JoinColumnAggregation } from '../../../types/page';
import type { GraphListColumn, GraphList } from '../../../auth/graphClient';
import { getListColumns, getSiteLists } from '../../../auth/graphClient';

const useStyles = makeStyles({
  container: {
    display: 'flex',
    flexDirection: 'column',
    gap: '8px',
  },
  joinItem: {
    backgroundColor: tokens.colorNeutralBackground2,
    borderRadius: tokens.borderRadiusMedium,
    overflow: 'hidden',
  },
  joinPanel: {
    display: 'flex',
    flexDirection: 'column',
    gap: '12px',
    padding: '12px',
  },
  fieldRow: {
    display: 'flex',
    gap: '12px',
    flexWrap: 'wrap',
  },
  fieldHalf: {
    flex: 1,
    minWidth: '120px',
  },
  columnsSection: {
    display: 'flex',
    flexDirection: 'column',
    gap: '8px',
  },
  columnsLabel: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground2,
    fontWeight: tokens.fontWeightSemibold,
  },
  columnsList: {
    display: 'flex',
    flexDirection: 'column',
    gap: '4px',
    maxHeight: '200px',
    overflowY: 'auto',
  },
  columnConfigItem: {
    padding: '4px 8px',
    backgroundColor: tokens.colorNeutralBackground3,
    borderRadius: tokens.borderRadiusSmall,
  },
  columnConfigHeader: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
  },
  emptyState: {
    padding: '16px',
    textAlign: 'center',
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase200,
    backgroundColor: tokens.colorNeutralBackground2,
    borderRadius: tokens.borderRadiusMedium,
  },
  loadingContainer: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
    padding: '8px',
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase200,
  },
  directionIcon: {
    display: 'inline-flex',
    alignItems: 'center',
    marginRight: '4px',
  },
  joinOptionText: {
    display: 'flex',
    alignItems: 'center',
    gap: '4px',
  },
});

function generateId(): string {
  return `join-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`;
}

// Represents a joinable relationship
interface JoinableRelation {
  type: 'forward' | 'reverse';
  // For forward: the lookup column in primary list
  // For reverse: the lookup column in the other list
  sourceColumnName: string;
  sourceColumnDisplayName: string;
  // The other list
  targetList: GraphList;
  targetListColumns?: GraphListColumn[];
  // For reverse joins, the lookup column points to our list
  lookupColumnInTarget?: string;
}

interface JoinBuilderProps {
  joins: WebPartJoin[];
  primaryColumns: GraphListColumn[];
  primaryDataSource?: WebPartDataSource;
  onChange: (joins: WebPartJoin[]) => void;
}

export default function JoinBuilder({ joins, primaryColumns, primaryDataSource, onChange }: JoinBuilderProps) {
  const styles = useStyles();
  const { instance, accounts } = useMsal();
  const account = accounts[0];

  // Track columns for each join's target source
  const [targetColumnsMap, setTargetColumnsMap] = useState<Record<string, GraphListColumn[]>>({});
  const [loadingTargets, setLoadingTargets] = useState<Record<string, boolean>>({});

  // State for reverse lookup discovery
  const [siteLists, setSiteLists] = useState<GraphList[]>([]);
  const [siteListColumns, setSiteListColumns] = useState<Record<string, GraphListColumn[]>>({});
  const [loadingRelations, setLoadingRelations] = useState(false);

  // Get lookup columns from primary list (these are the ones that can be joined)
  const lookupColumns = primaryColumns.filter((col) => col.lookup);

  // Load site lists and discover reverse lookups
  useEffect(() => {
    async function discoverRelations() {
      if (!account || !primaryDataSource?.siteId || !primaryDataSource?.listId) {
        return;
      }

      setLoadingRelations(true);
      try {
        // Get all lists on the site
        const lists = await getSiteLists(instance, account, primaryDataSource.siteId);
        setSiteLists(lists);

        // For each list (except current), get columns to check for lookups to current list
        const columnsMap: Record<string, GraphListColumn[]> = {};

        for (const list of lists) {
          // Skip current list
          if (list.id === primaryDataSource.listId) continue;

          try {
            const cols = await getListColumns(instance, account, primaryDataSource.siteId, list.id);
            columnsMap[list.id] = cols;
          } catch (err) {
            console.warn(`Failed to load columns for list ${list.displayName}:`, err);
          }
        }

        setSiteListColumns(columnsMap);
      } catch (err) {
        console.error('Failed to discover relations:', err);
      } finally {
        setLoadingRelations(false);
      }
    }

    discoverRelations();
  }, [instance, account, primaryDataSource?.siteId, primaryDataSource?.listId]);

  // Build the list of joinable relations (forward + reverse)
  const joinableRelations = useMemo((): JoinableRelation[] => {
    const relations: JoinableRelation[] = [];

    // Forward lookups: lookup columns in the primary list
    for (const col of lookupColumns) {
      if (!col.lookup?.listId) continue;

      const targetList = siteLists.find((l) => l.id === col.lookup!.listId);
      if (targetList) {
        relations.push({
          type: 'forward',
          sourceColumnName: col.name,
          sourceColumnDisplayName: col.displayName,
          targetList,
          targetListColumns: siteListColumns[targetList.id],
        });
      }
    }

    // Reverse lookups: other lists that have a lookup to our list
    if (primaryDataSource?.listId) {
      for (const [listId, cols] of Object.entries(siteListColumns)) {
        const list = siteLists.find((l) => l.id === listId);
        if (!list) continue;

        // Find lookup columns that point to our primary list
        const lookupsToUs = cols.filter(
          (col) => col.lookup?.listId === primaryDataSource.listId
        );

        for (const lookupCol of lookupsToUs) {
          relations.push({
            type: 'reverse',
            sourceColumnName: lookupCol.name,
            sourceColumnDisplayName: lookupCol.displayName,
            targetList: list,
            targetListColumns: cols,
            lookupColumnInTarget: lookupCol.name,
          });
        }
      }
    }

    return relations;
  }, [lookupColumns, siteLists, siteListColumns, primaryDataSource?.listId]);

  // Load target columns when join target source changes
  useEffect(() => {
    async function loadTargetColumns() {
      for (const join of joins) {
        const key = `${join.targetSource?.siteId}-${join.targetSource?.listId}`;
        if (
          join.targetSource?.siteId &&
          join.targetSource?.listId &&
          !targetColumnsMap[key] &&
          !loadingTargets[key]
        ) {
          setLoadingTargets((prev) => ({ ...prev, [key]: true }));
          try {
            const cols = await getListColumns(
              instance,
              account,
              join.targetSource.siteId,
              join.targetSource.listId
            );
            setTargetColumnsMap((prev) => ({ ...prev, [key]: cols }));
          } catch (err) {
            console.error('Failed to load target columns:', err);
          } finally {
            setLoadingTargets((prev) => ({ ...prev, [key]: false }));
          }
        }
      }
    }

    if (account) {
      loadTargetColumns();
    }
  }, [joins, instance, account, targetColumnsMap, loadingTargets]);

  const handleAddJoinFromRelation = useCallback((relation: JoinableRelation) => {
    const newJoin: WebPartJoin = {
      id: generateId(),
      targetSource: {
        siteId: primaryDataSource?.siteId || '',
        listId: relation.targetList.id,
        listName: relation.targetList.displayName,
      },
      // For forward: source is the lookup column, target is 'id'
      // For reverse: source is 'id', target is the lookup column in the other list
      sourceColumn: relation.type === 'forward' ? relation.sourceColumnName : 'id',
      targetColumn: relation.type === 'forward' ? 'id' : relation.lookupColumnInTarget || 'id',
      joinType: 'left',
      columnsToInclude: [],
    };
    onChange([...joins, newJoin]);
  }, [joins, primaryDataSource?.siteId, onChange]);

  const handleRemoveJoin = useCallback(
    (id: string) => {
      onChange(joins.filter((j) => j.id !== id));
    },
    [joins, onChange]
  );

  const handleJoinChange = useCallback(
    (id: string, field: keyof WebPartJoin, value: unknown) => {
      onChange(
        joins.map((j) => {
          if (j.id !== id) return j;

          // If target source changed, reset columns to include
          if (field === 'targetSource') {
            return { ...j, targetSource: value as WebPartDataSource, columnsToInclude: [] };
          }

          return { ...j, [field]: value };
        })
      );
    },
    [joins, onChange]
  );

  const handleToggleColumn = useCallback(
    (joinId: string, columnName: string, checked: boolean, columns: GraphListColumn[]) => {
      onChange(
        joins.map((j) => {
          if (j.id !== joinId) return j;

          // Update columnsToInclude (legacy)
          const cols = new Set(j.columnsToInclude);
          // Update columnConfigs
          const configs = [...(j.columnConfigs || [])];

          if (checked) {
            cols.add(columnName);
            // Add config if doesn't exist
            if (!configs.find((c) => c.columnName === columnName)) {
              const col = columns.find((c) => c.name === columnName);
              configs.push({
                columnName,
                displayName: col?.displayName || columnName,
                aggregation: 'first',
              });
            }
          } else {
            cols.delete(columnName);
            // Remove config
            const idx = configs.findIndex((c) => c.columnName === columnName);
            if (idx >= 0) configs.splice(idx, 1);
          }

          return {
            ...j,
            columnsToInclude: Array.from(cols),
            columnConfigs: configs,
          };
        })
      );
    },
    [joins, onChange]
  );

  const handleUpdateColumnConfig = useCallback(
    (joinId: string, columnName: string, field: 'displayName' | 'aggregation', value: string) => {
      onChange(
        joins.map((j) => {
          if (j.id !== joinId) return j;

          const configs = [...(j.columnConfigs || [])];
          const config = configs.find((c) => c.columnName === columnName);
          if (config) {
            if (field === 'displayName') {
              config.displayName = value;
            } else if (field === 'aggregation') {
              config.aggregation = value as JoinColumnAggregation;
            }
          }

          return { ...j, columnConfigs: configs };
        })
      );
    },
    [joins, onChange]
  );

  // Get column config for a specific column
  const getColumnConfig = (join: WebPartJoin, columnName: string): JoinColumnConfig | undefined => {
    return join.columnConfigs?.find((c) => c.columnName === columnName);
  };

  const getTargetColumns = (join: WebPartJoin): GraphListColumn[] => {
    const key = `${join.targetSource?.siteId}-${join.targetSource?.listId}`;
    return targetColumnsMap[key] || [];
  };

  const isTargetLoading = (join: WebPartJoin): boolean => {
    const key = `${join.targetSource?.siteId}-${join.targetSource?.listId}`;
    return loadingTargets[key] || false;
  };

  return (
    <div className={styles.container}>
      {joins.length === 0 ? (
        <div className={styles.emptyState}>
          No joins configured. Click "Add Join" to link data from another list.
        </div>
      ) : (
        <Accordion collapsible>
          {joins.map((join, index) => {
            const targetColumns = getTargetColumns(join);
            const isLoading = isTargetLoading(join);

            return (
              <AccordionItem key={join.id} value={join.id} className={styles.joinItem}>
                <AccordionHeader>
                  <Text>
                    Join {index + 1}: {join.targetSource?.listName || 'Select a list'}
                  </Text>
                </AccordionHeader>
                <AccordionPanel>
                  <div className={styles.joinPanel}>
                    {/* Delete button */}
                    <div style={{ display: 'flex', justifyContent: 'flex-end', marginBottom: '8px' }}>
                      <Button
                        appearance="subtle"
                        size="small"
                        icon={<DeleteRegular />}
                        onClick={() => handleRemoveJoin(join.id)}
                      >
                        Remove
                      </Button>
                    </div>
                    {/* Join relationship info (read-only) */}
                    <div className={styles.fieldRow}>
                      <Field label="Source column" className={styles.fieldHalf}>
                        <Text>
                          {join.sourceColumn === 'id'
                            ? 'ID'
                            : primaryColumns.find((c) => c.name === join.sourceColumn)?.displayName || join.sourceColumn}
                        </Text>
                      </Field>
                      <Field label="Target" className={styles.fieldHalf}>
                        <Text>
                          {join.targetSource?.listName || 'Unknown'} ({join.targetColumn === 'id' ? 'ID' : targetColumns.find((c) => c.name === join.targetColumn)?.displayName || join.targetColumn})
                        </Text>
                      </Field>
                    </div>

                    <div className={styles.fieldRow}>
                      <Field label="Join type" className={styles.fieldHalf}>
                        <Dropdown
                          value={join.joinType === 'inner' ? 'Inner join' : 'Left join'}
                          selectedOptions={[join.joinType]}
                          onOptionSelect={(_, data) =>
                            handleJoinChange(join.id, 'joinType', data.optionValue)
                          }
                        >
                          <Option value="left">Left join (keep all primary records)</Option>
                          <Option value="inner">Inner join (only matching records)</Option>
                        </Dropdown>
                      </Field>
                    </div>

                    {/* Columns to Include */}
                    {isLoading ? (
                      <div className={styles.loadingContainer}>
                        <Spinner size="tiny" />
                        <span>Loading columns...</span>
                      </div>
                    ) : targetColumns.length > 0 ? (
                      <div className={styles.columnsSection}>
                        <Text className={styles.columnsLabel}>Columns to include</Text>
                        <div className={styles.columnsList}>
                          {targetColumns.map((col) => {
                            const isSelected = join.columnsToInclude.includes(col.name);
                            const config = getColumnConfig(join, col.name);

                            return (
                              <div key={col.name} className={isSelected ? styles.columnConfigItem : undefined}>
                                <div className={styles.columnConfigHeader}>
                                  <Checkbox
                                    label={col.displayName}
                                    checked={isSelected}
                                    onChange={(_, data) =>
                                      handleToggleColumn(join.id, col.name, data.checked as boolean, targetColumns)
                                    }
                                  />
                                  {isSelected && config && (
                                    <Dropdown
                                      size="small"
                                      style={{ minWidth: '100px' }}
                                      value={
                                        config.aggregation === 'first' ? 'First' :
                                        config.aggregation === 'count' ? 'Count' :
                                        config.aggregation === 'sum' ? 'Sum' :
                                        config.aggregation === 'avg' ? 'Avg' :
                                        config.aggregation === 'min' ? 'Min' :
                                        config.aggregation === 'max' ? 'Max' : 'First'
                                      }
                                      selectedOptions={[config.aggregation]}
                                      onOptionSelect={(_, data) =>
                                        handleUpdateColumnConfig(join.id, col.name, 'aggregation', data.optionValue as string)
                                      }
                                    >
                                      <Option value="first">First</Option>
                                      <Option value="count">Count</Option>
                                      <Option value="sum">Sum</Option>
                                      <Option value="avg">Avg</Option>
                                      <Option value="min">Min</Option>
                                      <Option value="max">Max</Option>
                                    </Dropdown>
                                  )}
                                </div>
                              </div>
                            );
                          })}
                        </div>
                      </div>
                    ) : null}
                  </div>
                </AccordionPanel>
              </AccordionItem>
            );
          })}
        </Accordion>
      )}

      {/* Add Join Dropdown */}
      {loadingRelations ? (
        <div className={styles.loadingContainer}>
          <Spinner size="tiny" />
          <span>Discovering related lists...</span>
        </div>
      ) : joinableRelations.length > 0 ? (
        <Field label="Add join">
          <Dropdown
            placeholder="Select a relationship..."
            selectedOptions={[]}
            onOptionSelect={(_, data) => {
              const relation = joinableRelations.find(
                (r) => `${r.type}-${r.targetList.id}-${r.sourceColumnName}` === data.optionValue
              );
              if (relation) {
                handleAddJoinFromRelation(relation);
              }
            }}
          >
            {/* Forward lookups group */}
            {joinableRelations.some((r) => r.type === 'forward') && (
              <OptionGroup label="Lookup columns (this list → other)">
                {joinableRelations
                  .filter((r) => r.type === 'forward')
                  .map((relation) => (
                    <Option
                      key={`${relation.type}-${relation.targetList.id}-${relation.sourceColumnName}`}
                      value={`${relation.type}-${relation.targetList.id}-${relation.sourceColumnName}`}
                      text={`${relation.sourceColumnDisplayName} → ${relation.targetList.displayName}`}
                    >
                      <span className={styles.joinOptionText}>
                        <ArrowRightRegular className={styles.directionIcon} />
                        {relation.sourceColumnDisplayName} → {relation.targetList.displayName}
                      </span>
                    </Option>
                  ))}
              </OptionGroup>
            )}
            {/* Reverse lookups group */}
            {joinableRelations.some((r) => r.type === 'reverse') && (
              <OptionGroup label="Related lists (other → this list)">
                {joinableRelations
                  .filter((r) => r.type === 'reverse')
                  .map((relation) => (
                    <Option
                      key={`${relation.type}-${relation.targetList.id}-${relation.sourceColumnName}`}
                      value={`${relation.type}-${relation.targetList.id}-${relation.sourceColumnName}`}
                      text={`${relation.targetList.displayName} (via ${relation.sourceColumnDisplayName})`}
                    >
                      <span className={styles.joinOptionText}>
                        <ArrowLeftRegular className={styles.directionIcon} />
                        {relation.targetList.displayName} (via {relation.sourceColumnDisplayName})
                      </span>
                    </Option>
                  ))}
              </OptionGroup>
            )}
          </Dropdown>
        </Field>
      ) : (
        <Text style={{ fontSize: tokens.fontSizeBase200, color: tokens.colorNeutralForeground3 }}>
          No related lists found. Create lookup columns to enable joins.
        </Text>
      )}
    </div>
  );
}
