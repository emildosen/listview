import { useState, useCallback, useEffect } from 'react';
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
} from '@fluentui/react-components';
import { AddRegular, DeleteRegular } from '@fluentui/react-icons';
import { useMsal } from '@azure/msal-react';
import type { WebPartJoin, WebPartDataSource } from '../../../types/page';
import type { GraphListColumn } from '../../../auth/graphClient';
import { getListColumns } from '../../../auth/graphClient';
import DataSourcePicker from './DataSourcePicker';

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
  joinHeader: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
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
    flexWrap: 'wrap',
    gap: '8px',
    maxHeight: '150px',
    overflowY: 'auto',
  },
  addButton: {
    marginTop: '8px',
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
  deleteButton: {
    marginLeft: 'auto',
  },
});

function generateId(): string {
  return `join-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`;
}

interface JoinBuilderProps {
  joins: WebPartJoin[];
  primaryColumns: GraphListColumn[];
  onChange: (joins: WebPartJoin[]) => void;
}

export default function JoinBuilder({ joins, primaryColumns, onChange }: JoinBuilderProps) {
  const styles = useStyles();
  const { instance, accounts } = useMsal();
  const account = accounts[0];

  // Track columns for each join's target source
  const [targetColumnsMap, setTargetColumnsMap] = useState<Record<string, GraphListColumn[]>>({});
  const [loadingTargets, setLoadingTargets] = useState<Record<string, boolean>>({});

  // Get lookup columns from primary list (these are the ones that can be joined)
  const lookupColumns = primaryColumns.filter((col) => col.lookup);

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

  const handleAddJoin = useCallback(() => {
    const newJoin: WebPartJoin = {
      id: generateId(),
      targetSource: { siteId: '', listId: '', listName: '' },
      sourceColumn: lookupColumns[0]?.name || '',
      targetColumn: 'id',
      joinType: 'left',
      columnsToInclude: [],
    };
    onChange([...joins, newJoin]);
  }, [joins, lookupColumns, onChange]);

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
    (joinId: string, columnName: string, checked: boolean) => {
      onChange(
        joins.map((j) => {
          if (j.id !== joinId) return j;
          const cols = new Set(j.columnsToInclude);
          if (checked) {
            cols.add(columnName);
          } else {
            cols.delete(columnName);
          }
          return { ...j, columnsToInclude: Array.from(cols) };
        })
      );
    },
    [joins, onChange]
  );

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
                  <div className={styles.joinHeader}>
                    <Text>
                      Join {index + 1}: {join.targetSource?.listName || 'Select a list'}
                    </Text>
                    <Button
                      appearance="subtle"
                      size="small"
                      icon={<DeleteRegular />}
                      className={styles.deleteButton}
                      onClick={(e) => {
                        e.stopPropagation();
                        handleRemoveJoin(join.id);
                      }}
                    />
                  </div>
                </AccordionHeader>
                <AccordionPanel>
                  <div className={styles.joinPanel}>
                    {/* Target List Selection */}
                    <DataSourcePicker
                      value={join.targetSource}
                      onChange={(source) => handleJoinChange(join.id, 'targetSource', source)}
                    />

                    {join.targetSource?.listId && (
                      <>
                        {/* Join Configuration */}
                        <div className={styles.fieldRow}>
                          <Field label="Source column" className={styles.fieldHalf}>
                            <Dropdown
                              placeholder="Select lookup column"
                              value={
                                primaryColumns.find((c) => c.name === join.sourceColumn)
                                  ?.displayName || ''
                              }
                              selectedOptions={join.sourceColumn ? [join.sourceColumn] : []}
                              onOptionSelect={(_, data) =>
                                handleJoinChange(join.id, 'sourceColumn', data.optionValue)
                              }
                            >
                              {lookupColumns.map((col) => (
                                <Option key={col.name} value={col.name}>
                                  {col.displayName}
                                </Option>
                              ))}
                            </Dropdown>
                          </Field>

                          <Field label="Target column" className={styles.fieldHalf}>
                            {isLoading ? (
                              <div className={styles.loadingContainer}>
                                <Spinner size="tiny" />
                                <span>Loading...</span>
                              </div>
                            ) : (
                              <Dropdown
                                placeholder="Select column"
                                value={
                                  targetColumns.find((c) => c.name === join.targetColumn)
                                    ?.displayName || ''
                                }
                                selectedOptions={join.targetColumn ? [join.targetColumn] : []}
                                onOptionSelect={(_, data) =>
                                  handleJoinChange(join.id, 'targetColumn', data.optionValue)
                                }
                              >
                                <Option value="id">ID</Option>
                                {targetColumns.map((col) => (
                                  <Option key={col.name} value={col.name}>
                                    {col.displayName}
                                  </Option>
                                ))}
                              </Dropdown>
                            )}
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
                        {!isLoading && targetColumns.length > 0 && (
                          <div className={styles.columnsSection}>
                            <Text className={styles.columnsLabel}>Columns to include</Text>
                            <div className={styles.columnsList}>
                              {targetColumns.map((col) => (
                                <Checkbox
                                  key={col.name}
                                  label={col.displayName}
                                  checked={join.columnsToInclude.includes(col.name)}
                                  onChange={(_, data) =>
                                    handleToggleColumn(join.id, col.name, data.checked as boolean)
                                  }
                                />
                              ))}
                            </div>
                          </div>
                        )}
                      </>
                    )}
                  </div>
                </AccordionPanel>
              </AccordionItem>
            );
          })}
        </Accordion>
      )}

      <Button
        appearance="subtle"
        icon={<AddRegular />}
        onClick={handleAddJoin}
        className={styles.addButton}
        disabled={lookupColumns.length === 0}
      >
        Add Join
      </Button>

      {lookupColumns.length === 0 && (
        <Text style={{ fontSize: tokens.fontSizeBase200, color: tokens.colorNeutralForeground3 }}>
          No lookup columns available. Joins require lookup columns in the primary list.
        </Text>
      )}
    </div>
  );
}
