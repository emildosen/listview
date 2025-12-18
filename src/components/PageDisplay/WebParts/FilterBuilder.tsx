import { useCallback } from 'react';
import {
  makeStyles,
  tokens,
  Dropdown,
  Option,
  Input,
  Button,
} from '@fluentui/react-components';
import { AddRegular, DeleteRegular } from '@fluentui/react-icons';
import type { WebPartFilter, WebPartFilterOperator } from '../../../types/page';
import type { GraphListColumn } from '../../../auth/graphClient';

const useStyles = makeStyles({
  container: {
    display: 'flex',
    flexDirection: 'column',
    gap: '8px',
  },
  filterRow: {
    display: 'flex',
    flexWrap: 'wrap',
    alignItems: 'center',
    gap: '8px',
    padding: '8px',
    backgroundColor: tokens.colorNeutralBackground2,
    borderRadius: tokens.borderRadiusMedium,
  },
  conjunctionDropdown: {
    minWidth: '70px',
  },
  columnDropdown: {
    minWidth: '120px',
    flex: 1,
  },
  operatorDropdown: {
    minWidth: '100px',
  },
  valueInput: {
    minWidth: '100px',
    flex: 1,
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
});

// Operators available for each column type
const OPERATORS_BY_TYPE: Record<string, { value: WebPartFilterOperator; label: string }[]> = {
  text: [
    { value: 'equals', label: 'Equals' },
    { value: 'notEquals', label: 'Not equals' },
    { value: 'contains', label: 'Contains' },
    { value: 'startsWith', label: 'Starts with' },
    { value: 'isEmpty', label: 'Is empty' },
    { value: 'isNotEmpty', label: 'Is not empty' },
  ],
  number: [
    { value: 'equals', label: 'Equals' },
    { value: 'notEquals', label: 'Not equals' },
    { value: 'greaterThan', label: 'Greater than' },
    { value: 'lessThan', label: 'Less than' },
    { value: 'isEmpty', label: 'Is empty' },
    { value: 'isNotEmpty', label: 'Is not empty' },
  ],
  choice: [
    { value: 'equals', label: 'Equals' },
    { value: 'notEquals', label: 'Not equals' },
    { value: 'isEmpty', label: 'Is empty' },
    { value: 'isNotEmpty', label: 'Is not empty' },
  ],
  boolean: [{ value: 'equals', label: 'Equals' }],
  date: [
    { value: 'equals', label: 'Equals' },
    { value: 'notEquals', label: 'Not equals' },
    { value: 'greaterThan', label: 'After' },
    { value: 'lessThan', label: 'Before' },
    { value: 'isEmpty', label: 'Is empty' },
    { value: 'isNotEmpty', label: 'Is not empty' },
  ],
  lookup: [
    { value: 'equals', label: 'Equals' },
    { value: 'notEquals', label: 'Not equals' },
    { value: 'isEmpty', label: 'Is empty' },
    { value: 'isNotEmpty', label: 'Is not empty' },
  ],
};

function getColumnType(column: GraphListColumn): string {
  if (column.text) return 'text';
  if (column.number) return 'number';
  if (column.boolean) return 'boolean';
  if (column.dateTime) return 'date';
  if (column.lookup) return 'lookup';
  if (column.choice) return 'choice';
  return 'text';
}

function generateId(): string {
  return `filter-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`;
}

interface FilterBuilderProps {
  filters: WebPartFilter[];
  columns: GraphListColumn[];
  onChange: (filters: WebPartFilter[]) => void;
}

export default function FilterBuilder({ filters, columns, onChange }: FilterBuilderProps) {
  const styles = useStyles();

  const handleAddFilter = useCallback(() => {
    const newFilter: WebPartFilter = {
      id: generateId(),
      column: columns[0]?.name || '',
      operator: 'equals',
      value: '',
      conjunction: 'and',
    };
    onChange([...filters, newFilter]);
  }, [filters, columns, onChange]);

  const handleRemoveFilter = useCallback(
    (id: string) => {
      onChange(filters.filter((f) => f.id !== id));
    },
    [filters, onChange]
  );

  const handleFilterChange = useCallback(
    (id: string, field: keyof WebPartFilter, value: unknown) => {
      onChange(
        filters.map((f) => {
          if (f.id !== id) return f;

          // If column changed, reset operator to a valid one for the new column type
          if (field === 'column') {
            const column = columns.find((c) => c.name === value);
            const colType = column ? getColumnType(column) : 'text';
            const validOperators = OPERATORS_BY_TYPE[colType] || OPERATORS_BY_TYPE.text;
            const currentOperatorValid = validOperators.some((op) => op.value === f.operator);

            return {
              ...f,
              column: value as string,
              operator: currentOperatorValid ? f.operator : validOperators[0].value,
              value: '',
            };
          }

          return { ...f, [field]: value };
        })
      );
    },
    [filters, columns, onChange]
  );

  const getOperatorsForColumn = (columnName: string) => {
    const column = columns.find((c) => c.name === columnName);
    const colType = column ? getColumnType(column) : 'text';
    return OPERATORS_BY_TYPE[colType] || OPERATORS_BY_TYPE.text;
  };

  const getChoicesForColumn = (columnName: string): string[] => {
    const column = columns.find((c) => c.name === columnName);
    if (column?.choice?.choices) {
      return column.choice.choices;
    }
    if (column?.boolean) {
      return ['Yes', 'No'];
    }
    return [];
  };

  const isValueRequired = (operator: WebPartFilterOperator): boolean => {
    return !['isEmpty', 'isNotEmpty'].includes(operator);
  };

  return (
    <div className={styles.container}>
      {filters.length === 0 ? (
        <div className={styles.emptyState}>
          No filters configured. Click "Add Filter" to create one.
        </div>
      ) : (
        filters.map((filter, index) => {
          const operators = getOperatorsForColumn(filter.column);
          const choices = getChoicesForColumn(filter.column);
          const column = columns.find((c) => c.name === filter.column);
          const isChoiceOrBoolean = column?.choice || column?.boolean;

          return (
            <div key={filter.id} className={styles.filterRow}>
              {/* Conjunction (for 2nd filter onwards) */}
              {index > 0 && (
                <Dropdown
                  className={styles.conjunctionDropdown}
                  size="small"
                  value={filter.conjunction.toUpperCase()}
                  selectedOptions={[filter.conjunction]}
                  onOptionSelect={(_, data) =>
                    handleFilterChange(filter.id, 'conjunction', data.optionValue)
                  }
                >
                  <Option value="and">AND</Option>
                  <Option value="or">OR</Option>
                </Dropdown>
              )}

              {/* Column */}
              <Dropdown
                className={styles.columnDropdown}
                size="small"
                placeholder="Select column"
                value={columns.find((c) => c.name === filter.column)?.displayName || ''}
                selectedOptions={filter.column ? [filter.column] : []}
                onOptionSelect={(_, data) =>
                  handleFilterChange(filter.id, 'column', data.optionValue)
                }
              >
                {columns.map((col) => (
                  <Option key={col.name} value={col.name}>
                    {col.displayName}
                  </Option>
                ))}
              </Dropdown>

              {/* Operator */}
              <Dropdown
                className={styles.operatorDropdown}
                size="small"
                value={operators.find((o) => o.value === filter.operator)?.label || ''}
                selectedOptions={[filter.operator]}
                onOptionSelect={(_, data) =>
                  handleFilterChange(filter.id, 'operator', data.optionValue)
                }
              >
                {operators.map((op) => (
                  <Option key={op.value} value={op.value}>
                    {op.label}
                  </Option>
                ))}
              </Dropdown>

              {/* Value */}
              {isValueRequired(filter.operator) && (
                <>
                  {isChoiceOrBoolean && choices.length > 0 ? (
                    <Dropdown
                      className={styles.valueInput}
                      size="small"
                      placeholder="Select value"
                      value={String(filter.value)}
                      selectedOptions={filter.value ? [String(filter.value)] : []}
                      onOptionSelect={(_, data) =>
                        handleFilterChange(filter.id, 'value', data.optionValue)
                      }
                    >
                      {choices.map((choice) => (
                        <Option key={choice} value={choice}>
                          {choice}
                        </Option>
                      ))}
                    </Dropdown>
                  ) : (
                    <Input
                      className={styles.valueInput}
                      size="small"
                      placeholder="Enter value"
                      value={String(filter.value)}
                      onChange={(_, data) =>
                        handleFilterChange(filter.id, 'value', data.value)
                      }
                    />
                  )}
                </>
              )}

              {/* Delete button */}
              <Button
                appearance="subtle"
                size="small"
                icon={<DeleteRegular />}
                onClick={() => handleRemoveFilter(filter.id)}
              />
            </div>
          );
        })
      )}

      <Button
        appearance="subtle"
        icon={<AddRegular />}
        onClick={handleAddFilter}
        className={styles.addButton}
        disabled={columns.length === 0}
      >
        Add Filter
      </Button>
    </div>
  );
}
