import { useCallback, useState } from 'react';
import {
  makeStyles,
  tokens,
  Input,
  Text,
  Button,
  Spinner,
  Dropdown,
  Option,
} from '@fluentui/react-components';
import {
  ReOrderDotsVerticalRegular,
  ArrowUpRegular,
  ArrowDownRegular,
  DeleteRegular,
} from '@fluentui/react-icons';
import type { WebPartDisplayColumn } from '../../../types/page';
import type { GraphListColumn } from '../../../auth/graphClient';

const useStyles = makeStyles({
  container: {
    display: 'flex',
    flexDirection: 'column',
    gap: '8px',
  },
  addColumnDropdown: {
    width: '100%',
  },
  selectedColumnList: {
    display: 'flex',
    flexDirection: 'column',
    gap: '4px',
  },
  columnItem: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
    padding: '8px',
    backgroundColor: tokens.colorNeutralBackground2,
    borderRadius: tokens.borderRadiusMedium,
  },
  dragHandle: {
    color: tokens.colorNeutralForeground3,
    cursor: 'grab',
    flexShrink: 0,
  },
  columnInfo: {
    flex: 1,
    minWidth: 0,
  },
  columnName: {
    display: 'block',
    fontWeight: tokens.fontWeightSemibold,
    fontSize: tokens.fontSizeBase200,
  },
  columnType: {
    display: 'block',
    fontSize: tokens.fontSizeBase100,
    color: tokens.colorNeutralForeground3,
  },
  displayNameInput: {
    flex: 1,
    maxWidth: '150px',
  },
  reorderButtons: {
    display: 'flex',
    flexDirection: 'column',
    gap: '2px',
  },
  loadingContainer: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
    padding: '16px',
    color: tokens.colorNeutralForeground3,
  },
  emptyState: {
    padding: '16px',
    textAlign: 'center',
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase200,
  },
  sectionLabel: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground2,
    marginBottom: '8px',
    fontWeight: tokens.fontWeightSemibold,
  },
});

function getColumnTypeLabel(column: GraphListColumn): string {
  if (column.text) return 'Text';
  if (column.number) return 'Number';
  if (column.boolean) return 'Yes/No';
  if (column.dateTime) return 'Date';
  if (column.lookup) return 'Lookup';
  if (column.choice) return 'Choice';
  if (column.hyperlinkOrPicture) return 'URL';
  return 'Unknown';
}

interface ColumnSelectorProps {
  availableColumns: GraphListColumn[];
  selectedColumns: WebPartDisplayColumn[];
  onChange: (columns: WebPartDisplayColumn[]) => void;
  loading?: boolean;
}

export default function ColumnSelector({
  availableColumns,
  selectedColumns,
  onChange,
  loading,
}: ColumnSelectorProps) {
  const styles = useStyles();
  const [dropdownValue, setDropdownValue] = useState<string>('');

  const handleAddColumn = useCallback(
    (columnName: string) => {
      const column = availableColumns.find((c) => c.name === columnName);
      if (column) {
        onChange([
          ...selectedColumns,
          {
            internalName: column.name,
            displayName: column.displayName,
          },
        ]);
        setDropdownValue('');
      }
    },
    [availableColumns, selectedColumns, onChange]
  );

  const handleRemoveColumn = useCallback(
    (internalName: string) => {
      onChange(selectedColumns.filter((c) => c.internalName !== internalName));
    },
    [selectedColumns, onChange]
  );

  const handleDisplayNameChange = useCallback(
    (internalName: string, displayName: string) => {
      onChange(
        selectedColumns.map((c) =>
          c.internalName === internalName ? { ...c, displayName } : c
        )
      );
    },
    [selectedColumns, onChange]
  );

  const handleMoveUp = useCallback(
    (index: number) => {
      if (index === 0) return;
      const newColumns = [...selectedColumns];
      [newColumns[index - 1], newColumns[index]] = [newColumns[index], newColumns[index - 1]];
      onChange(newColumns);
    },
    [selectedColumns, onChange]
  );

  const handleMoveDown = useCallback(
    (index: number) => {
      if (index === selectedColumns.length - 1) return;
      const newColumns = [...selectedColumns];
      [newColumns[index], newColumns[index + 1]] = [newColumns[index + 1], newColumns[index]];
      onChange(newColumns);
    },
    [selectedColumns, onChange]
  );

  if (loading) {
    return (
      <div className={styles.loadingContainer}>
        <Spinner size="tiny" />
        <span>Loading columns...</span>
      </div>
    );
  }

  if (availableColumns.length === 0) {
    return <div className={styles.emptyState}>No columns available</div>;
  }

  // Filter out already selected columns for the dropdown
  const unselectedColumns = availableColumns.filter(
    (col) => !selectedColumns.some((c) => c.internalName === col.name)
  );

  return (
    <div className={styles.container}>
      {/* Selected columns with reordering */}
      {selectedColumns.length > 0 && (
        <>
          <Text className={styles.sectionLabel}>
            Selected Columns ({selectedColumns.length})
          </Text>
          <div className={styles.selectedColumnList}>
            {selectedColumns.map((column, index) => {
              const sourceColumn = availableColumns.find(
                (c) => c.name === column.internalName
              );
              return (
                <div key={column.internalName} className={styles.columnItem}>
                  <ReOrderDotsVerticalRegular className={styles.dragHandle} />
                  <div className={styles.columnInfo}>
                    <Text className={styles.columnName}>
                      {sourceColumn?.displayName || column.internalName}
                    </Text>
                  </div>
                  <Input
                    className={styles.displayNameInput}
                    size="small"
                    value={column.displayName}
                    onChange={(_, data) =>
                      handleDisplayNameChange(column.internalName, data.value)
                    }
                    placeholder="Display name"
                  />
                  <div className={styles.reorderButtons}>
                    <Button
                      appearance="subtle"
                      size="small"
                      icon={<ArrowUpRegular />}
                      onClick={() => handleMoveUp(index)}
                      disabled={index === 0}
                    />
                    <Button
                      appearance="subtle"
                      size="small"
                      icon={<ArrowDownRegular />}
                      onClick={() => handleMoveDown(index)}
                      disabled={index === selectedColumns.length - 1}
                    />
                  </div>
                  <Button
                    appearance="subtle"
                    size="small"
                    icon={<DeleteRegular />}
                    onClick={() => handleRemoveColumn(column.internalName)}
                  />
                </div>
              );
            })}
          </div>
        </>
      )}

      {/* Add column dropdown */}
      <Dropdown
        className={styles.addColumnDropdown}
        placeholder="Add a column..."
        value={dropdownValue}
        selectedOptions={[]}
        onOptionSelect={(_, data) => {
          if (data.optionValue) {
            handleAddColumn(data.optionValue);
          }
        }}
        disabled={unselectedColumns.length === 0}
      >
        {unselectedColumns.map((column) => (
          <Option
            key={column.name}
            value={column.name}
            text={`${column.displayName} (${getColumnTypeLabel(column)})`}
          >
            {column.displayName} ({getColumnTypeLabel(column)})
          </Option>
        ))}
      </Dropdown>
    </div>
  );
}
