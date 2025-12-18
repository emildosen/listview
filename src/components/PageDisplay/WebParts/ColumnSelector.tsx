import { useCallback } from 'react';
import {
  makeStyles,
  tokens,
  Checkbox,
  Input,
  Text,
  Button,
  Spinner,
} from '@fluentui/react-components';
import {
  ReOrderDotsVerticalRegular,
  ArrowUpRegular,
  ArrowDownRegular,
} from '@fluentui/react-icons';
import type { WebPartDisplayColumn } from '../../../types/page';
import type { GraphListColumn } from '../../../auth/graphClient';

const useStyles = makeStyles({
  container: {
    display: 'flex',
    flexDirection: 'column',
    gap: '8px',
  },
  columnList: {
    display: 'flex',
    flexDirection: 'column',
    gap: '4px',
    maxHeight: '300px',
    overflowY: 'auto',
  },
  columnItem: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
    padding: '8px',
    backgroundColor: tokens.colorNeutralBackground2,
    borderRadius: tokens.borderRadiusMedium,
  },
  selectedColumnItem: {
    backgroundColor: tokens.colorNeutralBackground3,
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
  selectedSection: {
    marginTop: '12px',
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

  const handleToggleColumn = useCallback(
    (column: GraphListColumn, checked: boolean) => {
      if (checked) {
        // Add column
        onChange([
          ...selectedColumns,
          {
            internalName: column.name,
            displayName: column.displayName,
          },
        ]);
      } else {
        // Remove column
        onChange(selectedColumns.filter((c) => c.internalName !== column.name));
      }
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

  const isSelected = (columnName: string) =>
    selectedColumns.some((c) => c.internalName === columnName);

  return (
    <div className={styles.container}>
      {/* Available columns */}
      <Text className={styles.sectionLabel}>Available Columns</Text>
      <div className={styles.columnList}>
        {availableColumns.map((column) => (
          <div
            key={column.name}
            className={`${styles.columnItem} ${
              isSelected(column.name) ? styles.selectedColumnItem : ''
            }`}
          >
            <Checkbox
              checked={isSelected(column.name)}
              onChange={(_, data) => handleToggleColumn(column, data.checked as boolean)}
            />
            <div className={styles.columnInfo}>
              <Text className={styles.columnName}>{column.displayName}</Text>
              <Text className={styles.columnType}>{getColumnTypeLabel(column)}</Text>
            </div>
          </div>
        ))}
      </div>

      {/* Selected columns with reordering */}
      {selectedColumns.length > 0 && (
        <div className={styles.selectedSection}>
          <Text className={styles.sectionLabel}>
            Selected Columns ({selectedColumns.length})
          </Text>
          <div className={styles.columnList}>
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
                </div>
              );
            })}
          </div>
        </div>
      )}
    </div>
  );
}
