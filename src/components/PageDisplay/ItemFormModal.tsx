import { useState, useMemo } from 'react';
import {
  makeStyles,
  Dialog,
  DialogSurface,
  DialogBody,
  DialogTitle,
  DialogActions,
  Button,
  Input,
  Textarea,
  Dropdown,
  Option,
  Checkbox,
  Field,
  Spinner,
  MessageBar,
  MessageBarBody,
} from '@fluentui/react-components';
import type { GraphListColumn } from '../../auth/graphClient';

interface ItemFormModalProps {
  mode: 'create' | 'edit';
  columns: GraphListColumn[];
  initialValues: Record<string, unknown>;
  saving: boolean;
  onSave: (fields: Record<string, unknown>) => Promise<void>;
  onClose: () => void;
}

const useStyles = makeStyles({
  dialogSurface: {
    maxWidth: '500px',
  },
  formFields: {
    display: 'flex',
    flexDirection: 'column',
    gap: '12px',
  },
});

function ItemFormModal({
  mode,
  columns,
  initialValues,
  saving,
  onSave,
  onClose,
}: ItemFormModalProps) {
  const styles = useStyles();

  // Compute initial values once when the component mounts
  const computedInitialValues = useMemo(() => {
    const initial: Record<string, unknown> = {};
    columns.forEach((col) => {
      if (initialValues[col.name] !== undefined) {
        initial[col.name] = initialValues[col.name];
      } else if (mode === 'create' && col.defaultValue?.value) {
        // Use SharePoint default value for new items
        const defaultVal = col.defaultValue.value;

        if (col.boolean) {
          initial[col.name] = defaultVal === '1' || defaultVal.toLowerCase() === 'true';
        } else if (col.number) {
          initial[col.name] = Number(defaultVal) || '';
        } else if (col.dateTime) {
          // Handle SharePoint date formulas and formats
          if (defaultVal === '[today]' || defaultVal === '[Today]') {
            initial[col.name] = col.dateTime.format === 'dateOnly'
              ? formatDateForInput(new Date())
              : formatDateTimeForInput(new Date());
          } else {
            initial[col.name] = col.dateTime.format === 'dateOnly'
              ? formatDateForInput(defaultVal)
              : formatDateTimeForInput(defaultVal);
          }
        } else if (col.choice) {
          // Choice default - use as-is if it's a valid choice
          if (col.choice.choices.includes(defaultVal)) {
            initial[col.name] = defaultVal;
          } else {
            initial[col.name] = '';
          }
        } else {
          initial[col.name] = defaultVal;
        }
      } else {
        initial[col.name] = '';
      }
    });
    return initial;
  }, [columns, initialValues, mode]);

  const [values, setValues] = useState<Record<string, unknown>>(computedInitialValues);
  const [error, setError] = useState<string | null>(null);

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    setError(null);

    try {
      // Filter out empty values and format properly
      const submitValues: Record<string, unknown> = {};
      columns.forEach((col) => {
        const value = values[col.name];
        if (value !== undefined && value !== '') {
          submitValues[col.name] = value;
        }
      });

      await onSave(submitValues);
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Failed to save');
    }
  };

  const handleChange = (columnName: string, value: unknown) => {
    setValues((prev) => ({ ...prev, [columnName]: value }));
  };

  const renderField = (column: GraphListColumn) => {
    const value = values[column.name];

    // Skip read-only columns in create mode
    if (mode === 'create' && column.readOnly) {
      return null;
    }

    // Render based on column type
    if (column.choice?.choices) {
      // Choice column
      return (
        <Field key={column.id} label={column.displayName}>
          <Dropdown
            value={String(value || '')}
            selectedOptions={value ? [String(value)] : []}
            onOptionSelect={(_e, data) => handleChange(column.name, data.optionValue)}
            disabled={column.readOnly}
            size="small"
          >
            <Option value="">Select...</Option>
            {column.choice.choices.map((choice) => (
              <Option key={choice} value={choice}>
                {choice}
              </Option>
            ))}
          </Dropdown>
        </Field>
      );
    }

    if (column.lookup) {
      // Lookup column - for now just show as text (would need to load options)
      return (
        <Field key={column.id} label={column.displayName}>
          <Input
            value={typeof value === 'object' && value && 'LookupValue' in value
              ? (value as { LookupValue: string }).LookupValue
              : String(value || '')}
            disabled
            placeholder="Lookup field (read-only)"
            size="small"
          />
        </Field>
      );
    }

    // Boolean column
    if (column.boolean) {
      return (
        <Checkbox
          key={column.id}
          checked={Boolean(value)}
          onChange={(_e, data) => handleChange(column.name, data.checked)}
          disabled={column.readOnly}
          label={column.displayName}
        />
      );
    }

    // Number column
    if (column.number) {
      return (
        <Field key={column.id} label={column.displayName}>
          <Input
            type="number"
            value={value !== undefined && value !== null ? String(value) : ''}
            onChange={(_e, data) => handleChange(column.name, data.value ? Number(data.value) : '')}
            disabled={column.readOnly}
            size="small"
          />
        </Field>
      );
    }

    // DateTime column
    if (column.dateTime) {
      const isDateOnly = column.dateTime.format === 'dateOnly';
      return (
        <Field key={column.id} label={column.displayName}>
          <Input
            type={isDateOnly ? 'date' : 'datetime-local'}
            value={isDateOnly ? formatDateForInput(value) : formatDateTimeForInput(value)}
            onChange={(_e, data) => handleChange(column.name, data.value)}
            disabled={column.readOnly}
            size="small"
          />
        </Field>
      );
    }

    // Text column (check for multiline)
    if (column.text?.allowMultipleLines) {
      return (
        <Field key={column.id} label={column.displayName}>
          <Textarea
            value={String(value || '')}
            onChange={(_e, data) => handleChange(column.name, data.value)}
            disabled={column.readOnly}
            rows={3}
            size="small"
          />
        </Field>
      );
    }

    // Default: single-line text
    return (
      <Field key={column.id} label={column.displayName}>
        <Input
          value={String(value || '')}
          onChange={(_e, data) => handleChange(column.name, data.value)}
          disabled={column.readOnly}
          size="small"
        />
      </Field>
    );
  };

  return (
    <Dialog open onOpenChange={(_event, data) => { if (!data.open) onClose(); }}>
      <DialogSurface className={styles.dialogSurface}>
        <form onSubmit={handleSubmit}>
          <DialogTitle>
            {mode === 'create' ? 'Add New Item' : 'Edit Item'}
          </DialogTitle>
          <DialogBody>
            {error && (
              <MessageBar intent="error" style={{ marginBottom: '16px' }}>
                <MessageBarBody>{error}</MessageBarBody>
              </MessageBar>
            )}

            <div className={styles.formFields}>
              {columns.map(renderField)}
            </div>
          </DialogBody>
          <DialogActions>
            <Button appearance="secondary" onClick={onClose} disabled={saving}>
              Cancel
            </Button>
            <Button
              appearance="primary"
              type="submit"
              disabled={saving}
              icon={saving ? <Spinner size="tiny" /> : undefined}
            >
              {saving ? 'Saving...' : 'Save'}
            </Button>
          </DialogActions>
        </form>
      </DialogSurface>
    </Dialog>
  );
}

function formatLocalDate(d: Date): string {
  const year = d.getFullYear();
  const month = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

function formatLocalDateTime(d: Date): string {
  const hours = String(d.getHours()).padStart(2, '0');
  const minutes = String(d.getMinutes()).padStart(2, '0');
  return `${formatLocalDate(d)}T${hours}:${minutes}`;
}

function formatDateForInput(value: unknown): string {
  if (!value) return '';
  if (typeof value === 'string') {
    // If already in YYYY-MM-DD format, use as-is
    const match = value.match(/^(\d{4}-\d{2}-\d{2})/);
    if (match) return match[1];
    // Try to parse as date and format in local time
    const date = new Date(value);
    if (!isNaN(date.getTime())) {
      return formatLocalDate(date);
    }
  }
  if (value instanceof Date) {
    return formatLocalDate(value);
  }
  return '';
}

function formatDateTimeForInput(value: unknown): string {
  if (!value) return '';
  if (typeof value === 'string') {
    const date = new Date(value);
    if (!isNaN(date.getTime())) {
      return formatLocalDateTime(date);
    }
  }
  if (value instanceof Date) {
    return formatLocalDateTime(value);
  }
  return '';
}

export default ItemFormModal;
