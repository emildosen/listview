import { useState, useMemo, useEffect, useRef } from 'react';
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
import { useListFormConfig } from '../../hooks/useListFormConfig';
import type { FormFieldConfig } from '../../auth/graphClient';
import type { LookupOption } from '../../contexts/FormConfigContext';

// System columns that should never appear in forms
const SYSTEM_COLUMNS = new Set([
  'ContentType',
  'Attachments',
  'ID',
  'Created',
  'Modified',
  'Author',
  'Editor',
  '_UIVersionString',
  '_ModerationStatus',
  '_ModerationComments',
  'Edit',
  'LinkTitleNoMenu',
  'LinkTitle',
  'DocIcon',
  'ItemChildCount',
  'FolderChildCount',
  'AppAuthor',
  'AppEditor',
]);

function isSystemColumn(name: string): boolean {
  return SYSTEM_COLUMNS.has(name) || name.startsWith('_');
}

interface ItemFormModalProps {
  mode: 'create' | 'edit';
  siteId: string;
  listId: string;
  initialValues: Record<string, unknown>;
  saving: boolean;
  onSave: (fields: Record<string, unknown>) => Promise<void>;
  onClose: () => void;
  /** Lookup field that is pre-filled by parent (hidden from form and excluded from save) */
  prefillLookupField?: string;
}

const useStyles = makeStyles({
  loadingContainer: {
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    justifyContent: 'center',
    padding: '40px 20px',
    gap: '12px',
  },
  dialogSurface: {
    maxWidth: '500px',
  },
  dialogBody: {
    display: 'block',
    paddingTop: '16px',
    paddingBottom: '24px',
  },
  formFields: {
    display: 'flex',
    flexDirection: 'column',
    gap: '12px',
  },
  dateInput: {
    width: '180px',
  },
  dropdown: {
    minWidth: '80px',
  },
});

function ItemFormModal({
  mode,
  siteId,
  listId,
  initialValues,
  saving,
  onSave,
  onClose,
  prefillLookupField,
}: ItemFormModalProps) {
  const styles = useStyles();
  const { fields, loading, error: configError, getLookupOptions } = useListFormConfig(siteId, listId);

  // Track user changes separately - merged with computed defaults below
  const [userChanges, setUserChanges] = useState<Record<string, unknown>>({});
  const [error, setError] = useState<string | null>(null);

  // Lookup options state - keyed by field name
  const [lookupOptions, setLookupOptions] = useState<Record<string, LookupOption[]>>({});
  const [lookupLoading, setLookupLoading] = useState<Record<string, boolean>>({});

  // Track which lookups we've started loading to prevent re-fetching
  const loadedLookupsRef = useRef<Set<string>>(new Set());

  // Get visible (non-hidden, non-system) fields for the form
  // Also hide the pre-filled lookup field (it's set by the parent)
  const visibleFields = useMemo(() => {
    return fields.filter((f) =>
      !f.hidden &&
      !isSystemColumn(f.name) &&
      f.name !== prefillLookupField
    );
  }, [fields, prefillLookupField]);

  // Load lookup options for all lookup fields
  useEffect(() => {
    const lookupFields = visibleFields.filter((f) => f.lookup && !f.readOnly);
    if (lookupFields.length === 0) return;

    lookupFields.forEach(async (field) => {
      if (!field.lookup) return;

      // Use ref to track loaded lookups - prevents re-running on state changes
      if (loadedLookupsRef.current.has(field.name)) return;
      loadedLookupsRef.current.add(field.name);

      setLookupLoading((prev) => ({ ...prev, [field.name]: true }));

      try {
        // The lookup.listId is the target list's GUID
        // We need to use the same siteId since lookups are within the same site
        const options = await getLookupOptions(siteId, field.lookup.listId, field.lookup.columnName);
        setLookupOptions((prev) => ({ ...prev, [field.name]: options }));
      } catch (err) {
        console.error(`Failed to load lookup options for ${field.name}:`, err);
        setLookupOptions((prev) => ({ ...prev, [field.name]: [] }));
      } finally {
        setLookupLoading((prev) => ({ ...prev, [field.name]: false }));
      }
    });
  }, [visibleFields, siteId, getLookupOptions]);

  // Compute default values from fields and initialValues (no effect needed)
  const defaultValues = useMemo(() => {
    const initial: Record<string, unknown> = {};
    visibleFields.forEach((field) => {
      if (field.lookup) {
        // For lookups, extract the ID(s) from the initial value
        // SharePoint Graph API returns: { LookupId, LookupValue } or array of these
        // Also check FieldNameLookupId for the raw ID
        const lookupValue = initialValues[field.name];
        const lookupIdValue = initialValues[`${field.name}LookupId`];

        // Helper to extract ID from a lookup object (handles both cases)
        const extractId = (v: unknown): number | null => {
          if (typeof v === 'number') return v;
          if (typeof v === 'object' && v !== null) {
            // Check both LookupId and lookupId (Graph API uses LookupId)
            if ('LookupId' in v) return (v as { LookupId: number }).LookupId;
            if ('lookupId' in v) return (v as { lookupId: number }).lookupId;
          }
          return null;
        };

        // Helper to parse ID from various formats
        const parseId = (v: unknown): number | null => {
          if (typeof v === 'number') return v;
          if (typeof v === 'string' && v) {
            const parsed = parseInt(v, 10);
            return isNaN(parsed) ? null : parsed;
          }
          return null;
        };

        if (field.lookup.allowMultipleValues) {
          // Multi-select: extract array of IDs
          if (Array.isArray(lookupValue)) {
            initial[field.name] = lookupValue
              .map(extractId)
              .filter((id): id is number => id !== null);
          } else if (Array.isArray(lookupIdValue)) {
            initial[field.name] = lookupIdValue.map(parseId).filter((id): id is number => id !== null);
          } else {
            initial[field.name] = [];
          }
        } else {
          // Single select: try multiple sources for the ID
          // Priority: lookupValue object > FieldNameLookupId field
          const extractedId = extractId(lookupValue);
          const parsedLookupId = parseId(lookupIdValue);

          if (extractedId !== null) {
            initial[field.name] = extractedId;
          } else if (parsedLookupId !== null) {
            initial[field.name] = parsedLookupId;
          } else {
            initial[field.name] = null;
          }
        }
      } else if (field.choice?.allowMultipleValues) {
        // Multi-select choice: extract array from { results: [...] } format if present
        const choiceValue = initialValues[field.name];
        if (Array.isArray(choiceValue)) {
          initial[field.name] = choiceValue;
        } else if (typeof choiceValue === 'object' && choiceValue !== null && 'results' in choiceValue) {
          initial[field.name] = (choiceValue as { results: string[] }).results || [];
        } else if (choiceValue) {
          initial[field.name] = [String(choiceValue)];
        } else {
          initial[field.name] = [];
        }
      } else if (initialValues[field.name] !== undefined) {
        initial[field.name] = initialValues[field.name];
      } else if (mode === 'create' && field.defaultValue?.value) {
        // Use SharePoint default value for new items
        initial[field.name] = parseDefaultValue(field);
      } else {
        initial[field.name] = getEmptyValue(field);
      }
    });
    return initial;
  }, [visibleFields, initialValues, mode]);

  // Combine defaults with user changes - user changes override defaults
  const values = useMemo(
    () => ({ ...defaultValues, ...userChanges }),
    [defaultValues, userChanges]
  );

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    setError(null);

    try {
      // Only submit visible, non-readonly fields
      // Hidden columns are NOT included - SharePoint preserves them on edit,
      // and uses defaults on create
      // Pre-filled lookups are hidden from the form and handled by the parent.
      const submitValues: Record<string, unknown> = {};
      visibleFields.forEach((field) => {
        if (field.readOnly) return;

        const value = values[field.name];

        if (field.lookup) {
          // Lookup fields: submit as FieldNameId with the numeric ID(s)
          if (field.lookup.allowMultipleValues) {
            // Multi-select: always submit array (empty array clears all selections)
            submitValues[`${field.name}Id`] = Array.isArray(value) ? value : [];
          } else {
            // Single select: always submit (null clears the lookup)
            if (value !== undefined && value !== '' && value !== null) {
              submitValues[`${field.name}Id`] = typeof value === 'number' ? value : parseInt(String(value), 10);
            } else {
              submitValues[`${field.name}Id`] = null;
            }
          }
        } else if (field.choice?.allowMultipleValues) {
          // Multi-select choice: PnPjs expects a plain array
          submitValues[field.name] = Array.isArray(value) ? value : [];
        } else {
          // Regular fields: submit as-is
          if (value !== undefined && value !== '') {
            submitValues[field.name] = value;
          }
        }
      });

      await onSave(submitValues);
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Failed to save');
    }
  };

  const handleChange = (fieldName: string, value: unknown) => {
    setUserChanges((prev) => ({ ...prev, [fieldName]: value }));
  };

  // Calculate dropdown width based on longest option
  const getDropdownWidth = (choices: string[]): number => {
    const allOptions = ['Select...', ...choices];
    const longestOption = allOptions.reduce(
      (a, b) => (a.length > b.length ? a : b),
      ''
    );
    // Roughly 8px per character + 50px for padding/icon
    return Math.max(80, longestOption.length * 8 + 50);
  };

  const renderField = (field: FormFieldConfig) => {
    const value = values[field.name];

    // Skip read-only columns in create mode
    if (mode === 'create' && field.readOnly) {
      return null;
    }

    // Render based on field type
    if (field.choice?.choices) {
      // Choice column
      const dropdownWidth = getDropdownWidth(field.choice.choices);
      const isMultiSelect = field.choice.allowMultipleValues;

      // Normalize value for multi-select
      const selectedOptions: string[] = isMultiSelect
        ? (Array.isArray(value) ? value.map(String) : (value ? [String(value)] : []))
        : (value ? [String(value)] : []);

      const displayValue = isMultiSelect
        ? selectedOptions.join(', ')
        : String(value || '');

      return (
        <Field key={field.id} label={field.displayName} required={field.required}>
          <Dropdown
            value={displayValue}
            selectedOptions={selectedOptions}
            multiselect={isMultiSelect}
            onOptionSelect={(_e, data) => {
              if (isMultiSelect) {
                // Multi-select: use the full selected options array
                handleChange(field.name, data.selectedOptions);
              } else {
                // Single select
                handleChange(field.name, data.optionValue);
              }
            }}
            disabled={field.readOnly}
            size="small"
            className={styles.dropdown}
            style={{ width: `${dropdownWidth}px` }}
          >
            {!isMultiSelect && <Option value="">Select...</Option>}
            {field.choice.choices.map((choice) => (
              <Option key={choice} value={choice}>
                {choice}
              </Option>
            ))}
          </Dropdown>
        </Field>
      );
    }

    if (field.lookup) {
      const options = lookupOptions[field.name] || [];
      const isLoading = lookupLoading[field.name];
      const isMultiSelect = field.lookup.allowMultipleValues;

      // Value is now pre-extracted: number for single, number[] for multi
      const selectedIds: number[] = isMultiSelect
        ? (Array.isArray(value) ? value : [])
        : (typeof value === 'number' ? [value] : []);

      // Get display values for selected items (compare as numbers to handle type mismatches)
      const selectedOptions = options.filter((o) => selectedIds.some((id) => Number(id) === Number(o.id)));
      let displayValue = selectedOptions.map((o) => o.value).join(', ');

      // If options not loaded yet, try to get display value from initial data
      if (!displayValue && selectedIds.length > 0 && options.length === 0) {
        const initialLookupValue = initialValues[field.name];
        if (isMultiSelect && Array.isArray(initialLookupValue)) {
          displayValue = initialLookupValue
            .map((v) => (typeof v === 'object' && v && 'LookupValue' in v ? v.LookupValue : ''))
            .filter(Boolean)
            .join(', ');
        } else if (typeof initialLookupValue === 'object' && initialLookupValue && 'LookupValue' in initialLookupValue) {
          displayValue = (initialLookupValue as { LookupValue: string }).LookupValue;
        }
      }

      // If read-only, show as disabled input
      if (field.readOnly) {
        return (
          <Field key={field.id} label={field.displayName}>
            <Input value={displayValue || '-'} disabled size="small" />
          </Field>
        );
      }

      // Editable lookup - show as dropdown
      const dropdownWidth = getDropdownWidth(options.map((o) => o.value));
      return (
        <Field key={field.id} label={field.displayName} required={field.required}>
          <Dropdown
            value={displayValue}
            selectedOptions={selectedIds.map(String)}
            multiselect={isMultiSelect}
            onOptionSelect={(_e, data) => {
              if (isMultiSelect) {
                // Multi-select: toggle the selected option
                const selectedId = data.optionValue ? parseInt(data.optionValue, 10) : null;
                if (selectedId === null) return;
                const currentIds = Array.isArray(value) ? value : [];
                const newIds = currentIds.includes(selectedId)
                  ? currentIds.filter((id) => id !== selectedId)
                  : [...currentIds, selectedId];
                handleChange(field.name, newIds);
              } else {
                // Single select
                const selectedId = data.optionValue ? parseInt(data.optionValue, 10) : null;
                handleChange(field.name, selectedId);
              }
            }}
            disabled={isLoading}
            placeholder={isLoading ? 'Loading...' : 'Select...'}
            size="small"
            className={styles.dropdown}
            style={{ width: `${dropdownWidth}px` }}
          >
            {!isMultiSelect && <Option value="">Select...</Option>}
            {options.map((option) => (
              <Option key={option.id} value={String(option.id)}>
                {option.value}
              </Option>
            ))}
          </Dropdown>
        </Field>
      );
    }

    // Boolean column
    if (field.boolean) {
      return (
        <Checkbox
          key={field.id}
          checked={Boolean(value)}
          onChange={(_e, data) => handleChange(field.name, data.checked)}
          disabled={field.readOnly}
          label={field.displayName}
        />
      );
    }

    // Number column
    if (field.number) {
      return (
        <Field key={field.id} label={field.displayName} required={field.required}>
          <Input
            type="number"
            value={value !== undefined && value !== null ? String(value) : ''}
            onChange={(_e, data) =>
              handleChange(field.name, data.value ? Number(data.value) : '')
            }
            disabled={field.readOnly}
            size="small"
          />
        </Field>
      );
    }

    // DateTime column
    if (field.dateTime) {
      const isDateOnly = field.dateTime.format === 'dateOnly';
      return (
        <Field key={field.id} label={field.displayName} required={field.required}>
          <Input
            type={isDateOnly ? 'date' : 'datetime-local'}
            value={
              isDateOnly ? formatDateForInput(value) : formatDateTimeForInput(value)
            }
            onChange={(_e, data) => handleChange(field.name, data.value)}
            disabled={field.readOnly}
            size="small"
            className={styles.dateInput}
          />
        </Field>
      );
    }

    // Text column (check for multiline)
    if (field.text?.allowMultipleLines) {
      return (
        <Field key={field.id} label={field.displayName} required={field.required}>
          <Textarea
            value={String(value || '')}
            onChange={(_e, data) => handleChange(field.name, data.value)}
            disabled={field.readOnly}
            rows={3}
            size="small"
          />
        </Field>
      );
    }

    // Default: single-line text
    return (
      <Field key={field.id} label={field.displayName} required={field.required}>
        <Input
          value={String(value || '')}
          onChange={(_e, data) => handleChange(field.name, data.value)}
          disabled={field.readOnly}
          size="small"
        />
      </Field>
    );
  };

  return (
    <Dialog
      open
      modalType="non-modal"
      onOpenChange={(_event, data) => {
        if (!data.open) onClose();
      }}
    >
      <DialogSurface className={styles.dialogSurface}>
        <DialogTitle>{configError ? 'Error' : mode === 'create' ? 'Add New Item' : 'Edit Item'}</DialogTitle>
        <DialogBody className={styles.dialogBody}>
          {loading ? (
            <div className={styles.loadingContainer}>
              <Spinner size="medium" />
              <span>Loading form fields...</span>
            </div>
          ) : configError ? (
            <MessageBar intent="error"><MessageBarBody>{configError}</MessageBarBody></MessageBar>
          ) : (
            <>
              {error && <MessageBar intent="error" style={{ marginBottom: '16px' }}><MessageBarBody>{error}</MessageBarBody></MessageBar>}
              <div className={styles.formFields}>{visibleFields.map(renderField)}</div>
            </>
          )}
        </DialogBody>
        <DialogActions>
          {configError ? (
            <Button onClick={onClose}>Close</Button>
          ) : (
            <>
              <Button appearance="secondary" onClick={onClose} disabled={saving || loading}>Cancel</Button>
              <Button appearance="primary" onClick={handleSubmit} disabled={saving || loading} icon={saving ? <Spinner size="tiny" /> : undefined}>
                {saving ? 'Saving...' : 'Save'}
              </Button>
            </>
          )}
        </DialogActions>
      </DialogSurface>
    </Dialog>
  );
}

// Helper functions

function parseDefaultValue(field: FormFieldConfig): unknown {
  const defaultVal = field.defaultValue?.value;
  if (!defaultVal) return '';

  if (field.boolean) {
    return defaultVal === '1' || defaultVal.toLowerCase() === 'true';
  }

  if (field.number) {
    return Number(defaultVal) || '';
  }

  if (field.dateTime) {
    // Handle SharePoint date formulas
    if (defaultVal === '[today]' || defaultVal === '[Today]') {
      return field.dateTime.format === 'dateOnly'
        ? formatDateForInput(new Date())
        : formatDateTimeForInput(new Date());
    }
    return field.dateTime.format === 'dateOnly'
      ? formatDateForInput(defaultVal)
      : formatDateTimeForInput(defaultVal);
  }

  if (field.choice?.choices) {
    // Choice default - use as-is if it's a valid choice
    if (field.choice.choices.includes(defaultVal)) {
      return defaultVal;
    }
    return '';
  }

  return defaultVal;
}

function getEmptyValue(field: FormFieldConfig): unknown {
  if (field.boolean) return false;
  if (field.number) return '';
  return '';
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
