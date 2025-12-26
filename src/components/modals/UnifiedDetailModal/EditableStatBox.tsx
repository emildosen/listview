import { useState, useEffect, useRef } from 'react';
import {
  makeStyles,
  tokens,
  Text,
  Spinner,
  Tooltip,
  mergeClasses,
} from '@fluentui/react-components';
import { EditRegular, CheckmarkCircleRegular, DismissCircleRegular } from '@fluentui/react-icons';
import { InlineEditText } from './InlineEditText';
import { InlineEditChoice } from './InlineEditChoice';
import { InlineEditLookup } from './InlineEditLookup';
import { InlineEditNumber } from './InlineEditNumber';
import { InlineEditDate, formatDateForInput, formatDateTimeForInput } from './InlineEditDate';
import { InlineEditBoolean } from './InlineEditBoolean';
import { ClickableLookupValue } from './ClickableLookupValue';
import type { GraphListColumn, FormFieldConfig } from '../../../auth/graphClient';
import type { LookupOption } from '../../../contexts/FormConfigContext';

const useStyles = makeStyles({
  container: {
    position: 'relative',
    display: 'flex',
    flexDirection: 'column',
    gap: '2px',
    padding: '8px 12px',
    backgroundColor: tokens.colorNeutralBackground3,
    borderRadius: tokens.borderRadiusMedium,
    minWidth: '80px',
    cursor: 'pointer',
    transition: 'background-color 0.1s ease',
    '&:hover': {
      backgroundColor: tokens.colorNeutralBackground4,
    },
  },
  containerEditing: {
    cursor: 'default',
    backgroundColor: tokens.colorNeutralBackground1,
    '&:hover': {
      backgroundColor: tokens.colorNeutralBackground1,
    },
  },
  containerReadOnly: {
    cursor: 'default',
    '&:hover': {
      backgroundColor: tokens.colorNeutralBackground3,
    },
  },
  labelRow: {
    display: 'flex',
    alignItems: 'center',
    gap: '4px',
  },
  label: {
    fontSize: tokens.fontSizeBase100,
    color: tokens.colorNeutralForeground3,
    textTransform: 'uppercase',
    letterSpacing: '0.5px',
  },
  statusIcon: {
    fontSize: '12px',
  },
  successIcon: {
    color: tokens.colorPaletteGreenForeground1,
  },
  errorIcon: {
    color: tokens.colorPaletteRedForeground1,
    cursor: 'pointer',
  },
  valueRow: {
    display: 'flex',
    alignItems: 'center',
    gap: '6px',
    minHeight: '24px',
  },
  value: {
    fontSize: tokens.fontSizeBase400,
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground1,
    wordBreak: 'break-word',
    flex: 1,
  },
  editIcon: {
    opacity: 0,
    transition: 'opacity 0.15s ease',
    color: tokens.colorNeutralForeground3,
    fontSize: '14px',
    flexShrink: 0,
  },
  editIconVisible: {
    opacity: 1,
  },
  editContainer: {
    width: '100%',
    minWidth: '120px',
  },
  editRow: {
    display: 'flex',
    alignItems: 'center',
    gap: '6px',
  },
  cancelButton: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    cursor: 'pointer',
    color: tokens.colorNeutralForeground3,
    flexShrink: 0,
    '&:hover': {
      color: tokens.colorNeutralForeground1,
    },
  },
});

interface EditableStatBoxProps {
  fieldName: string;
  label: string;
  value: unknown;
  displayValue: string;
  formField: FormFieldConfig | undefined;
  columnMetadata: GraphListColumn | undefined;
  isEditing: boolean;
  isHovered: boolean;
  isSaving: boolean;
  error: string | null;
  siteId: string;
  siteUrl?: string;
  getLookupOptions: (siteId: string, listId: string, columnName: string) => Promise<LookupOption[]>;
  lookupOptions: Record<string, LookupOption[]>;
  setLookupOptions: React.Dispatch<React.SetStateAction<Record<string, LookupOption[]>>>;
  onStartEdit: () => void;
  onCancelEdit: (fieldName?: string) => void;
  onSave: (fieldName: string, value: unknown) => Promise<void>;
  onMouseEnter: () => void;
  onMouseLeave: () => void;
  onClearError: () => void;
}

export function EditableStatBox({
  fieldName,
  label,
  value,
  displayValue,
  formField,
  columnMetadata,
  isEditing,
  isHovered,
  isSaving,
  error,
  siteId,
  siteUrl,
  getLookupOptions,
  lookupOptions,
  setLookupOptions,
  onStartEdit,
  onCancelEdit,
  onSave,
  onMouseEnter,
  onMouseLeave,
  onClearError,
}: EditableStatBoxProps) {
  const styles = useStyles();
  const [editValue, setEditValue] = useState<unknown>(value);
  const [lookupLoading, setLookupLoading] = useState(false);
  const [showSuccess, setShowSuccess] = useState(false);
  const prevIsSaving = useRef(isSaving);
  const prevError = useRef(error);
  // Ref to track latest edit value for commit (avoids race condition with state updates)
  const editValueRef = useRef<unknown>(value);

  // Sync edit value with prop
  useEffect(() => {
    if (!isEditing) {
      setEditValue(value);
      editValueRef.current = value;
    }
  }, [value, isEditing]);

  // Helper to update both state and ref
  const updateEditValue = (newValue: unknown) => {
    setEditValue(newValue);
    editValueRef.current = newValue;
  };

  // Track when save completes successfully
  useEffect(() => {
    const wasSaving = prevIsSaving.current;
    const hadError = prevError.current;

    prevIsSaving.current = isSaving;
    prevError.current = error;

    if (wasSaving && !isSaving && !error && !hadError) {
      setShowSuccess(true);
    }
  }, [isSaving, error]);

  // Auto-hide success after 2 seconds
  useEffect(() => {
    if (showSuccess) {
      const timer = setTimeout(() => setShowSuccess(false), 2000);
      return () => clearTimeout(timer);
    }
  }, [showSuccess]);

  // Clear success when error appears
  useEffect(() => {
    if (error) {
      setShowSuccess(false);
    }
  }, [error]);

  // Load lookup options when entering edit mode
  useEffect(() => {
    if (!isEditing || !formField?.lookup) return;
    if (lookupOptions[fieldName]) return;

    const loadOptions = async () => {
      setLookupLoading(true);
      try {
        const options = await getLookupOptions(siteId, formField.lookup!.listId, formField.lookup!.columnName);
        setLookupOptions(prev => ({ ...prev, [fieldName]: options }));
      } catch {
        setLookupOptions(prev => ({ ...prev, [fieldName]: [] }));
      } finally {
        setLookupLoading(false);
      }
    };
    loadOptions();
  }, [isEditing, formField, fieldName, siteId, getLookupOptions, lookupOptions, setLookupOptions]);

  const handleCommit = async (directValue?: unknown) => {
    try {
      // Use directly passed value if provided, otherwise use ref (avoids race condition)
      const valueToSave = directValue !== undefined ? directValue : editValueRef.current;
      await onSave(fieldName, valueToSave);
      // Only close if this field is still being edited (user may have clicked another field)
      onCancelEdit(fieldName);
    } catch {
      // Error is handled in parent
    }
  };

  const isReadOnly = formField?.readOnly || columnMetadata?.readOnly;

  const handleClick = () => {
    if (!isReadOnly && !isEditing && !isSaving) {
      onStartEdit();
    }
  };

  const handleKeyDown = (e: React.KeyboardEvent) => {
    if (!isReadOnly && !isEditing && !isSaving && (e.key === 'Enter' || e.key === ' ')) {
      e.preventDefault();
      onStartEdit();
    }
  };

  const showEditIcon = isHovered && !isEditing && !isReadOnly && !isSaving;

  // Render status indicator
  const renderStatus = () => {
    if (isSaving) {
      return <Spinner size="tiny" />;
    }
    if (error) {
      return (
        <Tooltip content={error} relationship="label">
          <DismissCircleRegular
            className={mergeClasses(styles.statusIcon, styles.errorIcon)}
            onClick={(e) => {
              e.stopPropagation();
              onClearError();
            }}
          />
        </Tooltip>
      );
    }
    if (showSuccess) {
      return <CheckmarkCircleRegular className={mergeClasses(styles.statusIcon, styles.successIcon)} />;
    }
    return null;
  };

  // Render edit component based on field type
  const renderEditComponent = () => {
    // Choice field
    if (formField?.choice?.choices) {
      const isMultiSelect = formField.choice.allowMultipleValues ?? false;
      // Normalize value: multi-select uses array, single-select uses string
      // Handle SharePoint's { results: [...] } format
      let normalizedValue: string | string[];
      if (isMultiSelect) {
        if (Array.isArray(editValue)) {
          normalizedValue = editValue.map(String);
        } else if (typeof editValue === 'object' && editValue !== null && 'results' in editValue) {
          normalizedValue = ((editValue as { results: string[] }).results || []).map(String);
        } else if (editValue) {
          normalizedValue = [String(editValue)];
        } else {
          normalizedValue = [];
        }
      } else {
        normalizedValue = String(editValue ?? '');
      }

      return (
        <InlineEditChoice
          value={normalizedValue}
          choices={formField.choice.choices}
          isMultiSelect={isMultiSelect}
          onChange={(v) => updateEditValue(v)}
          onCommit={handleCommit}
          onCancel={onCancelEdit}
        />
      );
    }

    // Lookup field
    if (formField?.lookup) {
      const extractId = (v: unknown): number | null => {
        if (typeof v === 'number') return v;
        if (typeof v === 'object' && v !== null && 'LookupId' in v) {
          return (v as { LookupId: number }).LookupId;
        }
        return null;
      };

      const currentId = formField.lookup.allowMultipleValues
        ? (Array.isArray(editValue) ? editValue.map(extractId).filter((id): id is number => id !== null) : [])
        : extractId(editValue);

      return (
        <InlineEditLookup
          value={currentId}
          options={lookupOptions[fieldName] ?? []}
          isLoading={lookupLoading}
          isMultiSelect={formField.lookup.allowMultipleValues ?? false}
          onChange={(v) => updateEditValue(v)}
          onCommit={handleCommit}
          onCancel={onCancelEdit}
        />
      );
    }

    // Boolean field
    if (formField?.boolean || columnMetadata?.boolean) {
      return (
        <InlineEditBoolean
          value={Boolean(editValue)}
          onChange={(v) => updateEditValue(v)}
          onCommit={handleCommit}
          onCancel={onCancelEdit}
        />
      );
    }

    // Number field
    if (formField?.number || columnMetadata?.number) {
      return (
        <InlineEditNumber
          value={typeof editValue === 'number' ? editValue : null}
          onChange={(v) => updateEditValue(v)}
          onCommit={handleCommit}
          onCancel={onCancelEdit}
        />
      );
    }

    // DateTime field
    if (formField?.dateTime || columnMetadata?.dateTime) {
      const isDateOnly = (formField?.dateTime?.format ?? columnMetadata?.dateTime?.format) === 'dateOnly';
      const formattedValue = isDateOnly
        ? formatDateForInput(editValue)
        : formatDateTimeForInput(editValue);

      return (
        <InlineEditDate
          value={formattedValue}
          dateOnly={isDateOnly}
          onChange={(v) => updateEditValue(v)}
          onCommit={handleCommit}
          onCancel={onCancelEdit}
        />
      );
    }

    // Multiline text
    if (formField?.text?.allowMultipleLines || columnMetadata?.text?.allowMultipleLines) {
      return (
        <InlineEditText
          value={String(editValue ?? '')}
          multiline
          onChange={(v) => updateEditValue(v)}
          onCommit={handleCommit}
          onCancel={onCancelEdit}
        />
      );
    }

    // Default: single line text
    return (
      <InlineEditText
        value={String(editValue ?? '')}
        onChange={(v) => updateEditValue(v)}
        onCommit={handleCommit}
        onCancel={onCancelEdit}
      />
    );
  };

  return (
    <div
      className={mergeClasses(
        styles.container,
        isEditing && styles.containerEditing,
        isReadOnly && styles.containerReadOnly
      )}
      onClick={handleClick}
      onKeyDown={handleKeyDown}
      onMouseEnter={onMouseEnter}
      onMouseLeave={onMouseLeave}
      role={isReadOnly ? undefined : 'button'}
      tabIndex={isReadOnly || isEditing ? -1 : 0}
      aria-label={isReadOnly ? undefined : `Edit ${label}`}
    >
      <div className={styles.labelRow}>
        <Text className={styles.label}>{label}</Text>
        {renderStatus()}
      </div>
      {isEditing ? (
        <div className={styles.editContainer}>
          <div className={styles.editRow}>
            {renderEditComponent()}
            <DismissCircleRegular
              className={styles.cancelButton}
              onMouseDown={(e) => {
                e.preventDefault(); // Prevent blur from firing on the input
                e.stopPropagation();
                onCancelEdit();
              }}
              title="Cancel (Esc)"
            />
          </div>
        </div>
      ) : (
        <div className={styles.valueRow}>
          <Text className={styles.value}>
            {(() => {
              const lookupInfo = formField?.lookup ?? columnMetadata?.lookup;
              if (lookupInfo?.listId && value !== null && value !== undefined) {
                return (
                  <ClickableLookupValue
                    value={value}
                    targetListId={lookupInfo.listId}
                    targetListName={label}
                    siteId={siteId}
                    siteUrl={siteUrl}
                    isMultiSelect={lookupInfo.allowMultipleValues ?? false}
                  />
                );
              }
              return displayValue || '-';
            })()}
          </Text>
          <EditRegular
            className={mergeClasses(styles.editIcon, showEditIcon && styles.editIconVisible)}
          />
        </div>
      )}
    </div>
  );
}
