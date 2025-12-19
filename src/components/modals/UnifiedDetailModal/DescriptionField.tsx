import { useState, useRef, useEffect, useCallback } from 'react';
import {
  makeStyles,
  tokens,
  Text,
  Textarea,
  Spinner,
  mergeClasses,
} from '@fluentui/react-components';
import { RichTextEditor } from '../../common/RichTextEditor';

const useStyles = makeStyles({
  container: {
    marginTop: '16px',
    marginBottom: '16px',
  },
  label: {
    display: 'block',
    fontSize: tokens.fontSizeBase200,
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground3,
    textTransform: 'uppercase',
    letterSpacing: '0.02em',
    marginBottom: '8px',
  },
  textareaWrapper: {
    position: 'relative',
    minHeight: '80px',
    borderRadius: tokens.borderRadiusMedium,
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    backgroundColor: tokens.colorNeutralBackground2,
    transitionProperty: 'border-color, background-color',
    transitionDuration: '0.15s',
    transitionTimingFunction: 'ease',
    ':hover': {
      border: `1px solid ${tokens.colorNeutralStroke1Hover}`,
    },
  },
  textareaWrapperFocused: {
    border: `1px solid ${tokens.colorBrandStroke1}`,
    backgroundColor: tokens.colorNeutralBackground1,
  },
  textareaWrapperReadOnly: {
    cursor: 'default',
  },
  textarea: {
    width: '100%',
    minHeight: '80px',
    resize: 'vertical',
    border: 'none',
    backgroundColor: 'transparent',
    '& textarea': {
      backgroundColor: 'transparent',
    },
  },
  savingIndicator: {
    position: 'absolute',
    top: '12px',
    right: '12px',
  },
  error: {
    marginTop: '4px',
    color: tokens.colorPaletteRedForeground1,
    fontSize: tokens.fontSizeBase200,
  },
});

interface DescriptionFieldProps {
  value: string;
  isRichText: boolean;
  isSaving?: boolean;
  readOnly?: boolean;
  placeholder?: string;
  onSave: (value: string) => Promise<void>;
}

export function DescriptionField({
  value,
  isRichText,
  isSaving = false,
  readOnly = false,
  placeholder = 'Add a description...',
  onSave,
}: DescriptionFieldProps) {
  const styles = useStyles();
  const [localValue, setLocalValue] = useState(value);
  const [error, setError] = useState<string | null>(null);
  const [saving, setSaving] = useState(false);
  const textareaRef = useRef<HTMLTextAreaElement>(null);
  const lastSavedValue = useRef(value);

  // Sync local value with prop when external changes occur
  useEffect(() => {
    if (value !== lastSavedValue.current) {
      setLocalValue(value);
      lastSavedValue.current = value;
    }
  }, [value]);

  const handleSave = useCallback(async () => {
    if (localValue === lastSavedValue.current) {
      return; // No changes to save
    }

    setSaving(true);
    setError(null);
    try {
      await onSave(localValue);
      lastSavedValue.current = localValue;
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Failed to save');
    } finally {
      setSaving(false);
    }
  }, [localValue, onSave]);

  const handleTextareaChange = useCallback((_e: unknown, data: { value: string }) => {
    setLocalValue(data.value);
  }, []);

  const handleTextareaKeyDown = useCallback((e: React.KeyboardEvent) => {
    if (e.key === 'Escape') {
      e.preventDefault();
      setLocalValue(lastSavedValue.current);
      textareaRef.current?.blur();
    } else if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) {
      e.preventDefault();
      handleSave();
      textareaRef.current?.blur();
    }
  }, [handleSave]);

  const showSaving = isSaving || saving;

  // Rich text: use TinyMCE editor
  if (isRichText) {
    return (
      <div className={styles.container}>
        <Text className={styles.label}>Description</Text>
        <div style={{ position: 'relative' }}>
          <RichTextEditor
            value={localValue}
            onChange={setLocalValue}
            onBlur={handleSave}
            placeholder={placeholder}
            readOnly={readOnly}
            minHeight={80}
          />
          {showSaving && <Spinner size="tiny" className={styles.savingIndicator} />}
        </div>
        {error && <Text className={styles.error}>{error}</Text>}
      </div>
    );
  }

  // Plain text: use Fluent UI Textarea (always visible, no mode toggle)
  return (
    <div className={styles.container}>
      <Text className={styles.label}>Description</Text>
      <div
        className={mergeClasses(
          styles.textareaWrapper,
          readOnly && styles.textareaWrapperReadOnly
        )}
      >
        <Textarea
          ref={textareaRef}
          value={localValue}
          onChange={handleTextareaChange}
          onKeyDown={handleTextareaKeyDown}
          onBlur={handleSave}
          placeholder={placeholder}
          className={styles.textarea}
          resize="vertical"
          disabled={readOnly}
          appearance="outline"
        />
        {showSaving && <Spinner size="tiny" className={styles.savingIndicator} />}
      </div>
      {error && <Text className={styles.error}>{error}</Text>}
    </div>
  );
}
