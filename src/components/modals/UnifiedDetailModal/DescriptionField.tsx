import { useState, useEffect, useCallback, useRef } from 'react';
import {
  makeStyles,
  tokens,
  Text,
  Spinner,
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
  editorWrapper: {
    position: 'relative',
  },
  savingIndicator: {
    position: 'absolute',
    top: '12px',
    right: '12px',
    zIndex: 10,
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

  const showSaving = isSaving || saving;

  return (
    <div className={styles.container}>
      <Text className={styles.label}>Description</Text>
      <div className={styles.editorWrapper}>
        <RichTextEditor
          value={localValue}
          onChange={setLocalValue}
          onBlur={handleSave}
          placeholder={placeholder}
          readOnly={readOnly}
          minHeight={80}
          showToolbar={isRichText}
        />
        {showSaving && <Spinner size="tiny" className={styles.savingIndicator} />}
      </div>
      {error && <Text className={styles.error}>{error}</Text>}
    </div>
  );
}
