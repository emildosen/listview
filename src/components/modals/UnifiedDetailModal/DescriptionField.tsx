import { useState, useCallback, useRef } from 'react';
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
  const [error, setError] = useState<string | null>(null);
  const [saving, setSaving] = useState(false);
  const lastSavedValue = useRef(value);
  const pendingValue = useRef<string | null>(null);

  // Called when editor content changes (on blur)
  const handleChange = useCallback((newValue: string) => {
    pendingValue.current = newValue;
  }, []);

  // Called after onChange, triggers save
  const handleBlur = useCallback(async () => {
    const newValue = pendingValue.current;
    pendingValue.current = null;

    // No pending change or same as last saved
    if (newValue === null || newValue === lastSavedValue.current) {
      return;
    }

    setSaving(true);
    setError(null);
    try {
      await onSave(newValue);
      lastSavedValue.current = newValue;
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Failed to save');
    } finally {
      setSaving(false);
    }
  }, [onSave]);

  const showSaving = isSaving || saving;

  return (
    <div className={styles.container}>
      <Text className={styles.label}>Description</Text>
      <div className={styles.editorWrapper}>
        <RichTextEditor
          value={value}
          onChange={handleChange}
          onBlur={handleBlur}
          placeholder={placeholder}
          readOnly={readOnly}
          minHeight={200}
          showToolbar={isRichText}
        />
        {showSaving && <Spinner size="tiny" className={styles.savingIndicator} />}
      </div>
      {error && <Text className={styles.error}>{error}</Text>}
    </div>
  );
}
