import { useState, useCallback, useRef, useEffect } from 'react';
import {
  makeStyles,
  tokens,
  Text,
  Spinner,
  Tooltip,
} from '@fluentui/react-components';
import { CheckmarkCircleRegular, DismissCircleRegular } from '@fluentui/react-icons';
import { RichTextEditor } from '../../common/RichTextEditor';

const useStyles = makeStyles({
  container: {
    marginTop: '16px',
    marginBottom: '16px',
  },
  labelContainer: {
    display: 'flex',
    alignItems: 'center',
    gap: '6px',
    marginBottom: '8px',
  },
  label: {
    fontSize: tokens.fontSizeBase200,
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground3,
    textTransform: 'uppercase',
    letterSpacing: '0.02em',
  },
  labelStatus: {
    display: 'flex',
    alignItems: 'center',
  },
  successIcon: {
    color: tokens.colorPaletteGreenForeground1,
    fontSize: '14px',
  },
  errorIcon: {
    color: tokens.colorPaletteRedForeground1,
    fontSize: '14px',
    cursor: 'pointer',
  },
  savingSpinner: {
    flexShrink: 0,
  },
  editorWrapper: {
    position: 'relative',
  },
});

interface DescriptionFieldProps {
  value: string;
  isRichText: boolean;
  isSaving?: boolean;
  readOnly?: boolean;
  placeholder?: string;
  label?: string;
  onSave: (value: string) => Promise<void>;
}

export function DescriptionField({
  value,
  isRichText,
  isSaving = false,
  readOnly = false,
  placeholder = 'Add a description...',
  label = 'Description',
  onSave,
}: DescriptionFieldProps) {
  const styles = useStyles();
  const [error, setError] = useState<string | null>(null);
  const [saving, setSaving] = useState(false);
  const [showSuccess, setShowSuccess] = useState(false);
  const lastSavedValue = useRef(value);
  const pendingValue = useRef<string | null>(null);

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
    setShowSuccess(false);
    try {
      await onSave(newValue);
      lastSavedValue.current = newValue;
      setShowSuccess(true);
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Failed to save');
    } finally {
      setSaving(false);
    }
  }, [onSave]);

  const handleClearError = useCallback(() => {
    setError(null);
  }, []);

  const showSaving = isSaving || saving;

  // Render label status indicator (spinner, checkmark, or error)
  const renderLabelStatus = () => {
    if (showSaving) {
      return <Spinner size="tiny" className={styles.savingSpinner} />;
    }
    if (error) {
      return (
        <Tooltip content={error} relationship="label">
          <DismissCircleRegular
            className={styles.errorIcon}
            onClick={handleClearError}
          />
        </Tooltip>
      );
    }
    if (showSuccess) {
      return <CheckmarkCircleRegular className={styles.successIcon} />;
    }
    return null;
  };

  return (
    <div className={styles.container}>
      <div className={styles.labelContainer}>
        <Text className={styles.label}>{label}</Text>
        <span className={styles.labelStatus}>{renderLabelStatus()}</span>
      </div>
      <div className={styles.editorWrapper}>
        <RichTextEditor
          value={value}
          onChange={handleChange}
          onBlur={handleBlur}
          placeholder={placeholder}
          readOnly={readOnly}
          minHeight={100}
          showToolbar={isRichText}
        />
      </div>
    </div>
  );
}
