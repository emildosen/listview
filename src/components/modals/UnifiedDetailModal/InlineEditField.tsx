import { type ReactNode, useState, useEffect, useRef } from 'react';
import {
  makeStyles,
  tokens,
  Spinner,
  Text,
  Tooltip,
  mergeClasses,
} from '@fluentui/react-components';
import { EditRegular, CheckmarkCircleRegular, DismissCircleRegular } from '@fluentui/react-icons';

const useStyles = makeStyles({
  container: {
    position: 'relative',
    display: 'flex',
    alignItems: 'flex-start',
    gap: '8px',
    minHeight: '32px',
  },
  labelContainer: {
    display: 'flex',
    alignItems: 'center',
    gap: '6px',
    minWidth: '100px',
    flexShrink: 0,
    paddingTop: '6px',
  },
  label: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground3,
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
  valueContainer: {
    flex: 1,
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
    minHeight: '32px',
    padding: '4px 8px',
    borderRadius: tokens.borderRadiusMedium,
    cursor: 'pointer',
    transition: 'background-color 0.1s ease',
    '&:hover': {
      backgroundColor: tokens.colorNeutralBackground3,
    },
  },
  valueContainerEditing: {
    backgroundColor: tokens.colorNeutralBackground1,
    cursor: 'default',
    '&:hover': {
      backgroundColor: tokens.colorNeutralBackground1,
    },
  },
  valueContainerReadOnly: {
    cursor: 'default',
    '&:hover': {
      backgroundColor: 'transparent',
    },
  },
  value: {
    flex: 1,
    wordBreak: 'break-word',
  },
  editIcon: {
    opacity: 0,
    transition: 'opacity 0.15s ease',
    color: tokens.colorNeutralForeground3,
    flexShrink: 0,
  },
  editIconVisible: {
    opacity: 1,
  },
  editRow: {
    display: 'flex',
    alignItems: 'center',
    gap: '6px',
    flex: 1,
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

interface InlineEditFieldProps {
  label: string;
  isEditing: boolean;
  isHovered: boolean;
  isSaving: boolean;
  error: string | null;
  readOnly?: boolean;
  children: ReactNode;
  editComponent: ReactNode;
  onStartEdit: () => void;
  onCancel?: () => void;
  onMouseEnter: () => void;
  onMouseLeave: () => void;
  onClearError: () => void;
}

export function InlineEditField({
  label,
  isEditing,
  isHovered,
  isSaving,
  error,
  readOnly = false,
  children,
  editComponent,
  onStartEdit,
  onCancel,
  onMouseEnter,
  onMouseLeave,
  onClearError,
}: InlineEditFieldProps) {
  const styles = useStyles();
  const [showSuccess, setShowSuccess] = useState(false);
  const prevIsSaving = useRef(isSaving);
  const prevError = useRef(error);

  // Track when save completes successfully (isSaving: true → false, no error)
  useEffect(() => {
    const wasSaving = prevIsSaving.current;
    const hadError = prevError.current;

    // Update refs for next render
    prevIsSaving.current = isSaving;
    prevError.current = error;

    // Detect successful save completion
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

  const handleClick = () => {
    if (!readOnly && !isEditing && !isSaving) {
      onStartEdit();
    }
  };

  const handleKeyDown = (e: React.KeyboardEvent) => {
    if (!readOnly && !isEditing && !isSaving && (e.key === 'Enter' || e.key === ' ')) {
      e.preventDefault();
      onStartEdit();
    }
  };

  const showEditIcon = isHovered && !isEditing && !readOnly && !isSaving;

  // Render label status indicator (spinner, checkmark, or error)
  const renderLabelStatus = () => {
    if (isSaving) {
      return <Spinner size="tiny" className={styles.savingSpinner} />;
    }
    if (error) {
      return (
        <Tooltip content={error} relationship="label">
          <DismissCircleRegular
            className={styles.errorIcon}
            onClick={(e) => {
              e.stopPropagation();
              onClearError();
            }}
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
      <div
        className={mergeClasses(
          styles.valueContainer,
          isEditing && styles.valueContainerEditing,
          readOnly && styles.valueContainerReadOnly
        )}
        onClick={handleClick}
        onKeyDown={handleKeyDown}
        onMouseEnter={onMouseEnter}
        onMouseLeave={onMouseLeave}
        role={readOnly ? undefined : 'button'}
        tabIndex={readOnly || isEditing ? -1 : 0}
        aria-label={readOnly ? undefined : `Edit ${label}`}
      >
        {isEditing ? (
          <div className={styles.editRow}>
            {editComponent}
            {onCancel && (
              <DismissCircleRegular
                className={styles.cancelButton}
                onMouseDown={(e) => {
                  e.preventDefault(); // Prevent blur from firing on the input
                  e.stopPropagation();
                  onCancel();
                }}
                title="Cancel (Esc)"
              />
            )}
          </div>
        ) : (
          <>
            <span className={styles.value}>{children}</span>
            <EditRegular
              className={mergeClasses(styles.editIcon, showEditIcon && styles.editIconVisible)}
            />
          </>
        )}
      </div>
    </div>
  );
}
