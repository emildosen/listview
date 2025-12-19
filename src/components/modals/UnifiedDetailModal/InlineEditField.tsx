import { type ReactNode } from 'react';
import {
  makeStyles,
  tokens,
  Spinner,
  Text,
  Tooltip,
  mergeClasses,
} from '@fluentui/react-components';
import { EditRegular, ErrorCircleRegular } from '@fluentui/react-icons';

const useStyles = makeStyles({
  container: {
    position: 'relative',
    display: 'flex',
    alignItems: 'flex-start',
    gap: '8px',
    minHeight: '32px',
  },
  label: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground3,
    minWidth: '100px',
    flexShrink: 0,
    paddingTop: '6px',
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
  savingSpinner: {
    flexShrink: 0,
  },
  errorIcon: {
    color: tokens.colorPaletteRedForeground1,
    flexShrink: 0,
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
  onMouseEnter,
  onMouseLeave,
  onClearError,
}: InlineEditFieldProps) {
  const styles = useStyles();

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

  return (
    <div className={styles.container}>
      <Text className={styles.label}>{label}</Text>
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
          editComponent
        ) : (
          <>
            <span className={styles.value}>{children}</span>
            {isSaving && <Spinner size="tiny" className={styles.savingSpinner} />}
            {error && (
              <Tooltip content={error} relationship="label">
                <ErrorCircleRegular
                  className={styles.errorIcon}
                  onClick={(e) => {
                    e.stopPropagation();
                    onClearError();
                  }}
                />
              </Tooltip>
            )}
            <EditRegular
              className={mergeClasses(styles.editIcon, showEditIcon && styles.editIconVisible)}
            />
          </>
        )}
      </div>
    </div>
  );
}
