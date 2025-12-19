import { useState, useRef, useEffect } from 'react';
import {
  makeStyles,
  tokens,
  Text,
  Textarea,
  Spinner,
  mergeClasses,
} from '@fluentui/react-components';
import { EditRegular } from '@fluentui/react-icons';

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
  contentWrapper: {
    position: 'relative',
    minHeight: '80px',
    padding: '12px 16px',
    borderRadius: tokens.borderRadiusMedium,
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    backgroundColor: tokens.colorNeutralBackground2,
    cursor: 'pointer',
    transitionProperty: 'border-color, background-color',
    transitionDuration: '0.15s',
    transitionTimingFunction: 'ease',
  },
  contentWrapperHover: {
    border: `1px solid ${tokens.colorNeutralStroke1Hover}`,
    backgroundColor: tokens.colorNeutralBackground3,
  },
  contentWrapperEditing: {
    cursor: 'default',
    border: `1px solid ${tokens.colorBrandStroke1}`,
    backgroundColor: tokens.colorNeutralBackground1,
  },
  contentWrapperReadOnly: {
    cursor: 'default',
  },
  placeholder: {
    color: tokens.colorNeutralForeground4,
    fontStyle: 'italic',
  },
  content: {
    whiteSpace: 'pre-wrap',
    wordBreak: 'break-word',
    lineHeight: '1.5',
  },
  richContent: {
    '& p': {
      margin: '0 0 8px 0',
    },
    '& p:last-child': {
      marginBottom: 0,
    },
    '& ul, & ol': {
      margin: '0 0 8px 0',
      paddingLeft: '24px',
    },
    '& a': {
      color: tokens.colorBrandForegroundLink,
      textDecoration: 'none',
      '&:hover': {
        textDecoration: 'underline',
      },
    },
  },
  editIcon: {
    position: 'absolute',
    top: '12px',
    right: '12px',
    opacity: 0,
    transition: 'opacity 0.15s ease',
    color: tokens.colorNeutralForeground3,
  },
  editIconVisible: {
    opacity: 1,
  },
  textarea: {
    width: '100%',
    minHeight: '120px',
    resize: 'vertical',
  },
  savingIndicator: {
    position: 'absolute',
    top: '12px',
    right: '12px',
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
  const [isEditing, setIsEditing] = useState(false);
  const [isHovered, setIsHovered] = useState(false);
  const [editValue, setEditValue] = useState(value);
  const [error, setError] = useState<string | null>(null);
  const textareaRef = useRef<HTMLTextAreaElement>(null);

  // Sync edit value with prop when not editing
  useEffect(() => {
    if (!isEditing) {
      setEditValue(value);
    }
  }, [value, isEditing]);

  const handleStartEdit = () => {
    if (readOnly || isEditing || isSaving) return;
    setEditValue(value);
    setIsEditing(true);
    setError(null);
    setTimeout(() => {
      textareaRef.current?.focus();
      textareaRef.current?.select();
    }, 0);
  };

  const handleCommit = async () => {
    if (editValue === value) {
      setIsEditing(false);
      return;
    }

    try {
      await onSave(editValue);
      setIsEditing(false);
      setError(null);
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Failed to save');
    }
  };

  const handleCancel = () => {
    setEditValue(value);
    setIsEditing(false);
    setError(null);
  };

  const handleKeyDown = (e: React.KeyboardEvent) => {
    if (e.key === 'Escape') {
      e.preventDefault();
      handleCancel();
    } else if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) {
      e.preventDefault();
      handleCommit();
    }
  };

  const showEditIcon = isHovered && !isEditing && !readOnly && !isSaving;

  const renderContent = () => {
    if (!value) {
      return <span className={styles.placeholder}>{placeholder}</span>;
    }

    if (isRichText) {
      // Render HTML content safely (SharePoint rich text)
      return (
        <div
          className={mergeClasses(styles.content, styles.richContent)}
          dangerouslySetInnerHTML={{ __html: value }}
        />
      );
    }

    return <span className={styles.content}>{value}</span>;
  };

  return (
    <div className={styles.container}>
      <Text className={styles.label}>Description</Text>
      <div
        className={mergeClasses(
          styles.contentWrapper,
          isHovered && !isEditing && !readOnly && styles.contentWrapperHover,
          isEditing && styles.contentWrapperEditing,
          readOnly && styles.contentWrapperReadOnly
        )}
        onClick={handleStartEdit}
        onMouseEnter={() => setIsHovered(true)}
        onMouseLeave={() => setIsHovered(false)}
        role={readOnly ? undefined : 'button'}
        tabIndex={readOnly || isEditing ? -1 : 0}
        onKeyDown={(e) => {
          if (!readOnly && !isEditing && (e.key === 'Enter' || e.key === ' ')) {
            e.preventDefault();
            handleStartEdit();
          }
        }}
        aria-label={readOnly ? undefined : 'Edit description'}
      >
        {isEditing ? (
          <Textarea
            ref={textareaRef}
            value={editValue}
            onChange={(_e, data) => setEditValue(data.value)}
            onKeyDown={handleKeyDown}
            onBlur={handleCommit}
            placeholder={placeholder}
            className={styles.textarea}
            resize="vertical"
          />
        ) : (
          <>
            {renderContent()}
            {isSaving && <Spinner size="tiny" className={styles.savingIndicator} />}
            <EditRegular
              className={mergeClasses(styles.editIcon, showEditIcon && styles.editIconVisible)}
            />
          </>
        )}
      </div>
      {error && (
        <Text style={{ color: tokens.colorPaletteRedForeground1, fontSize: tokens.fontSizeBase200 }}>
          {error}
        </Text>
      )}
    </div>
  );
}
