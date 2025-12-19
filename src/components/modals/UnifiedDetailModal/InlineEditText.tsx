import { useEffect, useRef } from 'react';
import { makeStyles, Input, Textarea } from '@fluentui/react-components';

const useStyles = makeStyles({
  input: {
    width: '100%',
    minWidth: '200px',
  },
  textarea: {
    width: '100%',
    minWidth: '200px',
    minHeight: '80px',
  },
});

interface InlineEditTextProps {
  value: string;
  onChange: (value: string) => void;
  onCommit: () => void;
  onCancel: () => void;
  multiline?: boolean;
  placeholder?: string;
}

export function InlineEditText({
  value,
  onChange,
  onCommit,
  onCancel,
  multiline = false,
  placeholder,
}: InlineEditTextProps) {
  const styles = useStyles();
  const inputRef = useRef<HTMLInputElement | HTMLTextAreaElement>(null);

  useEffect(() => {
    // Focus and select on mount
    const input = inputRef.current;
    if (input) {
      input.focus();
      input.select();
    }
  }, []);

  const handleKeyDown = (e: React.KeyboardEvent) => {
    if (e.key === 'Escape') {
      e.preventDefault();
      onCancel();
    } else if (e.key === 'Enter' && !multiline) {
      e.preventDefault();
      onCommit();
    } else if (e.key === 'Enter' && multiline && (e.ctrlKey || e.metaKey)) {
      e.preventDefault();
      onCommit();
    }
  };

  const handleBlur = () => {
    onCommit();
  };

  if (multiline) {
    return (
      <Textarea
        ref={inputRef as React.RefObject<HTMLTextAreaElement>}
        value={value}
        onChange={(_e, data) => onChange(data.value)}
        onKeyDown={handleKeyDown}
        onBlur={handleBlur}
        placeholder={placeholder}
        className={styles.textarea}
        size="small"
        resize="vertical"
      />
    );
  }

  return (
    <Input
      ref={inputRef as React.RefObject<HTMLInputElement>}
      value={value}
      onChange={(_e, data) => onChange(data.value)}
      onKeyDown={handleKeyDown}
      onBlur={handleBlur}
      placeholder={placeholder}
      className={styles.input}
      size="small"
    />
  );
}
