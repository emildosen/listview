import { useEffect, useRef } from 'react';
import { makeStyles, Input } from '@fluentui/react-components';

const useStyles = makeStyles({
  input: {
    width: '120px',
  },
});

interface InlineEditNumberProps {
  value: number | null;
  onChange: (value: number | null) => void;
  onCommit: () => void;
  onCancel: () => void;
  placeholder?: string;
}

export function InlineEditNumber({
  value,
  onChange,
  onCommit,
  onCancel,
  placeholder,
}: InlineEditNumberProps) {
  const styles = useStyles();
  const inputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
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
    } else if (e.key === 'Enter') {
      e.preventDefault();
      onCommit();
    }
  };

  const handleChange = (_e: unknown, data: { value: string }) => {
    if (data.value === '') {
      onChange(null);
    } else {
      const num = Number(data.value);
      if (!isNaN(num)) {
        onChange(num);
      }
    }
  };

  return (
    <Input
      ref={inputRef}
      type="number"
      value={value !== null && value !== undefined ? String(value) : ''}
      onChange={handleChange}
      onKeyDown={handleKeyDown}
      onBlur={onCommit}
      placeholder={placeholder}
      className={styles.input}
      size="small"
    />
  );
}
