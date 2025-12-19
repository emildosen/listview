import { useEffect, useRef } from 'react';
import { makeStyles, Input } from '@fluentui/react-components';

const useStyles = makeStyles({
  input: {
    width: '180px',
  },
});

interface InlineEditDateProps {
  value: string;
  dateOnly?: boolean;
  onChange: (value: string) => void;
  onCommit: () => void;
  onCancel: () => void;
}

export function InlineEditDate({
  value,
  dateOnly = false,
  onChange,
  onCommit,
  onCancel,
}: InlineEditDateProps) {
  const styles = useStyles();
  const inputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    inputRef.current?.focus();
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

  return (
    <Input
      ref={inputRef}
      type={dateOnly ? 'date' : 'datetime-local'}
      value={value}
      onChange={(_e, data) => onChange(data.value)}
      onKeyDown={handleKeyDown}
      onBlur={onCommit}
      className={styles.input}
      size="small"
    />
  );
}

// Helper functions for date formatting
export function formatDateForInput(value: unknown): string {
  if (!value) return '';
  if (typeof value === 'string') {
    const match = value.match(/^(\d{4}-\d{2}-\d{2})/);
    if (match) return match[1];
    const date = new Date(value);
    if (!isNaN(date.getTime())) {
      return formatLocalDate(date);
    }
  }
  if (value instanceof Date) {
    return formatLocalDate(value);
  }
  return '';
}

export function formatDateTimeForInput(value: unknown): string {
  if (!value) return '';
  if (typeof value === 'string') {
    const date = new Date(value);
    if (!isNaN(date.getTime())) {
      return formatLocalDateTime(date);
    }
  }
  if (value instanceof Date) {
    return formatLocalDateTime(value);
  }
  return '';
}

function formatLocalDate(d: Date): string {
  const year = d.getFullYear();
  const month = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

function formatLocalDateTime(d: Date): string {
  const hours = String(d.getHours()).padStart(2, '0');
  const minutes = String(d.getMinutes()).padStart(2, '0');
  return `${formatLocalDate(d)}T${hours}:${minutes}`;
}

export function formatDateForDisplay(value: unknown): string {
  if (!value) return '-';
  const date = value instanceof Date ? value : new Date(String(value));
  if (isNaN(date.getTime())) return String(value);
  return date.toLocaleDateString();
}

export function formatDateTimeForDisplay(value: unknown): string {
  if (!value) return '-';
  const date = value instanceof Date ? value : new Date(String(value));
  if (isNaN(date.getTime())) return String(value);
  return date.toLocaleString();
}
