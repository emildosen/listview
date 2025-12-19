import { useEffect, useRef } from 'react';
import { Checkbox } from '@fluentui/react-components';

interface InlineEditBooleanProps {
  value: boolean;
  label: string;
  onChange: (value: boolean) => void;
  onCommit: () => void;
  onCancel: () => void;
}

export function InlineEditBoolean({
  value,
  label: _label,
  onChange,
  onCommit,
  onCancel,
}: InlineEditBooleanProps) {
  const checkboxRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    checkboxRef.current?.focus();
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

  const handleChange = (_e: unknown, data: { checked: boolean | 'mixed' }) => {
    const newValue = data.checked === true;
    onChange(newValue);
    // Commit immediately after toggle
    setTimeout(() => onCommit(), 0);
  };

  return (
    <Checkbox
      ref={checkboxRef}
      checked={value}
      onChange={handleChange}
      onKeyDown={handleKeyDown}
      onBlur={onCommit}
      label={value ? 'Yes' : 'No'}
    />
  );
}
