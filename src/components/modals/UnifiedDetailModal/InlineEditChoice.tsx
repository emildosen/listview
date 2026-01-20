import { useEffect, useRef, useState } from 'react';
import { makeStyles, Dropdown, Option } from '@fluentui/react-components';

const useStyles = makeStyles({
  dropdown: {
    minWidth: '150px',
  },
});

interface InlineEditChoiceProps {
  value: string | string[];
  choices: string[];
  isMultiSelect?: boolean;
  onChange: (value: string | string[]) => void;
  onCommit: (value?: string | string[]) => void;
  onCancel: () => void;
  placeholder?: string;
}

export function InlineEditChoice({
  value,
  choices,
  isMultiSelect = false,
  onChange,
  onCommit,
  onCancel,
  placeholder = 'Select...',
}: InlineEditChoiceProps) {
  const styles = useStyles();
  const dropdownRef = useRef<HTMLButtonElement>(null);
  const [isOpen, setIsOpen] = useState(false);
  // Track current selections for multi-select (to pass to onCommit when closing)
  const currentValueRef = useRef(value);
  currentValueRef.current = value;

  useEffect(() => {
    // Focus and open dropdown on mount
    dropdownRef.current?.focus();
  }, []);

  const handleKeyDown = (e: React.KeyboardEvent) => {
    if (e.key === 'Escape') {
      e.preventDefault();
      onCancel();
    }
  };

  const handleOptionSelect = (_e: unknown, data: { optionValue?: string; selectedOptions: string[] }) => {
    if (isMultiSelect) {
      // For multi-select, just update the value - don't commit yet
      onChange(data.selectedOptions);
    } else {
      // For single-select, commit immediately
      onChange(data.optionValue || '');
      setTimeout(() => onCommit(data.optionValue || ''), 0);
    }
  };

  const handleOpenChange = (_e: unknown, data: { open: boolean }) => {
    setIsOpen(data.open);

    // For multi-select, commit when dropdown closes
    if (isMultiSelect && !data.open) {
      // Use setTimeout to ensure the value state has been updated
      setTimeout(() => onCommit(currentValueRef.current), 0);
    }
  };

  // Normalize value to array for selectedOptions
  const selectedOptions = isMultiSelect
    ? (Array.isArray(value) ? value : (value ? [value] : []))
    : (value && !Array.isArray(value) ? [value] : []);

  // Display value for the dropdown
  const displayValue = isMultiSelect
    ? (Array.isArray(value) && value.length > 0 ? value.join(', ') : placeholder)
    : (typeof value === 'string' && value ? value : placeholder);

  // Calculate dropdown width based on longest option
  const longestOption = [...choices, placeholder].reduce((a, b) => (a.length > b.length ? a : b), '');
  const dropdownWidth = Math.max(150, longestOption.length * 8 + 60);

  return (
    <Dropdown
      ref={dropdownRef}
      open={isMultiSelect ? isOpen : undefined}
      value={displayValue}
      selectedOptions={selectedOptions}
      multiselect={isMultiSelect}
      onOptionSelect={handleOptionSelect}
      onOpenChange={handleOpenChange}
      onKeyDown={handleKeyDown}
      onBlur={isMultiSelect ? undefined : () => onCommit()}
      className={styles.dropdown}
      style={{ width: `${dropdownWidth}px` }}
      size="small"
    >
      {!isMultiSelect && <Option value="">{placeholder}</Option>}
      {choices.map((choice) => (
        <Option key={choice} value={choice}>
          {choice}
        </Option>
      ))}
    </Dropdown>
  );
}
