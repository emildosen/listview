import { useEffect, useRef } from 'react';
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

  useEffect(() => {
    // Focus dropdown on mount
    dropdownRef.current?.focus();
  }, []);

  const handleKeyDown = (e: React.KeyboardEvent) => {
    if (e.key === 'Escape') {
      e.preventDefault();
      onCancel();
    }
    // For multi-select, Enter commits
    if (e.key === 'Enter' && isMultiSelect) {
      e.preventDefault();
      onCommit();
    }
  };

  const handleOptionSelect = (_e: unknown, data: { optionValue?: string; selectedOptions: string[] }) => {
    if (isMultiSelect) {
      // For multi-select, update with the full selected options array
      onChange(data.selectedOptions);
    } else {
      // For single-select, commit immediately
      onChange(data.optionValue || '');
      setTimeout(() => onCommit(data.optionValue || ''), 0);
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
      value={displayValue}
      selectedOptions={selectedOptions}
      multiselect={isMultiSelect}
      onOptionSelect={handleOptionSelect}
      onKeyDown={handleKeyDown}
      onBlur={() => onCommit()}
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
