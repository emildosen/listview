import { useEffect, useRef } from 'react';
import { makeStyles, Dropdown, Option } from '@fluentui/react-components';

const useStyles = makeStyles({
  dropdown: {
    minWidth: '150px',
  },
});

interface InlineEditChoiceProps {
  value: string;
  choices: string[];
  onChange: (value: string) => void;
  onCommit: () => void;
  onCancel: () => void;
  placeholder?: string;
}

export function InlineEditChoice({
  value,
  choices,
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
  };

  const handleOptionSelect = (_e: unknown, data: { optionValue?: string }) => {
    onChange(data.optionValue || '');
    // Commit immediately after selection
    setTimeout(() => onCommit(), 0);
  };

  // Calculate dropdown width based on longest option
  const longestOption = [...choices, placeholder].reduce((a, b) => (a.length > b.length ? a : b), '');
  const dropdownWidth = Math.max(150, longestOption.length * 8 + 60);

  return (
    <Dropdown
      ref={dropdownRef}
      value={value || placeholder}
      selectedOptions={value ? [value] : []}
      onOptionSelect={handleOptionSelect}
      onKeyDown={handleKeyDown}
      onBlur={onCommit}
      className={styles.dropdown}
      style={{ width: `${dropdownWidth}px` }}
      size="small"
    >
      <Option value="">{placeholder}</Option>
      {choices.map((choice) => (
        <Option key={choice} value={choice}>
          {choice}
        </Option>
      ))}
    </Dropdown>
  );
}
