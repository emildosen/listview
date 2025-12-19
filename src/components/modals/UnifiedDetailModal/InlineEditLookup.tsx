import { useEffect, useRef } from 'react';
import { makeStyles, Dropdown, Option, Spinner } from '@fluentui/react-components';
import type { LookupOption } from '../../../contexts/FormConfigContext';

const useStyles = makeStyles({
  dropdown: {
    minWidth: '180px',
  },
  loadingOption: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
  },
});

interface InlineEditLookupProps {
  value: number | number[] | null;
  options: LookupOption[];
  isLoading: boolean;
  isMultiSelect: boolean;
  onChange: (value: number | number[] | null) => void;
  onCommit: () => void;
  onCancel: () => void;
  placeholder?: string;
}

export function InlineEditLookup({
  value,
  options,
  isLoading,
  isMultiSelect,
  onChange,
  onCommit,
  onCancel,
  placeholder = 'Select...',
}: InlineEditLookupProps) {
  const styles = useStyles();
  const dropdownRef = useRef<HTMLButtonElement>(null);

  useEffect(() => {
    dropdownRef.current?.focus();
  }, []);

  const handleKeyDown = (e: React.KeyboardEvent) => {
    if (e.key === 'Escape') {
      e.preventDefault();
      onCancel();
    }
  };

  const handleOptionSelect = (_e: unknown, data: { optionValue?: string; selectedOptions?: string[] }) => {
    if (isMultiSelect) {
      // For multi-select, parse all selected options
      const selectedIds = (data.selectedOptions || [])
        .filter((v) => v !== '')
        .map((v) => parseInt(v, 10))
        .filter((id) => !isNaN(id));
      onChange(selectedIds);
    } else {
      // For single select, parse the selected option
      const selectedId = data.optionValue ? parseInt(data.optionValue, 10) : null;
      onChange(isNaN(selectedId!) ? null : selectedId);
      // Commit immediately for single select
      setTimeout(() => onCommit(), 0);
    }
  };

  // Get display value
  const selectedIds: number[] = isMultiSelect
    ? (Array.isArray(value) ? value : [])
    : (typeof value === 'number' ? [value] : []);

  const selectedOptions = options.filter((o) =>
    selectedIds.some((id) => Number(id) === Number(o.id))
  );
  const displayValue = selectedOptions.map((o) => o.value).join(', ') || placeholder;

  // Calculate dropdown width
  const allLabels = [placeholder, ...options.map((o) => o.value)];
  const longestLabel = allLabels.reduce((a, b) => (a.length > b.length ? a : b), '');
  const dropdownWidth = Math.max(180, longestLabel.length * 8 + 60);

  if (isLoading) {
    return (
      <Dropdown
        value="Loading..."
        selectedOptions={[]}
        disabled
        className={styles.dropdown}
        style={{ width: `${dropdownWidth}px` }}
        size="small"
      >
        <Option text="Loading..." value="">
          <span className={styles.loadingOption}>
            <Spinner size="tiny" /> Loading...
          </span>
        </Option>
      </Dropdown>
    );
  }

  return (
    <Dropdown
      ref={dropdownRef}
      value={displayValue}
      selectedOptions={selectedIds.map(String)}
      multiselect={isMultiSelect}
      onOptionSelect={handleOptionSelect}
      onKeyDown={handleKeyDown}
      onBlur={() => {
        // Only commit on blur for multi-select
        if (isMultiSelect) {
          onCommit();
        }
      }}
      className={styles.dropdown}
      style={{ width: `${dropdownWidth}px` }}
      size="small"
    >
      {!isMultiSelect && <Option value="">{placeholder}</Option>}
      {options.map((option) => (
        <Option key={option.id} value={String(option.id)}>
          {option.value}
        </Option>
      ))}
    </Dropdown>
  );
}
