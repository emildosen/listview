import {
  Dropdown,
  Option,
} from '@fluentui/react-components';
import {
  TableRegular,
  ChartMultipleRegular,
  DismissCircleRegular,
} from '@fluentui/react-icons';
import type { WebPartType } from '../../types/page';

interface WebPartOption {
  value: WebPartType | '';
  label: string;
  icon: React.ReactNode;
}

const webPartOptions: WebPartOption[] = [
  { value: '', label: 'None (empty)', icon: <DismissCircleRegular /> },
  { value: 'list-items', label: 'List Items', icon: <TableRegular /> },
  { value: 'chart', label: 'Chart', icon: <ChartMultipleRegular /> },
];

interface WebPartPickerProps {
  value: WebPartType | null;
  onChange: (type: WebPartType | null) => void;
}

export default function WebPartPicker({ value, onChange }: WebPartPickerProps) {
  return (
    <Dropdown
      size="small"
      value={value ? webPartOptions.find(o => o.value === value)?.label || '' : 'None (empty)'}
      selectedOptions={[value || '']}
      onOptionSelect={(_, data) => {
        const selectedValue = data.optionValue as WebPartType | '';
        onChange(selectedValue === '' ? null : selectedValue);
      }}
    >
      {webPartOptions.map((option) => (
        <Option key={option.value} value={option.value} text={option.label}>
          <span style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
            {option.icon}
            {option.label}
          </span>
        </Option>
      ))}
    </Dropdown>
  );
}
