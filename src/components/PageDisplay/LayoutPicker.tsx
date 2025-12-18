import {
  makeStyles,
  tokens,
  mergeClasses,
  Tooltip,
} from '@fluentui/react-components';
import { useTheme } from '../../contexts/ThemeContext';
import type { SectionLayout } from '../../types/page';

const useStyles = makeStyles({
  container: {
    display: 'flex',
    gap: '8px',
    flexWrap: 'wrap',
  },
  option: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    width: '60px',
    height: '40px',
    padding: '6px',
    borderRadius: tokens.borderRadiusMedium,
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    backgroundColor: tokens.colorNeutralBackground1,
    cursor: 'pointer',
    transition: 'background-color 0.15s ease',
    '&:hover': {
      backgroundColor: tokens.colorNeutralBackground1Hover,
    },
  },
  optionDark: {
    backgroundColor: '#252525',
    border: '1px solid #444',
    '&:hover': {
      backgroundColor: '#333',
    },
  },
  optionSelected: {
    border: `2px solid ${tokens.colorBrandStroke1}`,
    backgroundColor: tokens.colorBrandBackground2,
  },
  optionSelectedDark: {
    backgroundColor: 'rgba(0, 120, 212, 0.2)',
  },
  layoutPreview: {
    display: 'flex',
    gap: '2px',
    width: '100%',
    height: '100%',
  },
  column: {
    backgroundColor: tokens.colorNeutralForeground3,
    borderRadius: '2px',
    height: '100%',
    opacity: 0.4,
  },
  columnSelected: {
    backgroundColor: tokens.colorBrandForeground1,
    opacity: 0.7,
  },
});

interface LayoutOption {
  value: SectionLayout;
  label: string;
  columns: number[];  // Width percentages for each column
}

const layoutOptions: LayoutOption[] = [
  { value: 'one-column', label: 'Full width', columns: [100] },
  { value: 'two-column', label: 'Two columns (50/50)', columns: [50, 50] },
  { value: 'three-column', label: 'Three columns (33/33/33)', columns: [33, 33, 34] },
  { value: 'one-third-left', label: 'One-third left (33/67)', columns: [33, 67] },
  { value: 'one-third-right', label: 'One-third right (67/33)', columns: [67, 33] },
];

interface LayoutPickerProps {
  value: SectionLayout;
  onChange: (layout: SectionLayout) => void;
}

export default function LayoutPicker({ value, onChange }: LayoutPickerProps) {
  const { theme } = useTheme();
  const styles = useStyles();

  return (
    <div className={styles.container}>
      {layoutOptions.map((option) => {
        const isSelected = value === option.value;
        return (
          <Tooltip content={option.label} relationship="label" key={option.value}>
            <button
              type="button"
              className={mergeClasses(
                styles.option,
                theme === 'dark' && styles.optionDark,
                isSelected && styles.optionSelected,
                isSelected && theme === 'dark' && styles.optionSelectedDark
              )}
              onClick={() => onChange(option.value)}
              aria-label={option.label}
              aria-pressed={isSelected}
            >
              <div className={styles.layoutPreview}>
                {option.columns.map((width, idx) => (
                  <div
                    key={idx}
                    className={mergeClasses(
                      styles.column,
                      isSelected && styles.columnSelected
                    )}
                    style={{ flex: width }}
                  />
                ))}
              </div>
            </button>
          </Tooltip>
        );
      })}
    </div>
  );
}
