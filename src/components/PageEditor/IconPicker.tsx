import { useState, useMemo } from 'react';
import {
  makeStyles,
  tokens,
  Input,
  mergeClasses,
} from '@fluentui/react-components';
import { SearchRegular } from '@fluentui/react-icons';
import { PAGE_ICONS, PAGE_ICON_OPTIONS, getIconDisplayName } from '../../utils/iconMap';
import { useTheme } from '../../contexts/ThemeContext';

const useStyles = makeStyles({
  container: {
    display: 'flex',
    flexDirection: 'column',
    gap: '12px',
  },
  searchInput: {
    width: '100%',
  },
  grid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(6, 1fr)',
    gap: '8px',
  },
  iconButton: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    width: '48px',
    height: '48px',
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    borderRadius: '4px',
    backgroundColor: tokens.colorNeutralBackground1,
    cursor: 'pointer',
    transitionProperty: 'all',
    transitionDuration: '0.1s',
    transitionTimingFunction: 'ease',
  },
  iconButtonHover: {
    backgroundColor: tokens.colorNeutralBackground1Hover,
  },
  iconButtonDark: {
    backgroundColor: '#252525',
    border: '1px solid #404040',
  },
  iconButtonDarkHover: {
    backgroundColor: '#333333',
  },
  iconButtonSelected: {
    backgroundColor: tokens.colorBrandBackground,
    border: `1px solid ${tokens.colorBrandStroke1}`,
    color: tokens.colorNeutralForegroundOnBrand,
  },
  iconButtonSelectedHover: {
    backgroundColor: tokens.colorBrandBackgroundHover,
  },
  noResults: {
    padding: '16px',
    textAlign: 'center',
    color: tokens.colorNeutralForeground3,
  },
});

interface IconPickerProps {
  value: string;
  onChange: (iconName: string) => void;
}

export function IconPicker({ value, onChange }: IconPickerProps) {
  const styles = useStyles();
  const { theme } = useTheme();
  const [search, setSearch] = useState('');
  const [hoveredIcon, setHoveredIcon] = useState<string | null>(null);

  const filteredIcons = useMemo(() => {
    if (!search.trim()) return PAGE_ICON_OPTIONS;
    const searchLower = search.toLowerCase();
    return PAGE_ICON_OPTIONS.filter((name) =>
      getIconDisplayName(name).toLowerCase().includes(searchLower)
    );
  }, [search]);

  return (
    <div className={styles.container}>
      <Input
        className={styles.searchInput}
        placeholder="Search icons..."
        value={search}
        onChange={(_, data) => setSearch(data.value)}
        contentBefore={<SearchRegular />}
      />
      {filteredIcons.length === 0 ? (
        <div className={styles.noResults}>No icons found</div>
      ) : (
        <div className={styles.grid}>
          {filteredIcons.map((iconName) => {
            const IconComponent = PAGE_ICONS[iconName];
            const isSelected = value === iconName;
            const isHovered = hoveredIcon === iconName;
            return (
              <button
                key={iconName}
                type="button"
                className={mergeClasses(
                  styles.iconButton,
                  theme === 'dark' && styles.iconButtonDark,
                  isSelected && styles.iconButtonSelected,
                  isHovered && !isSelected && (theme === 'dark' ? styles.iconButtonDarkHover : styles.iconButtonHover),
                  isHovered && isSelected && styles.iconButtonSelectedHover
                )}
                onClick={() => onChange(iconName)}
                onMouseEnter={() => setHoveredIcon(iconName)}
                onMouseLeave={() => setHoveredIcon(null)}
                title={getIconDisplayName(iconName)}
              >
                <IconComponent fontSize={24} />
              </button>
            );
          })}
        </div>
      )}
    </div>
  );
}
