import { useState, useMemo, useEffect, useRef } from 'react';
import {
  makeStyles,
  tokens,
  Input,
  Popover,
  PopoverTrigger,
  PopoverSurface,
  Button,
  Text,
  mergeClasses,
  Spinner,
} from '@fluentui/react-components';
import { SearchRegular, ChevronDownRegular } from '@fluentui/react-icons';
import {
  PAGE_ICONS,
  PAGE_ICON_OPTIONS,
  ALL_ICON_NAMES,
  getIconDisplayName,
  isCuratedIcon,
  loadIconByName,
} from '../../utils/iconMap';
import { useTheme } from '../../contexts/ThemeContext';
import type { FluentIcon } from '@fluentui/react-icons';

const useStyles = makeStyles({
  trigger: {
    width: '100%',
    justifyContent: 'space-between',
  },
  triggerContent: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
  },
  surface: {
    padding: '12px',
    width: '320px',
    maxHeight: '400px',
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
    gap: '6px',
    overflowY: 'auto',
    maxHeight: '280px',
  },
  iconButton: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    width: '44px',
    height: '44px',
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
    border: `1px solid ${tokens.colorNeutralStroke1Hover}`,
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
  exactMatch: {
    padding: '12px',
    backgroundColor: tokens.colorNeutralBackground3,
    borderRadius: tokens.borderRadiusMedium,
    display: 'flex',
    alignItems: 'center',
    gap: '12px',
  },
  exactMatchDark: {
    backgroundColor: '#252525',
  },
  exactMatchIcon: {
    width: '44px',
    height: '44px',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    backgroundColor: tokens.colorNeutralBackground1,
    borderRadius: '4px',
    cursor: 'pointer',
    border: `1px solid ${tokens.colorNeutralStroke1}`,
  },
  exactMatchInfo: {
    flex: 1,
    display: 'flex',
    flexDirection: 'column',
    gap: '2px',
  },
  exactMatchLabel: {
    fontSize: tokens.fontSizeBase100,
    color: tokens.colorNeutralForeground3,
  },
  exactMatchName: {
    fontSize: tokens.fontSizeBase200,
    fontWeight: tokens.fontWeightMedium,
  },
  loadingIcon: {
    width: '44px',
    height: '44px',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
  },
});

interface IconPickerProps {
  value: string;
  onChange: (iconName: string) => void;
}

export function IconPicker({ value, onChange }: IconPickerProps) {
  const styles = useStyles();
  const { theme } = useTheme();
  const [open, setOpen] = useState(false);
  const [search, setSearch] = useState('');
  const [hoveredIcon, setHoveredIcon] = useState<string | null>(null);
  const [dynamicIcon, setDynamicIcon] = useState<FluentIcon | null>(null);
  const [loadingDynamic, setLoadingDynamic] = useState(false);
  const searchInputRef = useRef<HTMLInputElement>(null);

  // Get the current icon component
  const CurrentIcon = PAGE_ICONS[value] || PAGE_ICONS['DocumentRegular'];

  // Filter curated icons based on search
  const filteredIcons = useMemo(() => {
    if (!search.trim()) return PAGE_ICON_OPTIONS;
    const searchLower = search.toLowerCase();
    return PAGE_ICON_OPTIONS.filter((name) =>
      getIconDisplayName(name).toLowerCase().includes(searchLower)
    );
  }, [search]);

  // Check for exact match in full icon set (but not in curated list)
  const exactMatchName = useMemo(() => {
    if (!search.trim()) return null;
    // Check if exact match exists in full list but NOT in curated list
    const exactName = ALL_ICON_NAMES.find(
      (name) => name.toLowerCase() === search.toLowerCase()
    );
    if (exactName && !isCuratedIcon(exactName)) {
      return exactName;
    }
    // Also check with "Regular" suffix
    const withSuffix = search.endsWith('Regular') ? search : `${search}Regular`;
    const exactWithSuffix = ALL_ICON_NAMES.find(
      (name) => name.toLowerCase() === withSuffix.toLowerCase()
    );
    if (exactWithSuffix && !isCuratedIcon(exactWithSuffix)) {
      return exactWithSuffix;
    }
    return null;
  }, [search]);

  // Load dynamic icon when exact match found
  useEffect(() => {
    if (!exactMatchName) {
      setDynamicIcon(null);
      return;
    }

    setLoadingDynamic(true);
    loadIconByName(exactMatchName)
      .then((icon) => {
        setDynamicIcon(icon);
      })
      .finally(() => {
        setLoadingDynamic(false);
      });
  }, [exactMatchName]);

  // Focus search input when popover opens
  useEffect(() => {
    if (open && searchInputRef.current) {
      setTimeout(() => searchInputRef.current?.focus(), 0);
    }
  }, [open]);

  const handleSelect = (iconName: string) => {
    onChange(iconName);
    setOpen(false);
    setSearch('');
  };

  return (
    <Popover open={open} onOpenChange={(_, data) => setOpen(data.open)}>
      <PopoverTrigger disableButtonEnhancement>
        <Button
          appearance="outline"
          className={styles.trigger}
          icon={<ChevronDownRegular />}
          iconPosition="after"
        >
          <span className={styles.triggerContent}>
            <CurrentIcon fontSize={20} />
            <span>{getIconDisplayName(value || 'DocumentRegular')}</span>
          </span>
        </Button>
      </PopoverTrigger>
      <PopoverSurface className={styles.surface}>
        <Input
          ref={searchInputRef}
          className={styles.searchInput}
          placeholder="Search icons..."
          value={search}
          onChange={(_, data) => setSearch(data.value)}
          contentBefore={<SearchRegular />}
        />

        {/* Exact match from full icon set */}
        {exactMatchName && (
          <div
            className={mergeClasses(
              styles.exactMatch,
              theme === 'dark' && styles.exactMatchDark
            )}
          >
            {loadingDynamic ? (
              <div className={styles.loadingIcon}>
                <Spinner size="tiny" />
              </div>
            ) : dynamicIcon ? (
              <div
                className={styles.exactMatchIcon}
                onClick={() => handleSelect(exactMatchName)}
                title={`Select ${getIconDisplayName(exactMatchName)}`}
              >
                {(() => {
                  const DynamicIcon = dynamicIcon;
                  return <DynamicIcon fontSize={24} />;
                })()}
              </div>
            ) : null}
            <div className={styles.exactMatchInfo}>
              <Text className={styles.exactMatchLabel}>Exact match found</Text>
              <Text className={styles.exactMatchName}>
                {getIconDisplayName(exactMatchName)}
              </Text>
            </div>
          </div>
        )}

        {/* Curated icons grid */}
        {filteredIcons.length === 0 && !exactMatchName ? (
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
                    isHovered &&
                      !isSelected &&
                      (theme === 'dark'
                        ? styles.iconButtonDarkHover
                        : styles.iconButtonHover),
                    isHovered && isSelected && styles.iconButtonSelectedHover
                  )}
                  onClick={() => handleSelect(iconName)}
                  onMouseEnter={() => setHoveredIcon(iconName)}
                  onMouseLeave={() => setHoveredIcon(null)}
                  title={getIconDisplayName(iconName)}
                >
                  <IconComponent fontSize={22} />
                </button>
              );
            })}
          </div>
        )}
      </PopoverSurface>
    </Popover>
  );
}
