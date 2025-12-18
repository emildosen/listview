import { useState, useEffect } from 'react';
import { useMsal } from '@azure/msal-react';
import {
  makeStyles,
  mergeClasses,
  tokens,
  Input,
  Dropdown,
  Option,
  Text,
  Button,
} from '@fluentui/react-components';
import { SearchRegular, DismissRegular } from '@fluentui/react-icons';
import type { GraphListColumn, GraphListItem } from '../../auth/graphClient';
import type { PageDefinition } from '../../types/page';
import { useTheme } from '../../contexts/ThemeContext';

interface SearchPanelProps {
  page: PageDefinition;
  columns: GraphListColumn[];
  items: GraphListItem[];
  filters: Record<string, string>;
  searchText: string;
  selectedItemId: string | null;
  onFilterChange: (filters: Record<string, string>) => void;
  onSearchChange: (text: string) => void;
  onSelectItem: (itemId: string | null) => void;
}

const useStyles = makeStyles({
  container: {
    height: '100%',
    display: 'flex',
    flexDirection: 'column',
    backgroundColor: tokens.colorNeutralBackground2,
    borderRadius: tokens.borderRadiusMedium,
  },
  containerDark: {
    backgroundColor: '#1a1a1a',
  },
  filtersSection: {
    padding: '16px',
    borderBottom: `1px solid ${tokens.colorNeutralStroke1}`,
  },
  filtersTitleRow: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
    marginBottom: '12px',
  },
  filtersTitle: {
    fontWeight: tokens.fontWeightMedium,
  },
  clearFiltersLink: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorBrandForeground1,
    cursor: 'pointer',
    ':hover': {
      textDecoration: 'underline',
    },
  },
  filterList: {
    display: 'flex',
    flexDirection: 'column',
    gap: '8px',
    marginBottom: '12px',
  },
  searchContainer: {
    position: 'relative',
    marginBottom: '8px',
  },
  searchIcon: {
    position: 'absolute',
    left: '8px',
    top: '50%',
    transform: 'translateY(-50%)',
    color: tokens.colorNeutralForeground3,
    pointerEvents: 'none',
  },
  searchInput: {
    paddingLeft: '32px',
  },
  clearButton: {
    position: 'absolute',
    right: '4px',
    top: '50%',
    transform: 'translateY(-50%)',
  },
  resultsSection: {
    flex: 1,
    overflowY: 'auto',
    padding: '8px',
  },
  resultsCount: {
    fontSize: tokens.fontSizeBase100,
    color: tokens.colorNeutralForeground2,
    padding: '0 8px',
    marginBottom: '8px',
  },
  emptyResults: {
    textAlign: 'center',
    padding: '32px',
    color: tokens.colorNeutralForeground2,
  },
  itemList: {
    display: 'flex',
    flexDirection: 'column',
    gap: '4px',
  },
  itemButton: {
    width: '100%',
    textAlign: 'left',
    padding: '12px',
    borderRadius: tokens.borderRadiusMedium,
    border: 'none',
    backgroundColor: 'transparent',
    cursor: 'pointer',
    ':hover': {
      backgroundColor: tokens.colorNeutralBackground3,
    },
  },
  itemButtonSelected: {
    backgroundColor: tokens.colorBrandBackground,
    color: tokens.colorNeutralForegroundOnBrand,
    ':hover': {
      backgroundColor: tokens.colorBrandBackgroundHover,
    },
  },
  itemTitle: {
    fontWeight: tokens.fontWeightMedium,
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap',
  },
  itemSubtitle: {
    fontSize: tokens.fontSizeBase200,
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap',
    marginTop: '2px',
  },
  itemSubtitleSelected: {
    opacity: 0.7,
  },
});

function SearchPanel({
  page,
  columns,
  items,
  filters,
  searchText,
  selectedItemId,
  onFilterChange,
  onSearchChange,
  onSelectItem,
}: SearchPanelProps) {
  const styles = useStyles();
  const { theme } = useTheme();
  const { accounts } = useMsal();
  const [filterOptions, setFilterOptions] = useState<Record<string, string[]>>({});
  const [loadingOptions, setLoadingOptions] = useState(false);
  const account = accounts[0];

  // Load filter options for choice and lookup columns
  useEffect(() => {
    const loadFilterOptions = async () => {
      if (!page.searchConfig?.filterColumns.length || !account) {
        return;
      }

      setLoadingOptions(true);
      const options: Record<string, string[]> = {};

      for (const filterCol of page.searchConfig.filterColumns) {
        const column = columns.find((c) => c.name === filterCol.internalName);

        if (column?.choice?.choices) {
          // Choice column - use predefined choices
          options[filterCol.internalName] = column.choice.choices;
        } else if (column?.lookup) {
          // Lookup column - fetch unique values from items
          const uniqueValues = new Set<string>();
          items.forEach((item) => {
            const value = item.fields[filterCol.internalName];
            if (value && typeof value === 'object' && 'LookupValue' in value) {
              uniqueValues.add((value as { LookupValue: string }).LookupValue);
            } else if (value && typeof value === 'string') {
              uniqueValues.add(value);
            }
          });
          options[filterCol.internalName] = Array.from(uniqueValues).sort();
        } else if (filterCol.type === 'boolean') {
          options[filterCol.internalName] = ['Yes', 'No'];
        } else {
          // Extract unique values from current items
          const uniqueValues = new Set<string>();
          items.forEach((item) => {
            const value = item.fields[filterCol.internalName];
            if (value !== null && value !== undefined && value !== '') {
              uniqueValues.add(String(value));
            }
          });
          options[filterCol.internalName] = Array.from(uniqueValues).sort();
        }
      }

      setFilterOptions(options);
      setLoadingOptions(false);
    };

    loadFilterOptions();
  }, [page.searchConfig?.filterColumns, columns, items, account]);

  const getItemDisplayValue = (item: GraphListItem, columnName: string): string => {
    const value = item.fields[columnName];
    if (value === null || value === undefined) return '';
    if (typeof value === 'object' && 'LookupValue' in value) {
      return (value as { LookupValue: string }).LookupValue;
    }
    if (typeof value === 'boolean') {
      return value ? 'Yes' : 'No';
    }
    return String(value);
  };

  return (
    <div className={mergeClasses(styles.container, theme === 'dark' && styles.containerDark)}>
      {/* Filters Section */}
      <div className={styles.filtersSection}>
        <div className={styles.filtersTitleRow}>
          <Text className={styles.filtersTitle}>Filters</Text>
          {(searchText || Object.values(filters).some(v => v)) && (
            <Text
              className={styles.clearFiltersLink}
              onClick={() => {
                onFilterChange({});
                onSearchChange('');
              }}
            >
              Clear filters
            </Text>
          )}
        </div>

        {/* Dropdown Filters */}
        {page.searchConfig?.filterColumns.length > 0 && (
          <div className={styles.filterList}>
            {page.searchConfig.filterColumns.map((filterCol) => (
              <Dropdown
                key={filterCol.internalName}
                value={filters[filterCol.internalName] ? filterCol.displayName + ': ' + filters[filterCol.internalName] : filterCol.displayName + ': All'}
                selectedOptions={filters[filterCol.internalName] ? [filters[filterCol.internalName]] : []}
                onOptionSelect={(_e, data) =>
                  onFilterChange({
                    ...filters,
                    [filterCol.internalName]: data.optionValue as string || '',
                  })
                }
                disabled={loadingOptions}
                size="small"
              >
                <Option text={`${filterCol.displayName}: All`} value="">{filterCol.displayName}: All</Option>
                {(filterOptions[filterCol.internalName] || []).map((option) => (
                  <Option key={option} value={option}>
                    {option}
                  </Option>
                ))}
              </Dropdown>
            ))}
          </div>
        )}

        {/* Text Search */}
        <div className={styles.searchContainer}>
          <SearchRegular className={styles.searchIcon} />
          <Input
            className={styles.searchInput}
            placeholder="Search..."
            value={searchText}
            onChange={(_e, data) => onSearchChange(data.value)}
            size="small"
            style={{ paddingLeft: '32px' }}
          />
          {searchText && (
            <Button
              className={styles.clearButton}
              appearance="subtle"
              size="small"
              icon={<DismissRegular />}
              onClick={() => onSearchChange('')}
            />
          )}
        </div>
      </div>

      {/* Results Section */}
      <div className={styles.resultsSection}>
        <Text className={styles.resultsCount}>
          {items.length} result{items.length !== 1 ? 's' : ''}
        </Text>

        {items.length === 0 ? (
          <div className={styles.emptyResults}>
            <Text>No items found</Text>
          </div>
        ) : (
          <div className={styles.itemList}>
            {items.map((item) => {
              const titleValue = getItemDisplayValue(
                item,
                page.searchConfig?.titleColumn || 'Title'
              );
              const subtitleValues = (page.searchConfig?.subtitleColumns || [])
                .map((col) => getItemDisplayValue(item, col))
                .filter(Boolean);
              const isSelected = selectedItemId === item.id;

              return (
                <button
                  key={item.id}
                  type="button"
                  className={`${styles.itemButton} ${isSelected ? styles.itemButtonSelected : ''}`}
                  onClick={() => onSelectItem(item.id)}
                >
                  <div className={styles.itemTitle}>{titleValue || 'Untitled'}</div>
                  {subtitleValues.length > 0 && (
                    <div
                      className={`${styles.itemSubtitle} ${isSelected ? styles.itemSubtitleSelected : ''}`}
                      style={{ color: isSelected ? undefined : tokens.colorNeutralForeground2 }}
                    >
                      {subtitleValues.join(' • ')}
                    </div>
                  )}
                </button>
              );
            })}
          </div>
        )}
      </div>
    </div>
  );
}

export default SearchPanel;
