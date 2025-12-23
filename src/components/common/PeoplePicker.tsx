import { useState, useEffect, useRef, useCallback, useMemo } from 'react';
import { useMsal } from '@azure/msal-react';
import {
  makeStyles,
  tokens,
  Input,
  Spinner,
  Text,
  Tag,
  TagGroup,
  Combobox,
  Option,
} from '@fluentui/react-components';
import {
  PersonRegular,
  DismissRegular,
} from '@fluentui/react-icons';
import { searchPeople, type PersonOrGroupOption } from '../../auth/graphClient';

const useStyles = makeStyles({
  container: {
    display: 'flex',
    flexDirection: 'column',
    gap: '4px',
  },
  selectedTags: {
    display: 'flex',
    flexWrap: 'wrap',
    gap: '4px',
    marginBottom: '4px',
  },
  tag: {
    display: 'flex',
    alignItems: 'center',
    gap: '4px',
  },
  option: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
  },
  optionDetails: {
    display: 'flex',
    flexDirection: 'column',
  },
  optionEmail: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground3,
  },
  noResults: {
    padding: '8px 12px',
    color: tokens.colorNeutralForeground3,
    fontStyle: 'italic',
  },
  loading: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
    padding: '8px 12px',
  },
});

interface PeoplePickerProps {
  value: PersonOrGroupOption | PersonOrGroupOption[] | null;
  onChange: (value: PersonOrGroupOption | PersonOrGroupOption[] | null) => void;
  allowMultiple?: boolean;
  chooseFromType?: 'peopleOnly' | 'peopleAndGroups';
  placeholder?: string;
  disabled?: boolean;
  size?: 'small' | 'medium' | 'large';
}

export function PeoplePicker({
  value,
  onChange,
  allowMultiple = false,
  chooseFromType = 'peopleOnly',
  placeholder = 'Search for people...',
  disabled = false,
  size = 'small',
}: PeoplePickerProps) {
  const styles = useStyles();
  const { instance, accounts } = useMsal();
  const account = accounts[0];

  const [searchQuery, setSearchQuery] = useState('');
  const [searchResults, setSearchResults] = useState<PersonOrGroupOption[]>([]);
  const [isSearching, setIsSearching] = useState(false);
  const [open, setOpen] = useState(false);
  const searchDebounceRef = useRef<ReturnType<typeof setTimeout> | null>(null);

  // Normalize value to array for easier handling
  const selectedValues: PersonOrGroupOption[] = useMemo(() => {
    return value ? (Array.isArray(value) ? value : [value]) : [];
  }, [value]);

  // Debounced search
  const performSearch = useCallback(async (query: string) => {
    if (!query.trim() || !account) {
      setSearchResults([]);
      setIsSearching(false);
      return;
    }

    setIsSearching(true);
    try {
      const results = await searchPeople(instance, account, query, chooseFromType);
      // Filter out already selected items
      const selectedIds = new Set(selectedValues.map(v => v.id));
      const filtered = results.filter(r => !selectedIds.has(r.id));
      setSearchResults(filtered);
    } catch (err) {
      console.error('Search failed:', err);
      setSearchResults([]);
    } finally {
      setIsSearching(false);
    }
  }, [instance, account, chooseFromType, selectedValues]);

  useEffect(() => {
    if (searchDebounceRef.current) {
      clearTimeout(searchDebounceRef.current);
    }

    if (searchQuery.trim().length >= 1) {
      searchDebounceRef.current = setTimeout(() => {
        performSearch(searchQuery);
      }, 300);
    } else {
      setSearchResults([]);
      setIsSearching(false);
    }

    return () => {
      if (searchDebounceRef.current) {
        clearTimeout(searchDebounceRef.current);
      }
    };
  }, [searchQuery, performSearch]);

  const handleSelect = (_event: unknown, data: { optionValue?: string }) => {
    const selectedId = data.optionValue;
    if (!selectedId) return;

    const selectedPerson = searchResults.find(r => r.id === selectedId);
    if (!selectedPerson) return;

    if (allowMultiple) {
      onChange([...selectedValues, selectedPerson]);
    } else {
      onChange(selectedPerson);
    }

    setSearchQuery('');
    setSearchResults([]);
    if (!allowMultiple) {
      setOpen(false);
    }
  };

  const handleRemove = (personId: string) => {
    const newValues = selectedValues.filter(v => v.id !== personId);
    if (allowMultiple) {
      onChange(newValues.length > 0 ? newValues : null);
    } else {
      onChange(null);
    }
  };

  // For single select with existing value, show as read-only input with clear button
  if (!allowMultiple && selectedValues.length > 0) {
    const selected = selectedValues[0];
    return (
      <div className={styles.container}>
        <Input
          value={selected.displayName}
          readOnly
          size={size}
          disabled={disabled}
          contentAfter={
            !disabled ? (
              <DismissRegular
                style={{ cursor: 'pointer' }}
                onClick={() => onChange(null)}
              />
            ) : undefined
          }
          contentBefore={<PersonRegular />}
        />
      </div>
    );
  }

  return (
    <div className={styles.container}>
      {/* Selected tags for multi-select */}
      {allowMultiple && selectedValues.length > 0 && (
        <TagGroup
          onDismiss={(_e, data) => handleRemove(data.value)}
          className={styles.selectedTags}
        >
          {selectedValues.map(person => (
            <Tag
              key={person.id}
              value={person.id}
              dismissible={!disabled}
              icon={<PersonRegular />}
            >
              {person.displayName}
            </Tag>
          ))}
        </TagGroup>
      )}

      {/* Search input */}
      <Combobox
        value={searchQuery}
        open={open}
        onOpenChange={(_e, data) => setOpen(data.open)}
        onInput={(e) => setSearchQuery((e.target as HTMLInputElement).value)}
        onOptionSelect={handleSelect}
        placeholder={placeholder}
        disabled={disabled}
        size={size}
        freeform
      >
        {isSearching ? (
          <div className={styles.loading}>
            <Spinner size="tiny" />
            <Text>Searching...</Text>
          </div>
        ) : searchResults.length > 0 ? (
          searchResults.map(person => (
            <Option key={person.id} value={person.id} text={person.displayName}>
              <div className={styles.option}>
                <PersonRegular />
                <div className={styles.optionDetails}>
                  <Text>{person.displayName}</Text>
                  {person.email && (
                    <Text className={styles.optionEmail}>{person.email}</Text>
                  )}
                </div>
              </div>
            </Option>
          ))
        ) : searchQuery.trim().length > 0 && !isSearching ? (
          <div className={styles.noResults}>No results found</div>
        ) : null}
      </Combobox>
    </div>
  );
}

export default PeoplePicker;
