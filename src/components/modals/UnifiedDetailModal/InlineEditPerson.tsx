import { useEffect, useRef, useState, useCallback, useMemo } from 'react';
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
  PeopleRegular,
  DismissRegular,
} from '@fluentui/react-icons';
import { searchPeople, type PersonOrGroupOption } from '../../../auth/graphClient';

const useStyles = makeStyles({
  container: {
    display: 'flex',
    flexDirection: 'column',
    gap: '4px',
    minWidth: '200px',
  },
  selectedTags: {
    display: 'flex',
    flexWrap: 'wrap',
    gap: '4px',
    marginBottom: '4px',
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
  singleSelectInput: {
    minWidth: '200px',
  },
});

interface InlineEditPersonProps {
  value: PersonOrGroupOption | PersonOrGroupOption[] | null;
  isMultiSelect: boolean;
  chooseFromType: 'peopleOnly' | 'peopleAndGroups';
  onChange: (value: PersonOrGroupOption | PersonOrGroupOption[] | null) => void;
  onCommit: (directValue?: PersonOrGroupOption | PersonOrGroupOption[] | null) => void;
  onCancel: () => void;
}

export function InlineEditPerson({
  value,
  isMultiSelect,
  chooseFromType,
  onChange,
  onCommit,
  onCancel,
}: InlineEditPersonProps) {
  const styles = useStyles();
  const { instance, accounts } = useMsal();
  const account = accounts[0];
  const inputRef = useRef<HTMLInputElement>(null);

  const [searchQuery, setSearchQuery] = useState('');
  const [searchResults, setSearchResults] = useState<PersonOrGroupOption[]>([]);
  const [isSearching, setIsSearching] = useState(false);
  const [open, setOpen] = useState(false);
  const searchDebounceRef = useRef<ReturnType<typeof setTimeout> | null>(null);

  // Normalize value to array for easier handling
  const selectedValues: PersonOrGroupOption[] = useMemo(() => {
    return value ? (Array.isArray(value) ? value : [value]) : [];
  }, [value]);

  // Focus on mount
  useEffect(() => {
    inputRef.current?.focus();
  }, []);

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

    if (isMultiSelect) {
      const newValue = [...selectedValues, selectedPerson];
      onChange(newValue);
    } else {
      onChange(selectedPerson);
      // Commit immediately for single select
      setTimeout(() => onCommit(selectedPerson), 0);
    }

    setSearchQuery('');
    setSearchResults([]);
    if (!isMultiSelect) {
      setOpen(false);
    }
  };

  const handleRemove = (personId: string) => {
    const newValues = selectedValues.filter(v => v.id !== personId);
    if (isMultiSelect) {
      onChange(newValues.length > 0 ? newValues : null);
    } else {
      onChange(null);
    }
  };

  const handleKeyDown = (e: React.KeyboardEvent) => {
    if (e.key === 'Escape') {
      e.preventDefault();
      onCancel();
    }
  };

  const handleBlur = () => {
    // Only commit on blur for multi-select
    if (isMultiSelect) {
      onCommit();
    }
  };

  // For single select with existing value, show as read-only input with clear button
  if (!isMultiSelect && selectedValues.length > 0) {
    const selected = selectedValues[0];
    return (
      <Input
        ref={inputRef}
        value={selected.displayName}
        readOnly
        size="small"
        className={styles.singleSelectInput}
        onKeyDown={handleKeyDown}
        contentAfter={
          <DismissRegular
            style={{ cursor: 'pointer' }}
            onClick={() => {
              onChange(null);
              // Don't commit here - let them search for a new person
            }}
          />
        }
        contentBefore={
          selected.type === 'group' ? <PeopleRegular /> : <PersonRegular />
        }
      />
    );
  }

  return (
    <div className={styles.container}>
      {/* Selected tags for multi-select */}
      {isMultiSelect && selectedValues.length > 0 && (
        <TagGroup
          onDismiss={(_e, data) => handleRemove(data.value)}
          className={styles.selectedTags}
        >
          {selectedValues.map(person => (
            <Tag
              key={person.id}
              value={person.id}
              dismissible
              icon={person.type === 'group' ? <PeopleRegular /> : <PersonRegular />}
            >
              {person.displayName}
            </Tag>
          ))}
        </TagGroup>
      )}

      {/* Search input */}
      <Combobox
        ref={inputRef}
        value={searchQuery}
        open={open}
        onOpenChange={(_e, data) => setOpen(data.open)}
        onInput={(e) => setSearchQuery((e.target as HTMLInputElement).value)}
        onOptionSelect={handleSelect}
        onKeyDown={handleKeyDown}
        onBlur={handleBlur}
        placeholder="Search for people..."
        size="small"
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
                {person.type === 'group' ? <PeopleRegular /> : <PersonRegular />}
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
