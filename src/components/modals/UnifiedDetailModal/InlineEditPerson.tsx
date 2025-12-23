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
  const isClearing = useRef(false); // Track when user is clearing to select a new person

  // Normalize value to array for easier handling
  const selectedValues: PersonOrGroupOption[] = useMemo(() => {
    return value ? (Array.isArray(value) ? value : [value]) : [];
  }, [value]);

  // Load initial users on mount
  useEffect(() => {
    inputRef.current?.focus();

    // Fetch initial users to show before user types
    if (account) {
      setIsSearching(true);
      searchPeople(instance, account, '', chooseFromType, 10)
        .then(results => {
          // Filter out already selected items
          const selectedIds = new Set(selectedValues.map(v => v.id));
          setSearchResults(results.filter(r => !selectedIds.has(r.id)));
        })
        .catch(console.error)
        .finally(() => setIsSearching(false));
    }
  }, []); // eslint-disable-line react-hooks/exhaustive-deps

  // Debounced search
  const performSearch = useCallback(async (query: string) => {
    if (!account) {
      setSearchResults([]);
      setIsSearching(false);
      return;
    }

    setIsSearching(true);
    try {
      // Empty query fetches initial users, non-empty query searches
      const results = await searchPeople(instance, account, query, chooseFromType, 10);
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
      // Debounce search when typing
      searchDebounceRef.current = setTimeout(() => {
        performSearch(searchQuery);
      }, 300);
    } else {
      // When cleared, reload initial users immediately
      performSearch('');
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
    } else if (e.key === 'Enter') {
      e.preventDefault();
      onCommit();
    }
  };

  const handleBlur = () => {
    // Commit on blur - this handles:
    // - Multi-select: commit current selections
    // - Single-select with no value: commit null (clearing the field)
    onCommit();
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
        onBlur={() => {
          // Don't commit if user is clearing to select a new person
          if (isClearing.current) {
            isClearing.current = false;
            return;
          }
          // Commit current value when clicking outside
          onCommit();
        }}
        contentAfter={
          <DismissRegular
            style={{ cursor: 'pointer' }}
            onClick={(e) => {
              e.stopPropagation();
              isClearing.current = true; // Mark that we're clearing to select new
              onChange(null);
              // Don't commit - let user select a new person or click outside to save empty
            }}
          />
        }
        contentBefore={
          <PersonRegular />
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
              icon={<PersonRegular />}
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
