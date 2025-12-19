import { useState, useEffect, useRef, useMemo, useCallback } from 'react';
import { useMsal } from '@azure/msal-react';
import type { SPFI } from '@pnp/sp';
import {
  makeStyles,
  mergeClasses,
  tokens,
  Input,
  Dropdown,
  Option,
  Text,
  Button,
  Table,
  TableHeader,
  TableRow,
  TableHeaderCell,
  TableBody,
  TableCell,
  Field,
  Link,
} from '@fluentui/react-components';
import { SearchRegular, DismissRegular, DocumentRegular, AddRegular } from '@fluentui/react-icons';
import type { GraphListColumn, GraphListItem } from '../../auth/graphClient';
import type { PageDefinition } from '../../types/page';
import { UnifiedDetailModal } from '../modals/UnifiedDetailModal';
import ItemFormModal from '../modals/ItemFormModal';
import { useTheme } from '../../contexts/ThemeContext';
import { SharePointLink } from '../common/SharePointLink';
import { isSharePointUrl } from '../../auth/graphClient';
import { createListItem, createSPClient } from '../../services/sharepoint';

// URL regex pattern
const URL_REGEX = /(https?:\/\/[^\s]+)/g;

// Extract filename from SharePoint URL
function getSharePointFileName(url: string): string | null {
  try {
    const urlObj = new URL(url);
    if (urlObj.hostname.endsWith('.sharepoint.com')) {
      const path = decodeURIComponent(urlObj.pathname);
      const parts = path.split('/');
      const fileName = parts[parts.length - 1];
      if (fileName && fileName.includes('.')) {
        return fileName;
      }
    }
  } catch {
    // Invalid URL
  }
  return null;
}

// Component to render text with clickable links
function TextWithLinks({ text }: { text: string }) {
  if (!text || typeof text !== 'string') {
    return <>{text}</>;
  }

  // If entire text starts with SharePoint URL, treat whole value as the URL
  // (may contain spaces, so don't rely on regex that breaks at whitespace)
  if (isSharePointUrl(text)) {
    return <SharePointLink url={text} />;
  }

  const urlMatch = text.match(URL_REGEX);
  if (!urlMatch) {
    return <>{text}</>;
  }

  const parts: React.ReactNode[] = [];
  let lastIndex = 0;
  let key = 0;

  text.replace(URL_REGEX, (match, _p1, offset) => {
    if (offset > lastIndex) {
      parts.push(<span key={key++}>{text.slice(lastIndex, offset)}</span>);
    }

    const spFileName = getSharePointFileName(match);
    const displayText = spFileName || match;

    parts.push(
      <Link
        key={key++}
        href={match}
        target="_blank"
        rel="noopener noreferrer"
        inline
        onClick={(e) => e.stopPropagation()}
        style={{ display: 'inline-flex', alignItems: 'center', gap: '4px' }}
      >
        {spFileName && <DocumentRegular style={{ fontSize: '12px' }} />}
        {displayText}
      </Link>
    );

    lastIndex = offset + match.length;
    return match;
  });

  if (lastIndex < text.length) {
    parts.push(<span key={key++}>{text.slice(lastIndex)}</span>);
  }

  return <>{parts}</>;
}

interface TableViewProps {
  page: PageDefinition;
  columns: GraphListColumn[];
  items: GraphListItem[];
  filters: Record<string, string>;
  searchText: string;
  onFilterChange: (filters: Record<string, string>) => void;
  onSearchChange: (text: string) => void;
  onItemCreated?: () => void;
  onItemUpdated?: () => void;
  onItemDeleted?: () => void;
}

const useStyles = makeStyles({
  container: {
    display: 'flex',
    gap: '24px',
    height: '100%',
  },
  // Filter panel - Azure style: sharp edges, subtle shadow
  filterPanel: {
    width: '280px',
    flexShrink: 0,
  },
  filterPanelInner: {
    backgroundColor: tokens.colorNeutralBackground1,
    borderRadius: '2px',
    border: '1px solid transparent',
    borderImage: 'linear-gradient(135deg, rgba(0,0,0,0.05) 0%, rgba(0,0,0,0.15) 100%) 1',
    boxShadow: '0 1px 2px rgba(0, 0, 0, 0.08), 0 2px 4px rgba(0, 0, 0, 0.04)',
    height: '100%',
    display: 'flex',
    flexDirection: 'column',
    overflow: 'hidden',
  },
  // Search at top - full width
  searchSection: {
    padding: '12px',
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  filterTitleRow: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
    padding: '12px 12px 8px',
  },
  filterTitle: {
    fontWeight: tokens.fontWeightSemibold,
    fontSize: tokens.fontSizeBase200,
    textTransform: 'uppercase',
    letterSpacing: '0.02em',
    color: tokens.colorNeutralForeground2,
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
    gap: '12px',
    padding: '0 12px 20px',
    flex: 1,
    overflow: 'auto',
  },
  searchContainer: {
    position: 'relative',
  },
  resultsCount: {
    padding: '12px',
    paddingTop: '10px',
    borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground3,
  },
  // Table container - Azure style: sharp edges, subtle shadow
  tableContainer: {
    flex: 1,
    display: 'flex',
    flexDirection: 'column',
    overflow: 'hidden',
  },
  tableWrapper: {
    backgroundColor: tokens.colorNeutralBackground1,
    borderRadius: '2px',
    border: '1px solid transparent',
    borderImage: 'linear-gradient(135deg, rgba(0,0,0,0.05) 0%, rgba(0,0,0,0.15) 100%) 1',
    boxShadow: '0 1px 2px rgba(0, 0, 0, 0.08), 0 2px 4px rgba(0, 0, 0, 0.04)',
    flex: 1,
    minHeight: 0,
    display: 'flex',
    flexDirection: 'column',
    overflow: 'hidden',
  },
  emptyTable: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    height: '100%',
    color: tokens.colorNeutralForeground2,
  },
  tableScrollContainer: {
    overflowX: 'auto',
    flex: 1,
  },
  tableHeader: {
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  tableHeaderCell: {
    fontWeight: tokens.fontWeightSemibold,
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground2,
    textTransform: 'uppercase',
    letterSpacing: '0.02em',
  },
  tableRow: {
    cursor: 'pointer',
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    ':hover': {
      backgroundColor: tokens.colorNeutralBackground1Hover,
    },
    ':last-child': {
      borderBottom: 'none',
    },
  },
  // Dark theme styles
  panelDark: {
    backgroundColor: '#1a1a1a',
    borderImage: 'none',
    border: '1px solid #333333',
  },
  // Header with title and actions
  tableCardHeader: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
    marginBottom: '12px',
  },
  tableCardTitle: {
    fontWeight: tokens.fontWeightSemibold,
    fontSize: tokens.fontSizeBase400,
  },
});

function TableView({
  page,
  columns,
  items,
  filters,
  searchText,
  onFilterChange,
  onSearchChange,
  onItemCreated,
  onItemUpdated,
  onItemDeleted,
}: TableViewProps) {
  const styles = useStyles();
  const { theme } = useTheme();
  const { instance, accounts } = useMsal();
  const account = accounts[0];

  const [selectedItem, setSelectedItemState] = useState<GraphListItem | null>(null);

  // Wrap setSelectedItem to add debug logging
  const setSelectedItem = useCallback((item: GraphListItem | null) => {
    console.log('[TableView] setSelectedItem called:', item ? `item ${item.id}` : 'null');
    console.trace('[TableView] setSelectedItem stack trace');
    setSelectedItemState(item);
  }, []);
  const [filterOptions, setFilterOptions] = useState<Record<string, string[]>>({});
  const [createModalOpen, setCreateModalOpen] = useState(false);
  const [saving, setSaving] = useState(false);

  const siteUrl = page.primarySource.siteUrl;

  // Create SP client for the primary source's site
  const primarySpClientRef = useRef<SPFI | null>(null);
  const [primarySpClientReady, setPrimarySpClientReady] = useState(false);

  useEffect(() => {
    if (!account || !siteUrl) {
      setPrimarySpClientReady(false);
      return;
    }

    const initClient = async () => {
      try {
        const client = await createSPClient(instance, account, siteUrl);
        primarySpClientRef.current = client;
        setPrimarySpClientReady(true);
      } catch (err) {
        console.error('Failed to create SP client for primary list:', err);
        setPrimarySpClientReady(false);
      }
    };

    initClient();
  }, [instance, account, siteUrl]);

  // Load filter options
  useState(() => {
    if (!page.searchConfig?.filterColumns.length) {
      return;
    }

    const options: Record<string, string[]> = {};
    for (const filterCol of page.searchConfig.filterColumns) {
      const column = columns.find((c) => c.name === filterCol.internalName);

      if (column?.choice?.choices) {
        options[filterCol.internalName] = column.choice.choices;
      } else if (filterCol.type === 'boolean') {
        options[filterCol.internalName] = ['Yes', 'No'];
      } else {
        const uniqueValues = new Set<string>();
        items.forEach((item) => {
          const value = item.fields[filterCol.internalName];
          if (value !== null && value !== undefined && value !== '') {
            if (typeof value === 'object' && 'LookupValue' in value) {
              uniqueValues.add((value as { LookupValue: string }).LookupValue);
            } else {
              uniqueValues.add(String(value));
            }
          }
        });
        options[filterCol.internalName] = Array.from(uniqueValues).sort();
      }
    }
    setFilterOptions(options);
  });

  const getDisplayValue = (item: GraphListItem, columnName: string): string => {
    const value = item.fields[columnName];
    if (value === null || value === undefined) return '-';
    if (typeof value === 'object') {
      if ('LookupValue' in value) {
        return (value as { LookupValue: string }).LookupValue;
      }
      return JSON.stringify(value);
    }
    if (typeof value === 'boolean') {
      return value ? 'Yes' : 'No';
    }
    if (value instanceof Date) {
      return value.toLocaleDateString();
    }
    if (typeof value === 'string' && /^\d{4}-\d{2}-\d{2}/.test(value)) {
      return new Date(value).toLocaleDateString();
    }
    return String(value);
  };

  // Get initial values for new item form based on current filters
  // Memoized to prevent re-renders from causing focus loss in the form modal
  const createInitialValues = useMemo((): Record<string, unknown> => {
    const initialValues: Record<string, unknown> = {};

    // Map filter values to form fields
    for (const [columnName, filterValue] of Object.entries(filters)) {
      if (filterValue && filterValue !== '') {
        const column = columns.find(c => c.name === columnName);
        if (column) {
          // For boolean columns, convert Yes/No to boolean
          if (column.boolean) {
            initialValues[columnName] = filterValue === 'Yes';
          } else {
            initialValues[columnName] = filterValue;
          }
        }
      }
    }

    return initialValues;
  }, [filters, columns]);

  // Handle creating a new item
  const handleCreate = async (fields: Record<string, unknown>) => {
    if (!primarySpClientRef.current) return;

    setSaving(true);
    try {
      await createListItem(primarySpClientRef.current, page.primarySource.listId, fields);
      setCreateModalOpen(false);
      onItemCreated?.();
    } catch (err) {
      console.error('Failed to create item:', err);
      throw err;
    } finally {
      setSaving(false);
    }
  };

  return (
    <>
      <div className={styles.container}>
        {/* Filter Panel */}
        <div className={styles.filterPanel}>
          <div className={mergeClasses(styles.filterPanelInner, theme === 'dark' && styles.panelDark)}>
            {/* Search at top - full width */}
            <div className={styles.searchSection}>
              <Input
                placeholder="Search..."
                value={searchText}
                onChange={(_e, data) => onSearchChange(data.value)}
                contentBefore={<SearchRegular />}
                contentAfter={
                  searchText ? (
                    <Button
                      appearance="subtle"
                      size="small"
                      icon={<DismissRegular />}
                      onClick={() => onSearchChange('')}
                    />
                  ) : undefined
                }
                style={{ width: '100%' }}
              />
            </div>

            {/* Filter section */}
            <div className={styles.filterTitleRow}>
              <Text className={styles.filterTitle}>Filters</Text>
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
                  <Field key={filterCol.internalName} label={filterCol.displayName} size="small">
                    <Dropdown
                      value={filters[filterCol.internalName] || 'All'}
                      selectedOptions={filters[filterCol.internalName] ? [filters[filterCol.internalName]] : []}
                      onOptionSelect={(_e, data) =>
                        onFilterChange({
                          ...filters,
                          [filterCol.internalName]: data.optionValue as string || '',
                        })
                      }
                      size="small"
                    >
                      <Option value="">All</Option>
                      {(filterOptions[filterCol.internalName] || []).map((option) => (
                        <Option key={option} value={option}>
                          {option}
                        </Option>
                      ))}
                    </Dropdown>
                  </Field>
                ))}
              </div>
            )}

            {/* Results count */}
            <Text className={styles.resultsCount}>
              {items.length} result{items.length !== 1 ? 's' : ''}
            </Text>
          </div>
        </div>

        {/* Table */}
        <div className={styles.tableContainer}>
          {/* Header with Title and Add Button */}
          <div className={styles.tableCardHeader}>
            <Text className={styles.tableCardTitle}>{page.primarySource.listName}</Text>
            <Button
              appearance="primary"
              size="small"
              icon={<AddRegular />}
              onClick={() => setCreateModalOpen(true)}
              disabled={!primarySpClientReady}
            >
              New
            </Button>
          </div>

          <div className={mergeClasses(styles.tableWrapper, theme === 'dark' && styles.panelDark)}>
            {items.length === 0 ? (
              <div className={styles.emptyTable}>
                <Text>No items found</Text>
              </div>
            ) : (
              <div className={styles.tableScrollContainer}>
                <Table>
                  <TableHeader className={styles.tableHeader}>
                    <TableRow>
                      {(page.searchConfig?.tableColumns || page.displayColumns).map((col) => (
                        <TableHeaderCell key={col.internalName} className={styles.tableHeaderCell}>
                          {col.displayName}
                        </TableHeaderCell>
                      ))}
                    </TableRow>
                  </TableHeader>
                  <TableBody>
                    {items.map((item) => (
                      <TableRow
                        key={item.id}
                        className={styles.tableRow}
                        onClick={() => setSelectedItem(item)}
                      >
                        {(page.searchConfig?.tableColumns || page.displayColumns).map((col) => (
                          <TableCell key={col.internalName}>
                            <TextWithLinks text={getDisplayValue(item, col.internalName)} />
                          </TableCell>
                        ))}
                      </TableRow>
                    ))}
                  </TableBody>
                </Table>
              </div>
            )}
          </div>
        </div>
      </div>

      {/* Detail Modal */}
      {selectedItem && (
        <UnifiedDetailModal
          listId={page.primarySource.listId}
          listName={page.primarySource.listName}
          siteId={page.primarySource.siteId}
          siteUrl={siteUrl}
          columns={columns}
          item={selectedItem}
          page={page}
          onClose={() => setSelectedItem(null)}
          onItemUpdated={onItemUpdated}
          onItemDeleted={() => {
            setSelectedItem(null);
            onItemDeleted?.();
          }}
        />
      )}

      {/* Create Item Modal */}
      {createModalOpen && (
        <ItemFormModal
          mode="create"
          siteId={page.primarySource.siteId}
          listId={page.primarySource.listId}
          initialValues={createInitialValues}
          saving={saving}
          onSave={handleCreate}
          onClose={() => setCreateModalOpen(false)}
        />
      )}
    </>
  );
}

export default TableView;
