import { useState } from 'react';
import type { SPFI } from '@pnp/sp';
import {
  makeStyles,
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
import { SearchRegular, DismissRegular, DocumentRegular } from '@fluentui/react-icons';
import type { GraphListColumn, GraphListItem } from '../../auth/graphClient';
import type { PageDefinition } from '../../types/page';
import ItemDetailModal from './ItemDetailModal';

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
  spClient: SPFI | null;
  onPageUpdate: (page: PageDefinition) => Promise<void>;
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
  filterTitle: {
    fontWeight: tokens.fontWeightSemibold,
    fontSize: tokens.fontSizeBase200,
    textTransform: 'uppercase',
    letterSpacing: '0.02em',
    color: tokens.colorNeutralForeground2,
    padding: '12px 12px 8px',
  },
  filterList: {
    display: 'flex',
    flexDirection: 'column',
    gap: '12px',
    padding: '0 12px 12px',
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
    overflow: 'hidden',
  },
  tableWrapper: {
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
});

function TableView({
  page,
  columns,
  items,
  filters,
  searchText,
  onFilterChange,
  onSearchChange,
  spClient,
  onPageUpdate,
}: TableViewProps) {
  const styles = useStyles();
  const [selectedItem, setSelectedItem] = useState<GraphListItem | null>(null);
  const [filterOptions, setFilterOptions] = useState<Record<string, string[]>>({});

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

  return (
    <>
      <div className={styles.container}>
        {/* Filter Panel */}
        <div className={styles.filterPanel}>
          <div className={styles.filterPanelInner}>
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
            <Text className={styles.filterTitle}>Filters</Text>

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
          <div className={styles.tableWrapper}>
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
        <ItemDetailModal
          page={page}
          columns={columns}
          item={selectedItem}
          spClient={spClient}
          onClose={() => setSelectedItem(null)}
          onPageUpdate={onPageUpdate}
        />
      )}
    </>
  );
}

export default TableView;
