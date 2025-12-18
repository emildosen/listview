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
  Dialog,
  DialogSurface,
  DialogBody,
  DialogTitle,
  Field,
  Link,
} from '@fluentui/react-components';
import { SearchRegular, DismissRegular, DocumentRegular } from '@fluentui/react-icons';
import type { GraphListColumn, GraphListItem } from '../../auth/graphClient';
import type { PageDefinition } from '../../types/page';
import DetailPanel from './DetailPanel';

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
}

const useStyles = makeStyles({
  container: {
    display: 'flex',
    gap: '24px',
    height: '100%',
  },
  filterPanel: {
    width: '256px',
    flexShrink: 0,
  },
  filterPanelInner: {
    backgroundColor: tokens.colorNeutralBackground2,
    borderRadius: tokens.borderRadiusMedium,
    padding: '16px',
    height: '100%',
  },
  filterTitle: {
    fontWeight: tokens.fontWeightMedium,
    marginBottom: '16px',
  },
  filterList: {
    display: 'flex',
    flexDirection: 'column',
    gap: '12px',
    marginBottom: '16px',
  },
  searchContainer: {
    position: 'relative',
  },
  resultsCount: {
    marginTop: '16px',
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground2,
  },
  tableContainer: {
    flex: 1,
    overflow: 'auto',
  },
  tableWrapper: {
    backgroundColor: tokens.colorNeutralBackground2,
    borderRadius: tokens.borderRadiusMedium,
    height: '100%',
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
    height: '100%',
  },
  tableRow: {
    cursor: 'pointer',
    ':hover': {
      backgroundColor: tokens.colorNeutralBackground3,
    },
  },
  dialogSurface: {
    maxWidth: '1200px',
    width: '100%',
    height: '90vh',
  },
  dialogBody: {
    height: '100%',
    overflow: 'auto',
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

            {/* Text Search */}
            <Field label="Search" size="small">
              <div className={styles.searchContainer}>
                <Input
                  placeholder="Search..."
                  value={searchText}
                  onChange={(_e, data) => onSearchChange(data.value)}
                  size="small"
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
                />
              </div>
            </Field>

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
                  <TableHeader>
                    <TableRow>
                      {(page.searchConfig?.tableColumns || page.displayColumns).map((col) => (
                        <TableHeaderCell key={col.internalName}>
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
      <Dialog
        open={selectedItem !== null}
        onOpenChange={(_event, data) => {
          if (!data.open) setSelectedItem(null);
        }}
      >
        <DialogSurface className={styles.dialogSurface}>
          <DialogTitle
            action={
              <Button
                appearance="subtle"
                icon={<DismissRegular />}
                onClick={() => setSelectedItem(null)}
              />
            }
          >
            Item Details
          </DialogTitle>
          <DialogBody className={styles.dialogBody}>
            <DetailPanel
              page={page}
              columns={columns}
              item={selectedItem}
              spClient={spClient}
            />
          </DialogBody>
        </DialogSurface>
      </Dialog>
    </>
  );
}

export default TableView;
