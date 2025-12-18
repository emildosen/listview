import { useState } from 'react';
import { Link, useNavigate } from 'react-router-dom';
import {
  makeStyles,
  tokens,
  Button,
  Card,
  Text,
  Title2,
  Spinner,
  Breadcrumb,
  BreadcrumbItem,
  BreadcrumbDivider,
  Table,
  TableHeader,
  TableRow,
  TableHeaderCell,
  TableBody,
  TableCell,
  TableCellLayout,
  Input,
  shorthands,
} from '@fluentui/react-components';
import {
  AddRegular,
  DocumentTextRegular,
  SettingsRegular,
  DeleteRegular,
  ArrowLeftRegular,
  SearchRegular,
  FilterRegular,
} from '@fluentui/react-icons';
import { useSettings } from '../contexts/SettingsContext';

const useStyles = makeStyles({
  container: {
    padding: '32px',
    flex: 1,
  },
  breadcrumb: {
    marginBottom: '24px',
  },
  breadcrumbLink: {
    textDecoration: 'none',
    color: 'inherit',
  },
  content: {
    maxWidth: '1024px',
  },
  header: {
    display: 'flex',
    alignItems: 'flex-start',
    justifyContent: 'space-between',
    marginBottom: '24px',
  },
  description: {
    color: tokens.colorNeutralForeground2,
    marginTop: '4px',
  },
  // Azure style: sharp edges, subtle shadow, gradient border
  tableCard: {
    boxShadow: '0 1px 2px rgba(0, 0, 0, 0.08), 0 2px 4px rgba(0, 0, 0, 0.04)',
    borderRadius: '2px',
    overflow: 'hidden',
    backgroundColor: tokens.colorNeutralBackground1,
    border: '1px solid transparent',
    borderImage: 'linear-gradient(135deg, rgba(0,0,0,0.05) 0%, rgba(0,0,0,0.15) 100%) 1',
  },
  // Search bar at top - full width
  searchBar: {
    padding: '16px',
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  searchInput: {
    width: '100%',
  },
  // Filter toolbar below search
  toolbar: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    padding: '12px 16px',
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  toolbarLeft: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
  },
  toolbarRight: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
    paddingTop: '4px',
  },
  itemCount: {
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase200,
  },
  // Table styles
  tableWrapper: {
    overflow: 'auto',
  },
  table: {
    minWidth: '100%',
  },
  tableHeader: {
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  tableHeaderCell: {
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground2,
    fontSize: tokens.fontSizeBase200,
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
  tableCell: {
    ...shorthands.padding('12px', '16px'),
  },
  // Empty state card - Azure style with gradient border
  emptyCard: {
    boxShadow: '0 1px 2px rgba(0, 0, 0, 0.08), 0 2px 4px rgba(0, 0, 0, 0.04)',
    borderRadius: '2px',
    border: '1px solid transparent',
    borderImage: 'linear-gradient(135deg, rgba(0,0,0,0.05) 0%, rgba(0,0,0,0.15) 100%) 1',
  },
  cardBody: {
    padding: '48px',
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    textAlign: 'center',
  },
  emptyIcon: {
    color: tokens.colorNeutralForeground3,
    marginBottom: '16px',
  },
  emptyText: {
    color: tokens.colorNeutralForeground2,
    marginBottom: '8px',
  },
  emptySubtext: {
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase200,
    marginBottom: '16px',
  },
  footer: {
    marginTop: '32px',
    paddingTop: '24px',
    borderTop: `1px solid ${tokens.colorNeutralStroke1}`,
  },
  pageDescription: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground2,
    whiteSpace: 'nowrap',
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    maxWidth: '300px',
  },
  sourceInfo: {
    color: tokens.colorNeutralForeground2,
  },
  actionsCell: {
    display: 'flex',
    alignItems: 'center',
    gap: '4px',
  },
  deleteButton: {
    color: tokens.colorPaletteRedForeground1,
  },
});

function PagesPage() {
  const styles = useStyles();
  const { pages, removePage } = useSettings();
  const navigate = useNavigate();
  const [deletingId, setDeletingId] = useState<string | null>(null);
  const [searchQuery, setSearchQuery] = useState('');

  const handleDelete = async (id: string) => {
    if (!confirm('Are you sure you want to delete this page?')) {
      return;
    }

    setDeletingId(id);
    try {
      await removePage(id);
    } catch (error) {
      console.error('Failed to delete page:', error);
    } finally {
      setDeletingId(null);
    }
  };

  const filteredPages = pages.filter(
    (page) =>
      page.name.toLowerCase().includes(searchQuery.toLowerCase()) ||
      page.description?.toLowerCase().includes(searchQuery.toLowerCase()) ||
      page.primarySource?.listName?.toLowerCase().includes(searchQuery.toLowerCase())
  );

  return (
    <div className={styles.container}>
      {/* Breadcrumb */}
      <Breadcrumb className={styles.breadcrumb}>
        <BreadcrumbItem>
          <Link to="/app" className={styles.breadcrumbLink}>
            Home
          </Link>
        </BreadcrumbItem>
        <BreadcrumbDivider />
        <BreadcrumbItem>
          <Text weight="semibold">Pages</Text>
        </BreadcrumbItem>
      </Breadcrumb>

      <div className={styles.content}>
        <div className={styles.header}>
          <div>
            <Title2 as="h1">Custom Pages</Title2>
            <Text className={styles.description}>
              Create entity detail pages with search, filters, and related data.
            </Text>
          </div>
          <Button appearance="primary" icon={<AddRegular />} onClick={() => navigate('/app/pages/new')}>
            Create Page
          </Button>
        </div>

        {/* Empty State */}
        {pages.length === 0 && (
          <Card className={styles.emptyCard}>
            <div className={styles.cardBody}>
              <DocumentTextRegular fontSize={48} className={styles.emptyIcon} />
              <Text className={styles.emptyText}>No custom pages created yet</Text>
              <Text className={styles.emptySubtext}>
                Create a page to view entity details with related data
              </Text>
              <Button appearance="primary" size="small" onClick={() => navigate('/app/pages/new')}>
                Create your first page
              </Button>
            </div>
          </Card>
        )}

        {/* Pages List */}
        {pages.length > 0 && (
          <div className={styles.tableCard}>
            {/* Search bar - full width */}
            <div className={styles.searchBar}>
              <Input
                className={styles.searchInput}
                contentBefore={<SearchRegular />}
                placeholder="Search pages..."
                value={searchQuery}
                onChange={(_, data) => setSearchQuery(data.value)}
              />
            </div>

            {/* Filter toolbar */}
            <div className={styles.toolbar}>
              <div className={styles.toolbarLeft}>
                <Button appearance="subtle" icon={<FilterRegular />} size="small">
                  Filter
                </Button>
              </div>
              <div className={styles.toolbarRight}>
                <Text className={styles.itemCount}>
                  {filteredPages.length} {filteredPages.length === 1 ? 'item' : 'items'}
                </Text>
              </div>
            </div>

            {/* Table */}
            <div className={styles.tableWrapper}>
              <Table className={styles.table}>
                <TableHeader className={styles.tableHeader}>
                  <TableRow>
                    <TableHeaderCell className={styles.tableHeaderCell}>Name</TableHeaderCell>
                    <TableHeaderCell className={styles.tableHeaderCell}>Primary List</TableHeaderCell>
                    <TableHeaderCell className={styles.tableHeaderCell}>Related Sections</TableHeaderCell>
                    <TableHeaderCell className={styles.tableHeaderCell} style={{ width: '96px' }}>Actions</TableHeaderCell>
                  </TableRow>
                </TableHeader>
                <TableBody>
                  {filteredPages.map((page) => (
                    <TableRow
                      key={page.id}
                      className={styles.tableRow}
                      onClick={() => navigate(`/app/pages/${page.id}`)}
                    >
                      <TableCell className={styles.tableCell}>
                        <TableCellLayout>
                          <div>
                            <Text weight="medium">{page.name}</Text>
                            {page.description && (
                              <div className={styles.pageDescription}>
                                {page.description}
                              </div>
                            )}
                          </div>
                        </TableCellLayout>
                      </TableCell>
                      <TableCell className={styles.tableCell}>
                        <Text className={styles.sourceInfo}>
                          {page.primarySource?.listName || 'Not configured'}
                        </Text>
                      </TableCell>
                      <TableCell className={styles.tableCell}>
                        <Text className={styles.sourceInfo}>
                          {page.relatedSections?.length || 0} section
                          {(page.relatedSections?.length || 0) !== 1 ? 's' : ''}
                        </Text>
                      </TableCell>
                      <TableCell className={styles.tableCell} onClick={(e) => e.stopPropagation()}>
                        <div className={styles.actionsCell}>
                          <Button
                            appearance="subtle"
                            size="small"
                            icon={<SettingsRegular />}
                            title="Edit page"
                            onClick={() => navigate(`/app/pages/${page.id}/edit`)}
                          />
                          <Button
                            appearance="subtle"
                            size="small"
                            icon={deletingId === page.id ? <Spinner size="tiny" /> : <DeleteRegular />}
                            className={styles.deleteButton}
                            onClick={() => page.id && handleDelete(page.id)}
                            disabled={deletingId === page.id}
                            title="Delete page"
                          />
                        </div>
                      </TableCell>
                    </TableRow>
                  ))}
                </TableBody>
              </Table>
            </div>
          </div>
        )}

        {/* Back Button */}
        <div className={styles.footer}>
          <Button appearance="subtle" icon={<ArrowLeftRegular />} onClick={() => navigate('/app')}>
            Back
          </Button>
        </div>
      </div>
    </div>
  );
}

export default PagesPage;
