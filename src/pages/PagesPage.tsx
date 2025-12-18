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
} from '@fluentui/react-components';
import {
  AddRegular,
  DocumentTextRegular,
  SettingsRegular,
  DeleteRegular,
  ArrowLeftRegular,
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
    maxWidth: '896px',
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
  tableRow: {
    cursor: 'pointer',
    ':hover': {
      backgroundColor: tokens.colorNeutralBackground3,
    },
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
          <Card>
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
          <Card>
            <Table>
              <TableHeader>
                <TableRow>
                  <TableHeaderCell>Name</TableHeaderCell>
                  <TableHeaderCell>Primary List</TableHeaderCell>
                  <TableHeaderCell>Related Sections</TableHeaderCell>
                  <TableHeaderCell style={{ width: '96px' }}>Actions</TableHeaderCell>
                </TableRow>
              </TableHeader>
              <TableBody>
                {pages.map((page) => (
                  <TableRow
                    key={page.id}
                    className={styles.tableRow}
                    onClick={() => navigate(`/app/pages/${page.id}`)}
                  >
                    <TableCell>
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
                    <TableCell>
                      <Text className={styles.sourceInfo}>
                        {page.primarySource?.listName || 'Not configured'}
                      </Text>
                    </TableCell>
                    <TableCell>
                      <Text className={styles.sourceInfo}>
                        {page.relatedSections?.length || 0} section
                        {(page.relatedSections?.length || 0) !== 1 ? 's' : ''}
                      </Text>
                    </TableCell>
                    <TableCell onClick={(e) => e.stopPropagation()}>
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
          </Card>
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
