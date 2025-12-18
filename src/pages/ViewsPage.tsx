import { useState } from 'react';
import { Link, useNavigate } from 'react-router-dom';
import {
  makeStyles,
  mergeClasses,
  tokens,
  Button,
  Text,
  Title2,
  Badge,
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
  TableRegular,
  SettingsRegular,
  DeleteRegular,
  ArrowLeftRegular,
} from '@fluentui/react-icons';
import { useSettings } from '../contexts/SettingsContext';
import { useTheme } from '../contexts/ThemeContext';

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
  // Azure style: sharp edges, subtle shadow
  card: {
    backgroundColor: tokens.colorNeutralBackground1,
    borderRadius: '2px',
    border: '1px solid transparent',
    borderImage: 'linear-gradient(135deg, rgba(0,0,0,0.05) 0%, rgba(0,0,0,0.15) 100%) 1',
    boxShadow: '0 1px 2px rgba(0, 0, 0, 0.08), 0 2px 4px rgba(0, 0, 0, 0.04)',
    overflow: 'hidden',
  },
  cardDark: {
    backgroundColor: '#1a1a1a',
    borderImage: 'none',
    border: '1px solid #333333',
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
  viewDescription: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground2,
    whiteSpace: 'nowrap',
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    maxWidth: '300px',
  },
  sourceCount: {
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

function ViewsPage() {
  const styles = useStyles();
  const { theme } = useTheme();
  const { views, removeView } = useSettings();
  const navigate = useNavigate();
  const [deletingId, setDeletingId] = useState<string | null>(null);

  const handleDelete = async (id: string) => {
    if (!confirm('Are you sure you want to delete this view?')) {
      return;
    }

    setDeletingId(id);
    try {
      await removeView(id);
    } catch (error) {
      console.error('Failed to delete view:', error);
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
          <Text weight="semibold">Views</Text>
        </BreadcrumbItem>
      </Breadcrumb>

      <div className={styles.content}>
        <div className={styles.header}>
          <div>
            <Title2 as="h1">Views</Title2>
            <Text className={styles.description}>
              Create custom views to combine and display data from multiple lists.
            </Text>
          </div>
          <Button appearance="primary" icon={<AddRegular />} onClick={() => navigate('/app/views/new')}>
            Create View
          </Button>
        </div>

        {/* Empty State */}
        {views.length === 0 && (
          <div className={mergeClasses(styles.card, theme === 'dark' && styles.cardDark)}>
            <div className={styles.cardBody}>
              <TableRegular fontSize={48} className={styles.emptyIcon} />
              <Text className={styles.emptyText}>No views created yet</Text>
              <Text className={styles.emptySubtext}>
                Create a view to combine data from multiple SharePoint lists
              </Text>
              <Button appearance="primary" size="small" onClick={() => navigate('/app/views/new')}>
                Create your first view
              </Button>
            </div>
          </div>
        )}

        {/* Views List */}
        {views.length > 0 && (
          <div className={mergeClasses(styles.card, theme === 'dark' && styles.cardDark)}>
            <Table>
              <TableHeader className={styles.tableHeader}>
                <TableRow>
                  <TableHeaderCell className={styles.tableHeaderCell}>Name</TableHeaderCell>
                  <TableHeaderCell className={styles.tableHeaderCell}>Mode</TableHeaderCell>
                  <TableHeaderCell className={styles.tableHeaderCell}>Sources</TableHeaderCell>
                  <TableHeaderCell className={styles.tableHeaderCell} style={{ width: '96px' }}>Actions</TableHeaderCell>
                </TableRow>
              </TableHeader>
              <TableBody>
                {views.map((view) => (
                  <TableRow
                    key={view.id}
                    className={styles.tableRow}
                    onClick={() => navigate(`/app/views/${view.id}`)}
                  >
                    <TableCell>
                      <TableCellLayout>
                        <div>
                          <Text weight="medium">{view.name}</Text>
                          {view.description && (
                            <div className={styles.viewDescription}>
                              {view.description}
                            </div>
                          )}
                        </div>
                      </TableCellLayout>
                    </TableCell>
                    <TableCell>
                      <Badge
                        appearance="tint"
                        color={view.mode === 'aggregate' ? 'important' : 'brand'}
                        size="small"
                      >
                        {view.mode === 'aggregate' ? 'Aggregate' : 'Union'}
                      </Badge>
                    </TableCell>
                    <TableCell>
                      <Text className={styles.sourceCount}>
                        {view.sources.length} list{view.sources.length !== 1 ? 's' : ''}
                      </Text>
                    </TableCell>
                    <TableCell onClick={(e) => e.stopPropagation()}>
                      <div className={styles.actionsCell}>
                        <Button
                          appearance="subtle"
                          size="small"
                          icon={<SettingsRegular />}
                          title="Edit view"
                          onClick={() => navigate(`/app/views/${view.id}/edit`)}
                        />
                        <Button
                          appearance="subtle"
                          size="small"
                          icon={deletingId === view.id ? <Spinner size="tiny" /> : <DeleteRegular />}
                          className={styles.deleteButton}
                          onClick={() => view.id && handleDelete(view.id)}
                          disabled={deletingId === view.id}
                          title="Delete view"
                        />
                      </div>
                    </TableCell>
                  </TableRow>
                ))}
              </TableBody>
            </Table>
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

export default ViewsPage;
