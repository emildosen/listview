import { useMemo } from 'react';
import { useParams, useNavigate, Link } from 'react-router-dom';
import {
  makeStyles,
  mergeClasses,
  tokens,
  Text,
  Title2,
  Button,
  Spinner,
  Breadcrumb,
  BreadcrumbItem,
  BreadcrumbDivider,
  MessageBar,
  MessageBarBody,
} from '@fluentui/react-components';
import { ArrowLeftRegular } from '@fluentui/react-icons';
import { useSettings } from '../contexts/SettingsContext';
import { useTheme } from '../contexts/ThemeContext';
import PageEditor from '../components/PageEditor/PageEditor';
import type { PageDefinition } from '../types/page';

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
    marginBottom: '24px',
  },
  title: {
    marginBottom: '4px',
  },
  description: {
    display: 'block',
    color: tokens.colorNeutralForeground2,
  },
  // Azure style: sharp edges, subtle shadow
  card: {
    backgroundColor: tokens.colorNeutralBackground1,
    borderRadius: '2px',
    border: '1px solid transparent',
    borderImage: 'linear-gradient(135deg, rgba(0,0,0,0.05) 0%, rgba(0,0,0,0.15) 100%) 1',
    boxShadow: '0 1px 2px rgba(0, 0, 0, 0.08), 0 2px 4px rgba(0, 0, 0, 0.04)',
  },
  cardDark: {
    backgroundColor: '#1a1a1a',
    borderImage: 'none',
    border: '1px solid #333333',
  },
  loadingContainer: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    padding: '48px',
  },
  backLink: {
    marginTop: '16px',
  },
  messageBar: {
    marginBottom: '16px',
  },
});

function PageEditorPage() {
  const styles = useStyles();
  const { theme } = useTheme();
  const { pageId } = useParams<{ pageId?: string }>();
  const navigate = useNavigate();
  const { pages, savePage } = useSettings();

  const isEditMode = !!pageId;

  // Find page for editing
  const initialPage = useMemo((): PageDefinition | undefined => {
    if (!pageId) return undefined;
    return pages.find((p) => p.id === pageId);
  }, [pageId, pages]);

  const loading = isEditMode && pages.length === 0;

  const handleSave = async (page: PageDefinition) => {
    const savedPage = await savePage(page);
    // Navigate to the new/updated page
    navigate(`/app/pages/${savedPage.id}`);
    return savedPage;
  };

  const handleCancel = () => {
    navigate('/app/pages');
  };

  if (loading) {
    return (
      <div className={styles.container}>
        <div className={mergeClasses(styles.card, theme === 'dark' && styles.cardDark)}>
          <div className={styles.loadingContainer}>
            <Spinner size="large" />
          </div>
        </div>
      </div>
    );
  }

  if (isEditMode && !initialPage) {
    return (
      <div className={styles.container}>
        <Breadcrumb className={styles.breadcrumb}>
          <BreadcrumbItem>
            <Link to="/app" className={styles.breadcrumbLink}>
              Home
            </Link>
          </BreadcrumbItem>
          <BreadcrumbDivider />
          <BreadcrumbItem>
            <Link to="/app/pages" className={styles.breadcrumbLink}>
              Pages
            </Link>
          </BreadcrumbItem>
          <BreadcrumbDivider />
          <BreadcrumbItem>
            <Text weight="semibold">Not Found</Text>
          </BreadcrumbItem>
        </Breadcrumb>
        <MessageBar intent="error" className={styles.messageBar}>
          <MessageBarBody>Page not found</MessageBarBody>
        </MessageBar>
        <div className={styles.backLink}>
          <Button appearance="subtle" icon={<ArrowLeftRegular />} onClick={() => navigate('/app/pages')}>
            Back to Pages
          </Button>
        </div>
      </div>
    );
  }

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
          <Link to="/app/pages" className={styles.breadcrumbLink}>
            Pages
          </Link>
        </BreadcrumbItem>
        <BreadcrumbDivider />
        <BreadcrumbItem>
          <Text weight="semibold">
            {isEditMode ? `Edit: ${initialPage?.name}` : 'Create Page'}
          </Text>
        </BreadcrumbItem>
      </Breadcrumb>

      <div className={styles.content}>
        <div className={styles.header}>
          <Title2 as="h1" className={styles.title}>
            {isEditMode ? 'Edit Page' : 'Create Page'}
          </Title2>
          <Text className={styles.description}>
            {isEditMode
              ? 'Modify the page configuration below.'
              : 'Configure a new entity detail page with search, filters, and related data.'}
          </Text>
        </div>

        <PageEditor
          initialPage={initialPage}
          onSave={handleSave}
          onCancel={handleCancel}
        />
      </div>
    </div>
  );
}

export default PageEditorPage;
