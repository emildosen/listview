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
import ViewEditor from '../components/ViewEditor';
import type { ViewDefinition } from '../types/view';

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

function ViewEditorPage() {
  const styles = useStyles();
  const { theme } = useTheme();
  const { viewId } = useParams<{ viewId?: string }>();
  const navigate = useNavigate();
  const { views, saveView } = useSettings();

  const isEditMode = !!viewId;

  // Find view for editing
  const initialView = useMemo((): ViewDefinition | undefined => {
    if (!viewId) return undefined;
    return views.find((v) => v.id === viewId);
  }, [viewId, views]);

  const loading = isEditMode && views.length === 0;

  const handleSave = async (view: ViewDefinition) => {
    await saveView(view);
    navigate('/app/views');
  };

  const handleCancel = () => {
    navigate('/app/views');
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

  if (isEditMode && !initialView) {
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
            <Link to="/app/views" className={styles.breadcrumbLink}>
              Views
            </Link>
          </BreadcrumbItem>
          <BreadcrumbDivider />
          <BreadcrumbItem>
            <Text weight="semibold">Not Found</Text>
          </BreadcrumbItem>
        </Breadcrumb>
        <MessageBar intent="error" className={styles.messageBar}>
          <MessageBarBody>View not found</MessageBarBody>
        </MessageBar>
        <div className={styles.backLink}>
          <Button appearance="subtle" icon={<ArrowLeftRegular />} onClick={() => navigate('/app/views')}>
            Back to Views
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
          <Link to="/app/views" className={styles.breadcrumbLink}>
            Views
          </Link>
        </BreadcrumbItem>
        <BreadcrumbDivider />
        <BreadcrumbItem>
          <Text weight="semibold">
            {isEditMode ? `Edit: ${initialView?.name}` : 'Create View'}
          </Text>
        </BreadcrumbItem>
      </Breadcrumb>

      <div className={styles.content}>
        <div className={styles.header}>
          <Title2 as="h1" className={styles.title}>
            {isEditMode ? 'Edit View' : 'Create View'}
          </Title2>
          <Text className={styles.description}>
            {isEditMode
              ? 'Modify the view configuration below.'
              : 'Configure a new view to combine and display data from multiple lists.'}
          </Text>
        </div>

        <ViewEditor
          initialView={initialView}
          onSave={handleSave}
          onCancel={handleCancel}
        />
      </div>
    </div>
  );
}

export default ViewEditorPage;
