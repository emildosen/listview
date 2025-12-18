import { Link, useNavigate } from 'react-router-dom';
import {
  makeStyles,
  mergeClasses,
  tokens,
  Button,
  Text,
  Title2,
  Badge,
  Breadcrumb,
  BreadcrumbItem,
  BreadcrumbDivider,
} from '@fluentui/react-components';
import {
  FolderRegular,
  OpenRegular,
  OptionsRegular,
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
  title: {
    marginBottom: '24px',
  },
  // Azure style: sharp edges, subtle shadow
  card: {
    backgroundColor: tokens.colorNeutralBackground1,
    borderRadius: '2px',
    border: '1px solid transparent',
    borderImage: 'linear-gradient(135deg, rgba(0,0,0,0.05) 0%, rgba(0,0,0,0.15) 100%) 1',
    boxShadow: '0 1px 2px rgba(0, 0, 0, 0.08), 0 2px 4px rgba(0, 0, 0, 0.04)',
    marginBottom: '24px',
  },
  cardDark: {
    backgroundColor: '#1a1a1a',
    borderImage: 'none',
    border: '1px solid #333333',
  },
  cardHeader: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
    padding: '16px',
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  cardTitle: {
    fontWeight: tokens.fontWeightSemibold,
    fontSize: tokens.fontSizeBase200,
    textTransform: 'uppercase',
    letterSpacing: '0.02em',
    color: tokens.colorNeutralForeground2,
  },
  cardBody: {
    padding: '16px',
  },
  settingsGrid: {
    display: 'flex',
    flexDirection: 'column',
    gap: '12px',
  },
  settingRow: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    padding: '12px',
    backgroundColor: tokens.colorNeutralBackground3,
    borderRadius: '2px',
  },
  settingRowDark: {
    backgroundColor: '#252525',
  },
  settingLabel: {
    fontWeight: tokens.fontWeightMedium,
  },
  settingValue: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground2,
    marginTop: '4px',
  },
  cardActions: {
    marginTop: '16px',
  },
  emptyPlaceholder: {
    marginTop: '16px',
    padding: '32px',
    border: `2px dashed ${tokens.colorNeutralStroke1}`,
    borderRadius: '2px',
    textAlign: 'center',
    color: tokens.colorNeutralForeground3,
  },
  backButton: {
    marginTop: '24px',
  },
  placeholderText: {
    color: tokens.colorNeutralForeground2,
    display: 'block',
    marginBottom: '16px',
  },
});

function SettingsPage() {
  const styles = useStyles();
  const { theme } = useTheme();
  const navigate = useNavigate();
  const {
    site,
    sitePath,
    isCustomSite,
    settingsList,
    clearSiteOverride,
    initialize,
  } = useSettings();

  const handleResetSite = () => {
    clearSiteOverride();
    initialize();
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
          <Text weight="semibold">Settings</Text>
        </BreadcrumbItem>
      </Breadcrumb>

      <div className={styles.content}>
        <Title2 as="h1" className={styles.title}>Settings</Title2>

        {/* Site Configuration Card */}
        <div className={mergeClasses(styles.card, theme === 'dark' && styles.cardDark)}>
          <div className={styles.cardHeader}>
            <FolderRegular fontSize={16} />
            <Text className={styles.cardTitle}>SharePoint Site</Text>
          </div>
          <div className={styles.cardBody}>
            <div className={styles.settingsGrid}>
              <div className={mergeClasses(styles.settingRow, theme === 'dark' && styles.settingRowDark)}>
                <div>
                  <Text className={styles.settingLabel}>Connected Site</Text>
                  <Text as="p" className={styles.settingValue}>
                    {site?.displayName}
                  </Text>
                </div>
                <Button
                  appearance="subtle"
                  size="small"
                  icon={<OpenRegular />}
                  iconPosition="after"
                  as="a"
                  href={site?.webUrl}
                  target="_blank"
                  rel="noopener noreferrer"
                >
                  Open in SharePoint
                </Button>
              </div>

              <div className={mergeClasses(styles.settingRow, theme === 'dark' && styles.settingRowDark)}>
                <div>
                  <Text className={styles.settingLabel}>Site Path</Text>
                  <Text as="p" className={styles.settingValue}>
                    <code>{sitePath}</code>
                  </Text>
                </div>
                {isCustomSite ? (
                  <Badge appearance="tint" color="warning">Custom</Badge>
                ) : (
                  <Badge appearance="tint" color="success">Standard</Badge>
                )}
              </div>

              <div className={mergeClasses(styles.settingRow, theme === 'dark' && styles.settingRowDark)}>
                <div>
                  <Text className={styles.settingLabel}>Settings List</Text>
                  <Text as="p" className={styles.settingValue}>
                    {settingsList?.displayName}
                  </Text>
                </div>
                <Badge appearance="ghost">Active</Badge>
              </div>
            </div>

            {isCustomSite && (
              <div className={styles.cardActions}>
                <Button
                  appearance="outline"
                  size="small"
                  onClick={handleResetSite}
                >
                  Reset to Standard Site
                </Button>
              </div>
            )}
          </div>
        </div>

        {/* App Settings Card */}
        <div className={mergeClasses(styles.card, theme === 'dark' && styles.cardDark)}>
          <div className={styles.cardHeader}>
            <OptionsRegular fontSize={16} />
            <Text className={styles.cardTitle}>Application Settings</Text>
          </div>
          <div className={styles.cardBody}>
            <Text className={styles.placeholderText}>
              App-specific settings will appear here as features are added.
            </Text>
            <div className={styles.emptyPlaceholder}>
              No settings configured yet
            </div>
          </div>
        </div>

        {/* Back to app */}
        <div className={styles.backButton}>
          <Button
            appearance="subtle"
            icon={<ArrowLeftRegular />}
            onClick={() => navigate('/app')}
          >
            Back to App
          </Button>
        </div>
      </div>
    </div>
  );
}

export default SettingsPage;
