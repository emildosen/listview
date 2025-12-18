import { Link, useNavigate } from 'react-router-dom';
import {
  makeStyles,
  tokens,
  Button,
  Card,
  CardHeader,
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
  card: {
    marginBottom: '24px',
  },
  cardHeader: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
  },
  settingsGrid: {
    display: 'flex',
    flexDirection: 'column',
    gap: '16px',
    marginTop: '16px',
  },
  settingRow: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    padding: '12px',
    backgroundColor: tokens.colorNeutralBackground3,
    borderRadius: tokens.borderRadiusMedium,
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
    borderRadius: tokens.borderRadiusMedium,
    textAlign: 'center',
    color: tokens.colorNeutralForeground3,
  },
  backButton: {
    marginTop: '24px',
  },
  cardBody: {
    padding: '16px',
  },
  placeholderText: {
    color: tokens.colorNeutralForeground2,
    display: 'block',
    marginBottom: '16px',
  },
});

function SettingsPage() {
  const styles = useStyles();
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
        <Card className={styles.card}>
          <CardHeader
            header={
              <div className={styles.cardHeader}>
                <FolderRegular fontSize={20} />
                <Text weight="semibold" size={400}>SharePoint Site</Text>
              </div>
            }
          />
          <div className={styles.cardBody}>
            <div className={styles.settingsGrid}>
              <div className={styles.settingRow}>
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

              <div className={styles.settingRow}>
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

              <div className={styles.settingRow}>
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
        </Card>

        {/* App Settings Card */}
        <Card className={styles.card}>
          <CardHeader
            header={
              <div className={styles.cardHeader}>
                <OptionsRegular fontSize={20} />
                <Text weight="semibold" size={400}>Application Settings</Text>
              </div>
            }
          />
          <div className={styles.cardBody}>
            <Text className={styles.placeholderText}>
              App-specific settings will appear here as features are added.
            </Text>
            <div className={styles.emptyPlaceholder}>
              No settings configured yet
            </div>
          </div>
        </Card>

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
