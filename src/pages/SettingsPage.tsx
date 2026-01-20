import { useState } from 'react';
import { Link, useNavigate } from 'react-router-dom';
import {
  makeStyles,
  mergeClasses,
  tokens,
  Button,
  Text,
  Title2,
  Breadcrumb,
  BreadcrumbItem,
  BreadcrumbDivider,
  Dialog,
  DialogTrigger,
  DialogSurface,
  DialogTitle,
  DialogBody,
  DialogActions,
  DialogContent,
} from '@fluentui/react-components';
import {
  OpenRegular,
  ArrowLeftRegular,
  WarningRegular,
  BuildingSwapRegular,
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
    maxWidth: '600px',
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
    padding: '20px',
  },
  siteInfo: {
    display: 'flex',
    flexDirection: 'column',
    gap: '16px',
  },
  siteName: {
    fontSize: tokens.fontSizeBase400,
    fontWeight: tokens.fontWeightSemibold,
  },
  siteLink: {
    display: 'inline-flex',
    alignItems: 'center',
    gap: '6px',
    color: tokens.colorBrandForeground1,
    textDecoration: 'none',
    fontSize: tokens.fontSizeBase300,
    ':hover': {
      textDecoration: 'underline',
    },
  },
  changeSiteSection: {
    marginTop: '24px',
    paddingTop: '20px',
    borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
    display: 'flex',
    flexDirection: 'column',
    gap: '12px',
  },
  changeSiteHeader: {
    fontWeight: tokens.fontWeightSemibold,
  },
  changeSiteDescription: {
    color: tokens.colorNeutralForeground2,
    fontSize: tokens.fontSizeBase200,
  },
  backButton: {
    marginTop: '24px',
  },
});

function SettingsPage() {
  const styles = useStyles();
  const { theme } = useTheme();
  const navigate = useNavigate();
  const {
    site,
    changePrimarySite,
  } = useSettings();
  const [showChangeDialog, setShowChangeDialog] = useState(false);

  const handleChangePrimarySite = () => {
    setShowChangeDialog(false);
    changePrimarySite();
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

        {/* Primary Site Card */}
        <div className={mergeClasses(styles.card, theme === 'dark' && styles.cardDark)}>
          <div className={styles.cardHeader}>
            <Text className={styles.cardTitle}>Primary Site</Text>
          </div>
          <div className={styles.cardBody}>
            <div className={styles.siteInfo}>
              <Text className={styles.siteName}>{site?.displayName}</Text>
              <a
                href={site?.webUrl}
                target="_blank"
                rel="noopener noreferrer"
                className={styles.siteLink}
              >
                <OpenRegular fontSize={16} />
                Open in SharePoint
              </a>
            </div>

            <div className={styles.changeSiteSection}>
              <Text className={styles.changeSiteHeader}>Change Primary Site</Text>
              <Text className={styles.changeSiteDescription}>
                Switch to a different SharePoint site for storing ListView settings.
              </Text>
              <div>
                <Dialog open={showChangeDialog} onOpenChange={(_, data) => setShowChangeDialog(data.open)}>
                  <DialogTrigger disableButtonEnhancement>
                    <Button
                      appearance="outline"
                      icon={<BuildingSwapRegular />}
                    >
                      Change Primary Site
                    </Button>
                  </DialogTrigger>
                  <DialogSurface>
                    <DialogBody>
                      <DialogTitle>
                        <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                          <WarningRegular color={tokens.colorPaletteYellowForeground1} />
                          Change Primary Site?
                        </div>
                      </DialogTitle>
                      <DialogContent>
                        <Text>
                          Changing your primary site will:
                        </Text>
                        <ul style={{ margin: '12px 0', paddingLeft: '20px' }}>
                          <li>Disconnect from the current site's settings and pages</li>
                          <li>Require you to select a new primary site</li>
                          <li>Not delete any data from the current site</li>
                        </ul>
                        <Text weight="semibold">
                          Settings stored in "{site?.displayName}" will remain there but won't be accessible until you reconnect.
                        </Text>
                      </DialogContent>
                      <DialogActions>
                        <Button appearance="secondary" onClick={() => setShowChangeDialog(false)}>
                          Cancel
                        </Button>
                        <Button appearance="primary" onClick={handleChangePrimarySite}>
                          Change Site
                        </Button>
                      </DialogActions>
                    </DialogBody>
                  </DialogSurface>
                </Dialog>
              </div>
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
