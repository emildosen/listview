import { useState, useEffect } from 'react';
import {
  makeStyles,
  tokens,
  Button,
  Card,
  Text,
  Title2,
  Spinner,
  MessageBar,
  MessageBarBody,
  Divider,
  Dropdown,
  Option,
  Field,
} from '@fluentui/react-components';
import {
  WarningRegular,
  CheckmarkCircleRegular,
  SettingsRegular,
  InfoRegular,
} from '@fluentui/react-icons';
import { useMsal } from '@azure/msal-react';
import { useSettings } from '../contexts/SettingsContext';
import { getAllSites, type GraphSite } from '../auth/graphClient';

const useStyles = makeStyles({
  loadingContainer: {
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    justifyContent: 'center',
    minHeight: '400px',
  },
  loadingText: {
    display: 'block',
    marginTop: '16px',
    color: tokens.colorNeutralForeground2,
  },
  container: {
    maxWidth: '512px',
    margin: '0 auto',
  },
  header: {
    textAlign: 'center',
    marginBottom: '32px',
  },
  headerIcon: {
    width: '64px',
    height: '64px',
    borderRadius: '50%',
    backgroundColor: tokens.colorBrandBackground2,
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    margin: '0 auto 16px',
    color: tokens.colorBrandForeground1,
  },
  headerDescription: {
    display: 'block',
    color: tokens.colorNeutralForeground2,
    marginTop: '8px',
  },
  cardBody: {
    padding: '24px',
  },
  siteConnected: {
    display: 'flex',
    alignItems: 'center',
    gap: '12px',
    marginBottom: '8px',
  },
  siteIcon: {
    width: '40px',
    height: '40px',
    borderRadius: '50%',
    backgroundColor: tokens.colorPaletteGreenBackground2,
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    color: tokens.colorPaletteGreenForeground1,
  },
  errorIcon: {
    width: '64px',
    height: '64px',
    borderRadius: '50%',
    backgroundColor: tokens.colorPaletteRedBackground2,
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    margin: '0 auto',
    color: tokens.colorPaletteRedForeground1,
  },
  errorIconSmall: {
    width: '40px',
    height: '40px',
  },
  siteName: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground2,
  },
  choiceDescription: {
    display: 'block',
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground2,
    marginBottom: '8px',
  },
  cardActions: {
    display: 'flex',
    justifyContent: 'flex-end',
    marginTop: '24px',
  },
  cardActionsEnd: {
    display: 'flex',
    gap: '8px',
  },
  formField: {
    marginTop: '16px',
  },
  infoBox: {
    marginTop: '16px',
    padding: '12px',
    backgroundColor: tokens.colorNeutralBackground3,
    borderRadius: tokens.borderRadiusMedium,
    display: 'flex',
    gap: '8px',
  },
  infoBoxText: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground2,
  },
  instructionsBox: {
    marginTop: '16px',
    padding: '16px',
    backgroundColor: tokens.colorNeutralBackground3,
    borderRadius: tokens.borderRadiusMedium,
  },
  instructionsList: {
    listStyle: 'decimal',
    listStylePosition: 'inside',
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground2,
    display: 'flex',
    flexDirection: 'column',
    gap: '4px',
  },
  loadingRow: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
    padding: '8px 0',
  },
});

export function SetupWizard() {
  const styles = useStyles();
  const { instance, accounts } = useMsal();
  const {
    setupStatus,
    site,
    error,
    configureSite,
    createList,
    initialize,
  } = useSettings();

  const [sites, setSites] = useState<GraphSite[]>([]);
  const [loadingSites, setLoadingSites] = useState(false);
  const [selectedSiteId, setSelectedSiteId] = useState<string | null>(null);
  const [isChecking, setIsChecking] = useState(false);
  const [checkError, setCheckError] = useState<string | null>(null);

  const selectedSite = sites.find((s) => s.id === selectedSiteId);

  // Load sites on mount when in no-site-configured state
  useEffect(() => {
    if (setupStatus === 'no-site-configured' && sites.length === 0 && !loadingSites) {
      loadSites();
    }
  }, [setupStatus, sites.length, loadingSites]);

  const loadSites = async () => {
    const account = accounts[0];
    if (!account) return;

    setLoadingSites(true);
    setCheckError(null);
    try {
      const fetchedSites = await getAllSites(instance, account);
      setSites(fetchedSites);
    } catch (err) {
      console.error('Failed to load sites:', err);
      setCheckError('Failed to load SharePoint sites. Please try again.');
    } finally {
      setLoadingSites(false);
    }
  };

  const handleSiteSelect = (_: unknown, data: { optionValue?: string }) => {
    setSelectedSiteId(data.optionValue || null);
    setCheckError(null);
  };

  const handleConnectToSite = async () => {
    if (!selectedSite) return;

    setIsChecking(true);
    setCheckError(null);

    // Extract site path from webUrl
    const url = new URL(selectedSite.webUrl);
    const sitePath = url.pathname;

    const success = await configureSite(sitePath);

    if (!success) {
      setCheckError('Unable to connect to this site. Please check your permissions.');
    }

    setIsChecking(false);
  };

  const handleCreateList = async () => {
    await createList();
  };

  const handleRetry = () => {
    setCheckError(null);
    initialize();
  };

  // Loading state
  if (setupStatus === 'loading') {
    return (
      <div className={styles.loadingContainer}>
        <Spinner size="large" />
        <Text className={styles.loadingText}>Connecting to SharePoint...</Text>
      </div>
    );
  }

  // Creating list state
  if (setupStatus === 'creating-list') {
    return (
      <div className={styles.loadingContainer}>
        <Spinner size="large" />
        <Text className={styles.loadingText}>Creating system lists...</Text>
      </div>
    );
  }

  // General error state
  if (setupStatus === 'error') {
    return (
      <div className={styles.container}>
        <Card>
          <div className={styles.cardBody} style={{ textAlign: 'center' }}>
            <div className={styles.errorIcon}>
              <WarningRegular fontSize={32} />
            </div>
            <Title2 style={{ marginTop: '16px' }}>Connection Error</Title2>
            <Text className={styles.loadingText}>
              Unable to connect to SharePoint. Please check your permissions and
              try again.
            </Text>
            {error && (
              <Text style={{ color: tokens.colorPaletteRedForeground1, marginTop: '8px', fontFamily: 'monospace', fontSize: tokens.fontSizeBase200 }}>
                {error}
              </Text>
            )}
            <div style={{ marginTop: '16px' }}>
              <Button appearance="primary" onClick={handleRetry}>
                Retry
              </Button>
            </div>
          </div>
        </Card>
      </div>
    );
  }

  // Site not found state
  if (setupStatus === 'site-not-found') {
    return (
      <div className={styles.container}>
        <Card>
          <div className={styles.cardBody}>
            <div className={styles.siteConnected}>
              <div className={`${styles.siteIcon} ${styles.errorIconSmall}`} style={{ backgroundColor: tokens.colorPaletteRedBackground2 }}>
                <WarningRegular fontSize={20} />
              </div>
              <div>
                <Text weight="semibold">Site Not Accessible</Text>
                <Text className={styles.siteName}>
                  The selected site could not be accessed
                </Text>
              </div>
            </div>

            <MessageBar intent="warning" style={{ marginTop: '16px' }}>
              <MessageBarBody>
                The site may have been deleted, or you may not have permission to access it.
                Please select a different site.
              </MessageBarBody>
            </MessageBar>

            <div className={styles.cardActions}>
              <Button appearance="primary" onClick={handleRetry}>
                Select Different Site
              </Button>
            </div>
          </div>
        </Card>
      </div>
    );
  }

  // List not found - prompt to create
  if (setupStatus === 'list-not-found') {
    return (
      <div className={styles.container}>
        <Card>
          <div className={styles.cardBody}>
            <div className={styles.siteConnected}>
              <div className={styles.siteIcon}>
                <CheckmarkCircleRegular fontSize={20} />
              </div>
              <Text weight="semibold">Site Connected</Text>
            </div>

            <Divider style={{ margin: '16px 0' }} />

            <Text weight="semibold">Create System Lists</Text>
            <Text className={styles.choiceDescription} style={{ display: 'block', marginTop: '8px' }}>
              The site exists but doesn't have the required ListView system lists yet.
            </Text>
            <Text className={styles.choiceDescription}>
              Click below to create the LV-Settings and LV-Pages lists.
            </Text>

            <div className={styles.cardActions}>
              <Button appearance="primary" onClick={handleCreateList}>
                Create System Lists
              </Button>
            </div>
          </div>
        </Card>
      </div>
    );
  }

  // List creation failed
  if (setupStatus === 'list-creation-failed') {
    return (
      <div className={styles.container}>
        <Card>
          <div className={styles.cardBody}>
            <div className={styles.siteConnected}>
              <div className={`${styles.siteIcon} ${styles.errorIconSmall}`} style={{ backgroundColor: tokens.colorPaletteRedBackground2 }}>
                <WarningRegular fontSize={20} />
              </div>
              <div>
                <Text weight="semibold">Failed to Create Lists</Text>
                <Text className={styles.siteName}>
                  Could not create the system lists on {site?.displayName}
                </Text>
              </div>
            </div>

            <MessageBar intent="error" style={{ marginTop: '16px' }}>
              <MessageBarBody>
                <Text weight="semibold">Access Denied</Text>
                <br />
                <Text size={200}>
                  {error || 'You may not have permission to create lists on this site.'}
                </Text>
              </MessageBarBody>
            </MessageBar>

            <div className={styles.instructionsBox}>
              <Text weight="medium" size={200}>To fix this:</Text>
              <ol className={styles.instructionsList}>
                <li>Ask a site owner to add you as a site member with Edit permissions</li>
                <li>Or create the "LV-Settings" and "LV-Pages" lists manually in SharePoint</li>
                <li>Then click Retry below</li>
              </ol>
            </div>

            <div className={styles.cardActions}>
              <div className={styles.cardActionsEnd}>
                <Button appearance="outline" onClick={handleCreateList}>
                  Retry
                </Button>
                <Button
                  appearance="primary"
                  as="a"
                  href={site?.webUrl}
                  target="_blank"
                  rel="noopener noreferrer"
                >
                  Open Site
                </Button>
              </div>
            </div>
          </div>
        </Card>
      </div>
    );
  }

  // Main site picker UI (no-site-configured state)
  return (
    <div className={styles.container}>
      <div className={styles.header}>
        <div className={styles.headerIcon}>
          <SettingsRegular fontSize={32} />
        </div>
        <Title2>Set Up ListView</Title2>
        <Text className={styles.headerDescription}>
          Choose a SharePoint site to store app settings and configurations.
        </Text>
      </div>

      <Card>
        <div className={styles.cardBody}>
          <Text weight="semibold" size={500}>Select Primary Site</Text>
          <Text className={styles.choiceDescription} style={{ marginTop: '4px' }}>
            ListView will create system lists on this site to store your settings and custom pages.
          </Text>

          <Field label="SharePoint Site" className={styles.formField}>
            {loadingSites ? (
              <div className={styles.loadingRow}>
                <Spinner size="tiny" />
                <Text size={200}>Loading your sites...</Text>
              </div>
            ) : (
              <Dropdown
                placeholder="Select a site"
                value={selectedSite?.displayName || ''}
                selectedOptions={selectedSiteId ? [selectedSiteId] : []}
                onOptionSelect={handleSiteSelect}
              >
                {sites.map((site) => (
                  <Option key={site.id} value={site.id}>
                    {site.displayName}
                  </Option>
                ))}
              </Dropdown>
            )}
          </Field>

          {selectedSite && (
            <Text size={200} style={{ color: tokens.colorNeutralForeground3, marginTop: '4px', display: 'block' }}>
              {selectedSite.webUrl}
            </Text>
          )}

          {checkError && (
            <MessageBar intent="warning" style={{ marginTop: '16px' }}>
              <MessageBarBody>{checkError}</MessageBarBody>
            </MessageBar>
          )}

          <div className={styles.infoBox}>
            <InfoRegular fontSize={16} style={{ flexShrink: 0, marginTop: '2px' }} />
            <div>
              <Text className={styles.infoBoxText} block>
                <strong>Tip:</strong> Consider creating a dedicated "ListView" site in SharePoint Admin just for app settings.
              </Text>
              <Text className={styles.infoBoxText} block style={{ marginTop: '4px' }}>
                You can still pull data from any site you have access to.
              </Text>
            </div>
          </div>

          <div className={styles.cardActions}>
            <Button
              appearance="primary"
              onClick={handleConnectToSite}
              disabled={!selectedSite || isChecking}
              icon={isChecking ? <Spinner size="tiny" /> : undefined}
            >
              {isChecking ? 'Connecting...' : 'Connect to Site'}
            </Button>
          </div>
        </div>
      </Card>
    </div>
  );
}
