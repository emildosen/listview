import { useState } from 'react';
import {
  makeStyles,
  tokens,
  Button,
  Card,
  Text,
  Title2,
  Spinner,
  Input,
  MessageBar,
  MessageBarBody,
  Badge,
  Divider,
  Link as FluentLink,
} from '@fluentui/react-components';
import {
  WarningRegular,
  CheckmarkCircleRegular,
  SettingsRegular,
  EditRegular,
} from '@fluentui/react-icons';
import { useSettings } from '../contexts/SettingsContext';
import { DEFAULT_SETTINGS_SITE_PATH } from '../services/sharepoint';

type SetupStep = 'choice' | 'standard-instructions' | 'custom-input';

const useStyles = makeStyles({
  loadingContainer: {
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    justifyContent: 'center',
    minHeight: '400px',
  },
  loadingText: {
    marginTop: '16px',
    color: tokens.colorNeutralForeground2,
  },
  container: {
    maxWidth: '672px',
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
    color: tokens.colorNeutralForeground2,
    marginTop: '8px',
  },
  choiceGrid: {
    display: 'grid',
    gap: '16px',
    '@media (min-width: 768px)': {
      gridTemplateColumns: 'repeat(2, 1fr)',
    },
  },
  choiceCard: {
    padding: '16px',
    cursor: 'pointer',
    ':hover': {
      backgroundColor: tokens.colorNeutralBackground3,
    },
  },
  choiceHeader: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
    marginBottom: '8px',
  },
  choiceDescription: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground2,
    marginBottom: '8px',
  },
  cardBody: {
    padding: '24px',
  },
  cardBodySmall: {
    maxWidth: '512px',
    margin: '0 auto',
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
  stepsList: {
    display: 'flex',
    flexDirection: 'column',
    gap: '16px',
  },
  step: {
    display: 'flex',
    gap: '12px',
  },
  stepNumber: {
    flexShrink: 0,
  },
  stepContent: {
    flex: 1,
  },
  stepDescription: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground2,
  },
  cardActions: {
    display: 'flex',
    justifyContent: 'space-between',
    marginTop: '16px',
  },
  cardActionsEnd: {
    display: 'flex',
    gap: '8px',
  },
  formField: {
    marginTop: '16px',
  },
  inputGroup: {
    display: 'flex',
    width: '100%',
  },
  inputPrefix: {
    display: 'flex',
    alignItems: 'center',
    padding: '0 12px',
    backgroundColor: tokens.colorNeutralBackground3,
    color: tokens.colorNeutralForeground2,
    fontSize: tokens.fontSizeBase200,
    borderRadius: `${tokens.borderRadiusMedium} 0 0 ${tokens.borderRadiusMedium}`,
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    borderRight: 'none',
  },
  inputWithPrefix: {
    flex: 1,
    '& input': {
      borderTopLeftRadius: 0,
      borderBottomLeftRadius: 0,
    },
  },
  hint: {
    fontSize: tokens.fontSizeBase100,
    color: tokens.colorNeutralForeground3,
    marginTop: '4px',
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
});

export function SetupWizard() {
  const styles = useStyles();
  const {
    setupStatus,
    hostname,
    sitePath,
    site,
    error,
    configureSite,
    createList,
    initialize,
  } = useSettings();
  const [step, setStep] = useState<SetupStep>('choice');
  const [customPath, setCustomPath] = useState('');
  const [isChecking, setIsChecking] = useState(false);
  const [checkError, setCheckError] = useState<string | null>(null);

  const sharePointUrl = hostname
    ? `https://${hostname}${DEFAULT_SETTINGS_SITE_PATH}`
    : null;

  const handleUseStandard = () => {
    setStep('standard-instructions');
  };

  const handleUseCustom = () => {
    setStep('custom-input');
    setCheckError(null);
  };

  const handleCheckStandardSite = async () => {
    setIsChecking(true);
    setCheckError(null);

    const success = await configureSite(DEFAULT_SETTINGS_SITE_PATH, false);

    if (!success) {
      setCheckError(
        'Site not found. Please create the site in SharePoint Admin Center first.'
      );
    }

    setIsChecking(false);
  };

  const handleCheckCustomSite = async () => {
    if (!customPath.trim()) {
      setCheckError('Please enter a site path');
      return;
    }

    // Normalize the path
    let normalizedPath = customPath.trim();
    if (!normalizedPath.startsWith('/sites/')) {
      if (normalizedPath.startsWith('/')) {
        normalizedPath = `/sites${normalizedPath}`;
      } else if (normalizedPath.startsWith('sites/')) {
        normalizedPath = `/${normalizedPath}`;
      } else {
        normalizedPath = `/sites/${normalizedPath}`;
      }
    }

    setIsChecking(true);
    setCheckError(null);

    const success = await configureSite(normalizedPath, true);

    if (!success) {
      setCheckError(
        `Site not found at ${normalizedPath}. Please check the path and try again.`
      );
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

  const handleBackToChoice = () => {
    setStep('choice');
    setCheckError(null);
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
        <Text className={styles.loadingText}>Creating settings list...</Text>
      </div>
    );
  }

  // General error state
  if (setupStatus === 'error') {
    return (
      <div className={styles.cardBodySmall}>
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

  // List not found - prompt to create
  if (setupStatus === 'list-not-found') {
    return (
      <div className={styles.cardBodySmall}>
        <Card>
          <div className={styles.cardBody}>
            <div className={styles.siteConnected}>
              <div className={styles.siteIcon}>
                <CheckmarkCircleRegular fontSize={20} />
              </div>
              <div>
                <Text weight="semibold">Site Connected</Text>
                <Text className={styles.siteName}>{site?.displayName}</Text>
              </div>
            </div>

            <Divider style={{ margin: '16px 0' }} />

            <Text weight="semibold">Create Settings List</Text>
            <Text className={styles.choiceDescription} style={{ marginTop: '4px' }}>
              The site exists but doesn't have an LV-Settings list yet.
              Click below to create it.
            </Text>

            <div className={styles.cardActions}>
              <Button appearance="subtle" onClick={handleBackToChoice}>
                Back
              </Button>
              <Button appearance="primary" onClick={handleCreateList}>
                Create Settings List
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
      <div className={styles.cardBodySmall}>
        <Card>
          <div className={styles.cardBody}>
            <div className={styles.siteConnected}>
              <div className={`${styles.siteIcon} ${styles.errorIconSmall}`} style={{ backgroundColor: tokens.colorPaletteRedBackground2 }}>
                <WarningRegular fontSize={20} />
              </div>
              <div>
                <Text weight="semibold">Failed to Create List</Text>
                <Text className={styles.siteName}>
                  Could not create the settings list on {site?.displayName}
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
                <li>Or create the "LV-Settings" list manually in SharePoint</li>
                <li>Then click Retry below</li>
              </ol>
            </div>

            <div className={styles.cardActions}>
              <Button appearance="subtle" onClick={handleBackToChoice}>
                Use Different Site
              </Button>
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

  // Site not found - redirect to appropriate step
  if (setupStatus === 'site-not-found' && step === 'choice') {
    setStep(
      sitePath === DEFAULT_SETTINGS_SITE_PATH
        ? 'standard-instructions'
        : 'custom-input'
    );
  }

  // Main wizard UI
  return (
    <div className={styles.container}>
      <div className={styles.header}>
        <div className={styles.headerIcon}>
          <SettingsRegular fontSize={32} />
        </div>
        <Title2>Set Up ListView</Title2>
        <Text className={styles.headerDescription}>
          ListView needs a SharePoint site to store app settings and data.
        </Text>
      </div>

      {step === 'choice' && (
        <div className={styles.choiceGrid}>
          <Card className={styles.choiceCard} onClick={handleUseStandard}>
            <div className={styles.choiceHeader}>
              <CheckmarkCircleRegular fontSize={20} color={tokens.colorBrandForeground1} />
              <Text weight="semibold">Standard Setup</Text>
            </div>
            <Text className={styles.choiceDescription}>
              Use the default <code>/sites/ListView</code> site. This is
              shared across all users in your organization.
            </Text>
            <Badge appearance="outline" color="brand">Recommended</Badge>
          </Card>

          <Card className={styles.choiceCard} onClick={handleUseCustom}>
            <div className={styles.choiceHeader}>
              <EditRegular fontSize={20} />
              <Text weight="semibold">Custom Site</Text>
            </div>
            <Text className={styles.choiceDescription}>
              Use a different SharePoint site. Each user must configure this
              manually.
            </Text>
            <Badge appearance="ghost">Advanced</Badge>
          </Card>
        </div>
      )}

      {step === 'standard-instructions' && (
        <Card>
          <div className={styles.cardBody}>
            <Text weight="semibold" size={500}>Create the ListView Site</Text>
            <Text className={styles.choiceDescription} style={{ marginTop: '4px' }}>
              A SharePoint site needs to be created at{' '}
              <code>/sites/ListView</code> by a SharePoint admin.
            </Text>

            <Divider style={{ margin: '16px 0' }} />

            <div className={styles.stepsList}>
              <div className={styles.step}>
                <Badge appearance="filled" color="brand" size="large" className={styles.stepNumber}>1</Badge>
                <div className={styles.stepContent}>
                  <Text weight="medium">Go to SharePoint Admin Center</Text>
                  <Text className={styles.stepDescription}>
                    Navigate to{' '}
                    <FluentLink
                      href="https://admin.microsoft.com/sharepoint"
                      target="_blank"
                      rel="noopener noreferrer"
                    >
                      admin.microsoft.com/sharepoint
                    </FluentLink>
                  </Text>
                </div>
              </div>

              <div className={styles.step}>
                <Badge appearance="filled" color="brand" size="large" className={styles.stepNumber}>2</Badge>
                <div className={styles.stepContent}>
                  <Text weight="medium">Create a new Team site</Text>
                  <Text className={styles.stepDescription}>
                    Click "Create" and select "Team site"
                  </Text>
                </div>
              </div>

              <div className={styles.step}>
                <Badge appearance="filled" color="brand" size="large" className={styles.stepNumber}>3</Badge>
                <div className={styles.stepContent}>
                  <Text weight="medium">Set the site address to "ListView"</Text>
                  <Text className={styles.stepDescription}>
                    The URL should be:{' '}
                    {sharePointUrl ? (
                      <code style={{ fontSize: tokens.fontSizeBase100 }}>{sharePointUrl}</code>
                    ) : (
                      <code style={{ fontSize: tokens.fontSizeBase100 }}>
                        https://[tenant].sharepoint.com/sites/ListView
                      </code>
                    )}
                  </Text>
                </div>
              </div>

              <div className={styles.step}>
                <Badge appearance="filled" color="brand" size="large" className={styles.stepNumber}>4</Badge>
                <div className={styles.stepContent}>
                  <Text weight="medium">Grant access to users</Text>
                  <Text className={styles.stepDescription}>
                    Add users who need to use ListView as site members
                  </Text>
                </div>
              </div>
            </div>

            {checkError && (
              <MessageBar intent="warning" style={{ marginTop: '16px' }}>
                <MessageBarBody>{checkError}</MessageBarBody>
              </MessageBar>
            )}

            <div className={styles.cardActions}>
              <Button appearance="subtle" onClick={() => setStep('choice')}>
                Back
              </Button>
              <Button
                appearance="primary"
                onClick={handleCheckStandardSite}
                disabled={isChecking}
                icon={isChecking ? <Spinner size="tiny" /> : undefined}
              >
                {isChecking ? 'Checking...' : "I've created the site"}
              </Button>
            </div>
          </div>
        </Card>
      )}

      {step === 'custom-input' && (
        <Card>
          <div className={styles.cardBody}>
            <Text weight="semibold" size={500}>Use Custom Site</Text>
            <Text className={styles.choiceDescription} style={{ marginTop: '4px' }}>
              Enter the path to an existing SharePoint site. This setting is
              stored locally and each user must configure it.
            </Text>

            <div className={styles.formField}>
              <Text weight="medium" size={200}>Site Path</Text>
              <div className={styles.inputGroup} style={{ marginTop: '4px' }}>
                <span className={styles.inputPrefix}>/sites/</span>
                <Input
                  className={styles.inputWithPrefix}
                  placeholder="MySite"
                  value={customPath}
                  onChange={(_e, data) => setCustomPath(data.value)}
                  onKeyDown={(e) => {
                    if (e.key === 'Enter') {
                      handleCheckCustomSite();
                    }
                  }}
                />
              </div>
              <Text className={styles.hint}>
                Example: "CRM" for /sites/CRM
              </Text>
            </div>

            {checkError && (
              <MessageBar intent="warning" style={{ marginTop: '8px' }}>
                <MessageBarBody>{checkError}</MessageBarBody>
              </MessageBar>
            )}

            <div className={styles.cardActions}>
              <Button
                appearance="subtle"
                onClick={() => {
                  setStep('choice');
                  setCheckError(null);
                }}
              >
                Back
              </Button>
              <Button
                appearance="primary"
                onClick={handleCheckCustomSite}
                disabled={isChecking || !customPath.trim()}
                icon={isChecking ? <Spinner size="tiny" /> : undefined}
              >
                {isChecking ? 'Checking...' : 'Connect to Site'}
              </Button>
            </div>
          </div>
        </Card>
      )}
    </div>
  );
}
