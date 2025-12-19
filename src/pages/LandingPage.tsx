import { useMsal, useIsAuthenticated } from '@azure/msal-react';
import { useNavigate } from 'react-router-dom';
import {
  FluentProvider,
  makeStyles,
} from '@fluentui/react-components';
import {
  LockClosedRegular,
  DatabaseRegular,
  CodeRegular,
  ShieldCheckmarkRegular,
  CloudRegular,
  PlugConnectedRegular,
} from '@fluentui/react-icons';
import { greenLightTheme } from '../themes/customTheme';
import { loginRequest } from '../auth/msalConfig';
import Logo from '../components/Logo';

const useStyles = makeStyles({
  root: {
    minHeight: '100vh',
    backgroundColor: '#ffffff',
  },
  hero: {
    minHeight: '100vh',
    display: 'flex',
    flexDirection: 'column',
    justifyContent: 'center',
    padding: '0 24px',
    background: 'linear-gradient(180deg, #f0fdf4 0%, #ffffff 50%)',
    position: 'relative',
    overflow: 'hidden',
  },
  heroContent: {
    maxWidth: '800px',
    margin: '0 auto',
    textAlign: 'center',
    position: 'relative',
    zIndex: 1,
  },
  heroLogo: {
    marginBottom: '24px',
  },
  badge: {
    display: 'inline-flex',
    alignItems: 'center',
    gap: '6px',
    backgroundColor: 'rgba(33, 115, 70, 0.1)',
    color: '#217346',
    padding: '6px 12px',
    borderRadius: '16px',
    fontSize: '13px',
    fontWeight: '500',
    marginBottom: '24px',
  },
  heroTitle: {
    fontSize: '44px',
    fontWeight: '700',
    letterSpacing: '-0.03em',
    color: '#111827',
    marginBottom: '20px',
    lineHeight: 1.15,
    '@media (min-width: 640px)': {
      fontSize: '56px',
    },
    '@media (min-width: 768px)': {
      fontSize: '64px',
    },
  },
  heroHighlight: {
    color: '#217346',
  },
  heroDescription: {
    fontSize: '18px',
    color: '#4b5563',
    marginBottom: '40px',
    maxWidth: '600px',
    margin: '0 auto 40px',
    lineHeight: 1.7,
    '@media (min-width: 768px)': {
      fontSize: '20px',
    },
  },
  heroButtons: {
    display: 'flex',
    flexDirection: 'column',
    gap: '12px',
    justifyContent: 'center',
    alignItems: 'center',
    '@media (min-width: 640px)': {
      flexDirection: 'row',
    },
  },
  primaryButton: {
    backgroundColor: '#217346',
    color: '#ffffff',
    padding: '14px 28px',
    fontSize: '16px',
    fontWeight: '600',
    borderRadius: '8px',
    border: 'none',
    cursor: 'pointer',
    display: 'inline-flex',
    alignItems: 'center',
    justifyContent: 'center',
    gap: '8px',
    transition: 'all 0.2s ease',
    boxShadow: '0 1px 2px rgba(0, 0, 0, 0.05)',
    '&:hover': {
      backgroundColor: '#1a5c38',
      transform: 'translateY(-1px)',
      boxShadow: '0 4px 12px rgba(33, 115, 70, 0.25)',
    },
  },
  secondaryButton: {
    backgroundColor: '#ffffff',
    color: '#374151',
    padding: '14px 28px',
    fontSize: '16px',
    fontWeight: '600',
    borderRadius: '8px',
    border: '1px solid #e5e7eb',
    cursor: 'pointer',
    display: 'inline-flex',
    alignItems: 'center',
    justifyContent: 'center',
    gap: '8px',
    textDecoration: 'none',
    transition: 'all 0.2s ease',
    ':hover': {
      backgroundColor: '#f9fafb',
      border: '1px solid #d1d5db',
    },
  },
  section: {
    padding: '80px 24px',
    backgroundColor: '#ffffff',
  },
  sectionGray: {
    backgroundColor: '#f9fafb',
  },
  sectionContent: {
    maxWidth: '1100px',
    margin: '0 auto',
  },
  sectionTitle: {
    fontSize: '28px',
    fontWeight: '700',
    textAlign: 'center',
    color: '#111827',
    marginBottom: '12px',
    letterSpacing: '-0.02em',
    '@media (min-width: 768px)': {
      fontSize: '36px',
    },
  },
  sectionSubtitle: {
    textAlign: 'center',
    color: '#6b7280',
    marginBottom: '56px',
    maxWidth: '600px',
    margin: '0 auto 56px',
    fontSize: '17px',
    lineHeight: 1.6,
  },
  featuresGrid: {
    display: 'grid',
    gap: '24px',
    '@media (min-width: 768px)': {
      gridTemplateColumns: 'repeat(2, 1fr)',
    },
    '@media (min-width: 1024px)': {
      gridTemplateColumns: 'repeat(3, 1fr)',
    },
  },
  featureCard: {
    padding: '28px',
    borderRadius: '12px',
    backgroundColor: '#ffffff',
    border: '1px solid #e5e7eb',
    transition: 'all 0.2s ease',
    ':hover': {
      border: '1px solid #d1d5db',
      boxShadow: '0 4px 12px rgba(0, 0, 0, 0.05)',
    },
  },
  featureCardGray: {
    backgroundColor: '#f9fafb',
    border: '1px solid #f3f4f6',
    ':hover': {
      border: '1px solid #e5e7eb',
    },
  },
  featureIcon: {
    width: '44px',
    height: '44px',
    borderRadius: '10px',
    backgroundColor: 'rgba(33, 115, 70, 0.1)',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    marginBottom: '16px',
    color: '#217346',
  },
  featureTitle: {
    fontWeight: '600',
    fontSize: '17px',
    color: '#111827',
    marginBottom: '8px',
  },
  featureDescription: {
    color: '#6b7280',
    lineHeight: 1.6,
    fontSize: '15px',
  },
  architectureSection: {
    display: 'grid',
    gap: '48px',
    alignItems: 'center',
    '@media (min-width: 768px)': {
      gridTemplateColumns: '1fr 1fr',
      gap: '64px',
    },
  },
  architectureContent: {
    '@media (min-width: 768px)': {
      paddingRight: '24px',
    },
  },
  architectureTitle: {
    fontSize: '28px',
    fontWeight: '700',
    color: '#111827',
    marginBottom: '16px',
    letterSpacing: '-0.02em',
    '@media (min-width: 768px)': {
      fontSize: '32px',
    },
  },
  architectureDescription: {
    color: '#4b5563',
    fontSize: '17px',
    lineHeight: 1.7,
    marginBottom: '32px',
  },
  architectureList: {
    display: 'flex',
    flexDirection: 'column',
    gap: '16px',
  },
  architectureItem: {
    display: 'flex',
    alignItems: 'flex-start',
    gap: '12px',
  },
  architectureItemIcon: {
    width: '24px',
    height: '24px',
    borderRadius: '50%',
    backgroundColor: 'rgba(33, 115, 70, 0.1)',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    flexShrink: 0,
    marginTop: '2px',
  },
  architectureItemText: {
    color: '#374151',
    fontSize: '15px',
    lineHeight: 1.5,
  },
  checkIcon: {
    color: '#217346',
    fontSize: '14px',
  },
  diagramContainer: {
    backgroundColor: '#f9fafb',
    borderRadius: '12px',
    padding: '32px',
    border: '1px solid #e5e7eb',
  },
  diagram: {
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    gap: '16px',
  },
  diagramBox: {
    backgroundColor: '#ffffff',
    border: '1px solid #e5e7eb',
    borderRadius: '8px',
    padding: '16px 24px',
    textAlign: 'center',
    width: '100%',
    maxWidth: '280px',
  },
  diagramBoxHighlight: {
    backgroundColor: 'rgba(33, 115, 70, 0.05)',
    border: '1px solid #217346',
  },
  diagramLabel: {
    fontSize: '14px',
    fontWeight: '600',
    color: '#111827',
    marginBottom: '4px',
  },
  diagramSublabel: {
    fontSize: '12px',
    color: '#6b7280',
  },
  diagramArrow: {
    color: '#9ca3af',
    fontSize: '20px',
  },
  securityGrid: {
    display: 'grid',
    gap: '16px',
    '@media (min-width: 768px)': {
      gridTemplateColumns: 'repeat(2, 1fr)',
    },
  },
  securityCard: {
    padding: '24px',
    borderRadius: '10px',
    backgroundColor: '#ffffff',
    border: '1px solid #e5e7eb',
  },
  securityTitle: {
    fontWeight: '600',
    fontSize: '15px',
    color: '#111827',
    marginBottom: '6px',
  },
  securityDescription: {
    color: '#6b7280',
    fontSize: '14px',
    lineHeight: 1.5,
  },
  ctaSection: {
    textAlign: 'center',
    maxWidth: '640px',
    margin: '0 auto',
  },
  ctaDescription: {
    color: '#4b5563',
    fontSize: '17px',
    marginBottom: '32px',
    lineHeight: 1.6,
  },
  footer: {
    padding: '40px 24px',
    backgroundColor: '#ffffff',
    borderTop: '1px solid #e5e7eb',
  },
  footerContent: {
    maxWidth: '1100px',
    margin: '0 auto',
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    gap: '20px',
    '@media (min-width: 768px)': {
      flexDirection: 'row',
      justifyContent: 'space-between',
    },
  },
  footerBrand: {
    display: 'flex',
    alignItems: 'center',
    gap: '10px',
    fontWeight: '600',
    fontSize: '18px',
    color: '#111827',
  },
  footerLinks: {
    display: 'flex',
    alignItems: 'center',
    gap: '20px',
  },
  footerLink: {
    display: 'flex',
    alignItems: 'center',
    gap: '6px',
    color: '#6b7280',
    textDecoration: 'none',
    fontSize: '14px',
    transition: 'color 0.2s ease',
    ':hover': {
      color: '#111827',
    },
  },
  footerDivider: {
    color: '#e5e7eb',
  },
  footerLicense: {
    color: '#9ca3af',
    fontSize: '14px',
  },
});

function MicrosoftIcon() {
  return (
    <svg
      xmlns="http://www.w3.org/2000/svg"
      viewBox="0 0 23 23"
      width="18"
      height="18"
      fill="currentColor"
    >
      <path d="M0 0h11v11H0zM12 0h11v11H12zM0 12h11v11H0zM12 12h11v11H12z" />
    </svg>
  );
}

function GitHubIcon() {
  return (
    <svg
      xmlns="http://www.w3.org/2000/svg"
      viewBox="0 0 24 24"
      fill="currentColor"
      width="18"
      height="18"
    >
      <path d="M12 0c-6.626 0-12 5.373-12 12 0 5.302 3.438 9.8 8.207 11.387.599.111.793-.261.793-.577v-2.234c-3.338.726-4.033-1.416-4.033-1.416-.546-1.387-1.333-1.756-1.333-1.756-1.089-.745.083-.729.083-.729 1.205.084 1.839 1.237 1.839 1.237 1.07 1.834 2.807 1.304 3.492.997.107-.775.418-1.305.762-1.604-2.665-.305-5.467-1.334-5.467-5.931 0-1.311.469-2.381 1.236-3.221-.124-.303-.535-1.524.117-3.176 0 0 1.008-.322 3.301 1.23.957-.266 1.983-.399 3.003-.404 1.02.005 2.047.138 3.006.404 2.291-1.552 3.297-1.23 3.297-1.23.653 1.653.242 2.874.118 3.176.77.84 1.235 1.911 1.235 3.221 0 4.609-2.807 5.624-5.479 5.921.43.372.823 1.102.823 2.222v3.293c0 .319.192.694.801.576 4.765-1.589 8.199-6.086 8.199-11.386 0-6.627-5.373-12-12-12z" />
    </svg>
  );
}

function CheckIcon() {
  return (
    <svg width="12" height="12" viewBox="0 0 12 12" fill="none" xmlns="http://www.w3.org/2000/svg">
      <path d="M10 3L4.5 8.5L2 6" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
    </svg>
  );
}

function LandingPage() {
  const styles = useStyles();
  const { instance } = useMsal();
  const isAuthenticated = useIsAuthenticated();
  const navigate = useNavigate();

  const handleSignIn = async () => {
    try {
      await instance.loginPopup(loginRequest);
      navigate('/app');
    } catch (error) {
      console.error('Login failed:', error);
    }
  };

  return (
    <FluentProvider theme={greenLightTheme}>
      <div className={styles.root}>
        {/* Hero Section */}
        <section className={styles.hero}>
          <div className={styles.heroContent}>
            <div className={styles.heroLogo}>
              <Logo size={72} />
            </div>
            <div className={styles.badge}>
              <CodeRegular fontSize={14} />
              Open source
            </div>
            <h1 className={styles.heroTitle}>
              SharePoint lists,{' '}
              <span className={styles.heroHighlight}>your way</span>
            </h1>
            <p className={styles.heroDescription}>
              A flexible web app for managing SharePoint list data. Like Notion,
              build what you need—tracking, dashboards, workflows—without
              being locked into a specific use case.
            </p>
            <div className={styles.heroButtons}>
              <button
                onClick={isAuthenticated ? () => navigate('/app') : handleSignIn}
                className={styles.primaryButton}
              >
                {!isAuthenticated && <MicrosoftIcon />}
                {isAuthenticated ? 'Launch App' : 'Sign in with Microsoft'}
              </button>
              <a
                href="https://github.com/emildosen/listview"
                target="_blank"
                rel="noopener noreferrer"
                className={styles.secondaryButton}
              >
                <GitHubIcon />
                View on GitHub
              </a>
            </div>
          </div>
        </section>

        {/* Features Section */}
        <section className={styles.section}>
          <div className={styles.sectionContent}>
            <h2 className={styles.sectionTitle}>Built for flexibility</h2>
            <p className={styles.sectionSubtitle}>
              Point ListView at any SharePoint list and start working immediately.
              No hardcoding, no custom development.
            </p>
            <div className={styles.featuresGrid}>
              <div className={`${styles.featureCard} ${styles.featureCardGray}`}>
                <div className={styles.featureIcon}>
                  <DatabaseRegular fontSize={22} />
                </div>
                <h3 className={styles.featureTitle}>Schema-driven UI</h3>
                <p className={styles.featureDescription}>
                  Components discover list structure via Graph metadata. The app
                  adapts to whatever lists you point it at.
                </p>
              </div>

              <div className={`${styles.featureCard} ${styles.featureCardGray}`}>
                <div className={styles.featureIcon}>
                  <PlugConnectedRegular fontSize={22} />
                </div>
                <h3 className={styles.featureTitle}>Works with any list</h3>
                <p className={styles.featureDescription}>
                  Connect to existing SharePoint lists across your tenant.
                  No migration or data restructuring needed.
                </p>
              </div>

              <div className={`${styles.featureCard} ${styles.featureCardGray}`}>
                <div className={styles.featureIcon}>
                  <CloudRegular fontSize={22} />
                </div>
                <h3 className={styles.featureTitle}>Pure client-side</h3>
                <p className={styles.featureDescription}>
                  No backend servers. All Graph API calls happen directly from
                  your browser using delegated permissions.
                </p>
              </div>
            </div>
          </div>
        </section>

        {/* Architecture Section */}
        <section className={`${styles.section} ${styles.sectionGray}`}>
          <div className={styles.sectionContent}>
            <div className={styles.architectureSection}>
              <div className={styles.architectureContent}>
                <h2 className={styles.architectureTitle}>Your data stays yours</h2>
                <p className={styles.architectureDescription}>
                  ListView runs entirely in your browser. Authentication goes through
                  Microsoft, and data flows directly between you and SharePoint.
                  There's no middleware, no third-party servers, and no data storage
                  outside your tenant.
                </p>
                <div className={styles.architectureList}>
                  <div className={styles.architectureItem}>
                    <div className={styles.architectureItemIcon}>
                      <span className={styles.checkIcon}><CheckIcon /></span>
                    </div>
                    <span className={styles.architectureItemText}>
                      Uses your existing Microsoft 365 credentials
                    </span>
                  </div>
                  <div className={styles.architectureItem}>
                    <div className={styles.architectureItemIcon}>
                      <span className={styles.checkIcon}><CheckIcon /></span>
                    </div>
                    <span className={styles.architectureItemText}>
                      All data remains in your SharePoint tenant
                    </span>
                  </div>
                  <div className={styles.architectureItem}>
                    <div className={styles.architectureItemIcon}>
                      <span className={styles.checkIcon}><CheckIcon /></span>
                    </div>
                    <span className={styles.architectureItemText}>
                      Settings stored in a dedicated SharePoint site you control
                    </span>
                  </div>
                </div>
              </div>
              <div className={styles.diagramContainer}>
                <div className={styles.diagram}>
                  <div className={styles.diagramBox}>
                    <div className={styles.diagramLabel}>Your Browser</div>
                    <div className={styles.diagramSublabel}>ListView SPA</div>
                  </div>
                  <div className={styles.diagramArrow}>↓</div>
                  <div className={`${styles.diagramBox} ${styles.diagramBoxHighlight}`}>
                    <div className={styles.diagramLabel}>Microsoft Graph API</div>
                    <div className={styles.diagramSublabel}>Delegated permissions</div>
                  </div>
                  <div className={styles.diagramArrow}>↓</div>
                  <div className={styles.diagramBox}>
                    <div className={styles.diagramLabel}>Your SharePoint</div>
                    <div className={styles.diagramSublabel}>Lists & data</div>
                  </div>
                </div>
              </div>
            </div>
          </div>
        </section>

        {/* Security Section */}
        <section className={styles.section}>
          <div className={styles.sectionContent}>
            <h2 className={styles.sectionTitle}>Defense in depth</h2>
            <p className={styles.sectionSubtitle}>
              ListView implements a strict Content Security Policy to restrict what
              code can run and where data can be sent.
            </p>
            <div className={styles.securityGrid}>
              <div className={styles.securityCard}>
                <div className={styles.featureIcon}>
                  <ShieldCheckmarkRegular fontSize={22} />
                </div>
                <h3 className={styles.securityTitle}>XSS mitigation</h3>
                <p className={styles.securityDescription}>
                  CSP blocks unauthorized scripts from executing, even if injection
                  occurs through other vulnerabilities.
                </p>
              </div>
              <div className={styles.securityCard}>
                <div className={styles.featureIcon}>
                  <LockClosedRegular fontSize={22} />
                </div>
                <h3 className={styles.securityTitle}>Data exfiltration prevention</h3>
                <p className={styles.securityDescription}>
                  Outbound connections are restricted to Microsoft services only.
                  Tokens cannot be sent to attacker-controlled servers.
                </p>
              </div>
              <div className={styles.securityCard}>
                <div className={styles.featureIcon}>
                  <CodeRegular fontSize={22} />
                </div>
                <h3 className={styles.securityTitle}>Supply chain defense</h3>
                <p className={styles.securityDescription}>
                  No external CDNs or third-party scripts. All code is bundled and
                  served from the same origin.
                </p>
              </div>
              <div className={styles.securityCard}>
                <div className={styles.featureIcon}>
                  <CloudRegular fontSize={22} />
                </div>
                <h3 className={styles.securityTitle}>Sovereign cloud support</h3>
                <p className={styles.securityDescription}>
                  Self-host for GCC, GCC High, DoD, or China cloud environments
                  with your own Entra ID app registration.
                </p>
              </div>
            </div>
          </div>
        </section>

        {/* CTA Section */}
        <section className={`${styles.section} ${styles.sectionGray}`}>
          <div className={styles.ctaSection}>
            <h2 className={styles.sectionTitle}>Get started</h2>
            <p className={styles.ctaDescription}>
              Free and open source under the MIT license. Works with any commercial
              M365 tenant—your admin may need to grant consent for the app.
            </p>
            <div className={styles.heroButtons}>
              <button
                onClick={isAuthenticated ? () => navigate('/app') : handleSignIn}
                className={styles.primaryButton}
              >
                {!isAuthenticated && <MicrosoftIcon />}
                {isAuthenticated ? 'Launch App' : 'Sign in with Microsoft'}
              </button>
            </div>
          </div>
        </section>

        {/* Footer */}
        <footer className={styles.footer}>
          <div className={styles.footerContent}>
            <div className={styles.footerBrand}>
              <Logo size={22} />
              ListView
            </div>
            <div className={styles.footerLinks}>
              <a
                href="https://github.com/emildosen/listview"
                target="_blank"
                rel="noopener noreferrer"
                className={styles.footerLink}
              >
                <GitHubIcon />
                GitHub
              </a>
              <span className={styles.footerDivider}>|</span>
              <span className={styles.footerLicense}>MIT License</span>
            </div>
          </div>
        </footer>
      </div>
    </FluentProvider>
  );
}

export default LandingPage;
