import { useMsal, useIsAuthenticated } from '@azure/msal-react';
import { useNavigate } from 'react-router-dom';
import { useEffect } from 'react';
import {
  FluentProvider,
  makeStyles,
} from '@fluentui/react-components';
import {
  LockClosedRegular,
  GridRegular,
  DataBarVerticalRegular,
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
    padding: '0 16px',
    background: 'linear-gradient(to bottom, #f0fdf4, #ffffff)',
  },
  heroContent: {
    maxWidth: '768px',
    margin: '0 auto',
    textAlign: 'center',
  },
  heroLogo: {
    marginBottom: '32px',
  },
  heroTitle: {
    fontSize: '48px',
    fontWeight: '700',
    letterSpacing: '-0.02em',
    color: '#111827',
    marginBottom: '24px',
    lineHeight: 1.1,
    '@media (min-width: 640px)': {
      fontSize: '60px',
    },
    '@media (min-width: 768px)': {
      fontSize: '72px',
    },
  },
  heroHighlight: {
    color: '#217346',
  },
  heroDescription: {
    fontSize: '20px',
    color: '#4b5563',
    marginBottom: '48px',
    maxWidth: '576px',
    margin: '0 auto 48px',
    lineHeight: 1.6,
    '@media (min-width: 768px)': {
      fontSize: '24px',
    },
  },
  heroButtons: {
    display: 'flex',
    flexDirection: 'column',
    gap: '16px',
    justifyContent: 'center',
    '@media (min-width: 640px)': {
      flexDirection: 'row',
    },
  },
  primaryButton: {
    backgroundColor: '#217346',
    color: '#ffffff',
    padding: '16px 32px',
    fontSize: '18px',
    fontWeight: '600',
    borderRadius: '8px',
    border: 'none',
    cursor: 'pointer',
    display: 'inline-flex',
    alignItems: 'center',
    justifyContent: 'center',
    gap: '8px',
    '&:hover': {
      backgroundColor: '#1a5c38',
    },
  },
  secondaryButton: {
    backgroundColor: 'transparent',
    color: '#374151',
    padding: '16px 32px',
    fontSize: '18px',
    fontWeight: '600',
    borderRadius: '8px',
    border: '2px solid #d1d5db',
    cursor: 'pointer',
    display: 'inline-flex',
    alignItems: 'center',
    justifyContent: 'center',
    gap: '8px',
    textDecoration: 'none',
    '&:hover': {
      border: '2px solid #9ca3af',
    },
  },
  section: {
    padding: '96px 16px',
    backgroundColor: '#ffffff',
  },
  sectionGray: {
    backgroundColor: '#f9fafb',
  },
  sectionContent: {
    maxWidth: '1024px',
    margin: '0 auto',
  },
  sectionContentNarrow: {
    maxWidth: '896px',
  },
  sectionTitle: {
    fontSize: '30px',
    fontWeight: '700',
    textAlign: 'center',
    color: '#111827',
    marginBottom: '16px',
    '@media (min-width: 768px)': {
      fontSize: '36px',
    },
  },
  sectionSubtitle: {
    textAlign: 'center',
    color: '#6b7280',
    marginBottom: '64px',
    maxWidth: '576px',
    margin: '0 auto 64px',
    fontSize: '18px',
  },
  featuresGrid: {
    display: 'grid',
    gap: '32px',
    '@media (min-width: 768px)': {
      gridTemplateColumns: 'repeat(3, 1fr)',
    },
  },
  featureCard: {
    padding: '32px',
    borderRadius: '16px',
    backgroundColor: '#f9fafb',
    border: '1px solid #f3f4f6',
  },
  featureIcon: {
    width: '48px',
    height: '48px',
    borderRadius: '12px',
    backgroundColor: 'rgba(33, 115, 70, 0.1)',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    marginBottom: '20px',
    color: '#217346',
  },
  featureTitle: {
    fontWeight: '600',
    fontSize: '20px',
    color: '#111827',
    marginBottom: '8px',
  },
  featureDescription: {
    color: '#6b7280',
    lineHeight: 1.6,
  },
  stepsList: {
    display: 'flex',
    flexDirection: 'column',
    gap: '40px',
  },
  step: {
    display: 'flex',
    alignItems: 'flex-start',
    gap: '24px',
  },
  stepNumber: {
    flexShrink: 0,
    width: '48px',
    height: '48px',
    borderRadius: '50%',
    backgroundColor: '#217346',
    color: '#ffffff',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    fontWeight: '700',
    fontSize: '18px',
  },
  stepTitle: {
    fontWeight: '600',
    fontSize: '20px',
    color: '#111827',
    marginBottom: '8px',
  },
  stepDescription: {
    color: '#6b7280',
    fontSize: '18px',
  },
  ctaSection: {
    textAlign: 'center',
    maxWidth: '672px',
    margin: '0 auto',
  },
  footer: {
    padding: '48px 16px',
    backgroundColor: '#f9fafb',
    borderTop: '1px solid #e5e7eb',
  },
  footerContent: {
    maxWidth: '1024px',
    margin: '0 auto',
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    gap: '24px',
    '@media (min-width: 768px)': {
      flexDirection: 'row',
      justifyContent: 'space-between',
    },
  },
  footerBrand: {
    display: 'flex',
    alignItems: 'center',
    gap: '10px',
    fontWeight: '700',
    fontSize: '20px',
    color: '#111827',
  },
  footerTagline: {
    color: '#6b7280',
  },
  footerLinks: {
    display: 'flex',
    alignItems: 'center',
    gap: '24px',
  },
  footerLink: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
    color: '#4b5563',
    textDecoration: 'none',
    ':hover': {
      color: '#111827',
    },
  },
  footerDivider: {
    color: '#d1d5db',
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
      width="20"
      height="20"
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
      width="20"
      height="20"
    >
      <path d="M12 0c-6.626 0-12 5.373-12 12 0 5.302 3.438 9.8 8.207 11.387.599.111.793-.261.793-.577v-2.234c-3.338.726-4.033-1.416-4.033-1.416-.546-1.387-1.333-1.756-1.333-1.756-1.089-.745.083-.729.083-.729 1.205.084 1.839 1.237 1.839 1.237 1.07 1.834 2.807 1.304 3.492.997.107-.775.418-1.305.762-1.604-2.665-.305-5.467-1.334-5.467-5.931 0-1.311.469-2.381 1.236-3.221-.124-.303-.535-1.524.117-3.176 0 0 1.008-.322 3.301 1.23.957-.266 1.983-.399 3.003-.404 1.02.005 2.047.138 3.006.404 2.291-1.552 3.297-1.23 3.297-1.23.653 1.653.242 2.874.118 3.176.77.84 1.235 1.911 1.235 3.221 0 4.609-2.807 5.624-5.479 5.921.43.372.823 1.102.823 2.222v3.293c0 .319.192.694.801.576 4.765-1.589 8.199-6.086 8.199-11.386 0-6.627-5.373-12-12-12z" />
    </svg>
  );
}

function LandingPage() {
  const styles = useStyles();
  const { instance } = useMsal();
  const isAuthenticated = useIsAuthenticated();
  const navigate = useNavigate();

  useEffect(() => {
    if (isAuthenticated) {
      navigate('/app');
    }
  }, [isAuthenticated, navigate]);

  const handleSignIn = async () => {
    try {
      await instance.loginPopup(loginRequest);
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
              <Logo size={80} />
            </div>
            <h1 className={styles.heroTitle}>
              SharePoint lists,
              <br />
              <span className={styles.heroHighlight}>simplified</span>
            </h1>
            <p className={styles.heroDescription}>
              A lightweight tool to manage CRM and knowledge data
              in SharePoint lists.
            </p>
            <div className={styles.heroButtons}>
              <button onClick={handleSignIn} className={styles.primaryButton}>
                <MicrosoftIcon />
                Sign in with Microsoft
              </button>
              <a
                href="https://github.com/emildosen/listview"
                target="_blank"
                rel="noopener noreferrer"
                className={styles.secondaryButton}
              >
                View on GitHub
              </a>
            </div>
          </div>
        </section>

        {/* Features Section */}
        <section className={styles.section}>
          <div className={styles.sectionContent}>
            <h2 className={styles.sectionTitle}>Built for simplicity</h2>
            <p className={styles.sectionSubtitle}>
              Everything you need to manage your data, nothing you don't.
            </p>
            <div className={styles.featuresGrid}>
              <div className={styles.featureCard}>
                <div className={styles.featureIcon}>
                  <LockClosedRegular fontSize={24} />
                </div>
                <h3 className={styles.featureTitle}>Secure by default</h3>
                <p className={styles.featureDescription}>
                  Uses your existing Microsoft 365 login. Your data stays in
                  your SharePoint—we never store it.
                </p>
              </div>

              <div className={styles.featureCard}>
                <div className={styles.featureIcon}>
                  <GridRegular fontSize={24} />
                </div>
                <h3 className={styles.featureTitle}>No code needed</h3>
                <p className={styles.featureDescription}>
                  Configure your lists and views without writing formulas or
                  learning Power Apps.
                </p>
              </div>

              <div className={styles.featureCard}>
                <div className={styles.featureIcon}>
                  <DataBarVerticalRegular fontSize={24} />
                </div>
                <h3 className={styles.featureTitle}>Custom reports</h3>
                <p className={styles.featureDescription}>
                  Create rollup views and reports across your lists to track
                  what matters.
                </p>
              </div>
            </div>
          </div>
        </section>

        {/* How it works */}
        <section className={`${styles.section} ${styles.sectionGray}`}>
          <div className={`${styles.sectionContent} ${styles.sectionContentNarrow}`}>
            <h2 className={styles.sectionTitle}>How it works</h2>
            <div className={styles.stepsList}>
              <div className={styles.step}>
                <div className={styles.stepNumber}>1</div>
                <div>
                  <h3 className={styles.stepTitle}>Sign in with Microsoft</h3>
                  <p className={styles.stepDescription}>
                    Use your existing M365 account. No new passwords to remember.
                  </p>
                </div>
              </div>
              <div className={styles.step}>
                <div className={styles.stepNumber}>2</div>
                <div>
                  <h3 className={styles.stepTitle}>Connect your SharePoint lists</h3>
                  <p className={styles.stepDescription}>
                    Point ListView at any SharePoint list you have access to.
                  </p>
                </div>
              </div>
              <div className={styles.step}>
                <div className={styles.stepNumber}>3</div>
                <div>
                  <h3 className={styles.stepTitle}>Start managing your data</h3>
                  <p className={styles.stepDescription}>
                    View, edit, and report on your data with a clean, simple
                    interface.
                  </p>
                </div>
              </div>
            </div>
          </div>
        </section>

        {/* CTA Section */}
        <section className={styles.section}>
          <div className={styles.ctaSection}>
            <h2 className={styles.sectionTitle}>Ready to simplify?</h2>
            <p className={styles.sectionSubtitle}>
              Free and open source. Start managing your SharePoint lists today.
            </p>
            <button onClick={handleSignIn} className={styles.primaryButton}>
              <MicrosoftIcon />
              Sign in with Microsoft
            </button>
          </div>
        </section>

        {/* Footer */}
        <footer className={styles.footer}>
          <div className={styles.footerContent}>
            <div>
              <p className={styles.footerBrand}>
                <Logo size={24} />
                ListView
              </p>
              <p className={styles.footerTagline}>
                Open source SharePoint list management
              </p>
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
