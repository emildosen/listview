import { useIsAuthenticated, useMsal } from '@azure/msal-react';
import { InteractionStatus } from '@azure/msal-browser';
import { useNavigate, Outlet } from 'react-router-dom';
import { useEffect } from 'react';
import { makeStyles, Spinner } from '@fluentui/react-components';
import { useSettings } from '../contexts/SettingsContext';
import { SetupWizard } from '../components/SetupWizard';
import AppLayout from '../components/AppLayout';

const useStyles = makeStyles({
  setupContainer: {
    padding: '32px',
    flex: 1,
    display: 'flex',
    flexDirection: 'column',
    justifyContent: 'center',
  },
  loadingContainer: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    flex: 1,
  },
  outletContainer: {
    flex: 1,
    display: 'flex',
    flexDirection: 'column',
  },
});

function AppShell() {
  const styles = useStyles();
  const { inProgress } = useMsal();
  const isAuthenticated = useIsAuthenticated();
  const navigate = useNavigate();
  const { setupStatus } = useSettings();

  const isLoading = inProgress !== InteractionStatus.None;

  useEffect(() => {
    // Only redirect when MSAL has finished initializing and user is not authenticated
    if (!isLoading && !isAuthenticated) {
      navigate('/');
    }
  }, [isLoading, isAuthenticated, navigate]);

  // Show loading while MSAL is initializing
  if (isLoading || !isAuthenticated) {
    return (
      <AppLayout>
        <div className={styles.loadingContainer}>
          <Spinner size="large" />
        </div>
      </AppLayout>
    );
  }

  const isReady = setupStatus === 'ready';
  const isSetupLoading = setupStatus === 'loading';

  // Show setup wizard if not ready
  if (!isReady && !isSetupLoading) {
    return (
      <AppLayout>
        <div className={styles.setupContainer}>
          <SetupWizard />
        </div>
      </AppLayout>
    );
  }

  // Show loading state
  if (isSetupLoading) {
    return (
      <AppLayout>
        <div className={styles.loadingContainer}>
          <Spinner size="large" />
        </div>
      </AppLayout>
    );
  }

  // Render child routes
  return (
    <AppLayout>
      <div className={styles.outletContainer}>
        <Outlet />
      </div>
    </AppLayout>
  );
}

export default AppShell;
