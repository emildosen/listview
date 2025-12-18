import { useIsAuthenticated } from '@azure/msal-react';
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
  const isAuthenticated = useIsAuthenticated();
  const navigate = useNavigate();
  const { setupStatus } = useSettings();

  useEffect(() => {
    if (!isAuthenticated) {
      navigate('/');
    }
  }, [isAuthenticated, navigate]);

  if (!isAuthenticated) {
    return null;
  }

  const isReady = setupStatus === 'ready';
  const isLoading = setupStatus === 'loading';

  // Show setup wizard if not ready
  if (!isReady && !isLoading) {
    return (
      <AppLayout>
        <div className={styles.setupContainer}>
          <SetupWizard />
        </div>
      </AppLayout>
    );
  }

  // Show loading state
  if (isLoading) {
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
