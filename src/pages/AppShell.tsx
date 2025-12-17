import { useIsAuthenticated } from '@azure/msal-react';
import { useNavigate, Outlet } from 'react-router-dom';
import { useEffect } from 'react';
import { useSettings } from '../contexts/SettingsContext';
import { SetupWizard } from '../components/SetupWizard';
import AppLayout from '../components/AppLayout';

function AppShell() {
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
        <div className="p-8">
          <SetupWizard />
        </div>
      </AppLayout>
    );
  }

  // Show loading state
  if (isLoading) {
    return (
      <AppLayout>
        <div className="flex items-center justify-center h-64">
          <span className="loading loading-spinner loading-lg" />
        </div>
      </AppLayout>
    );
  }

  // Render child routes
  return (
    <AppLayout>
      <Outlet />
    </AppLayout>
  );
}

export default AppShell;
