import { useIsAuthenticated } from '@azure/msal-react';
import { useNavigate } from 'react-router-dom';
import { useEffect } from 'react';
import { useSettings } from '../contexts/SettingsContext';
import { SetupWizard } from '../components/SetupWizard';
import AppLayout from '../components/AppLayout';

function AppPage() {
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

  return (
    <AppLayout>
      <div className="p-8">
        {!isReady ? (
          <SetupWizard />
        ) : (
          <div className="max-w-lg mx-auto text-center">
            <div className="mb-8">
              <div className="w-16 h-16 rounded-full bg-primary/10 flex items-center justify-center mx-auto mb-4">
                <svg
                  xmlns="http://www.w3.org/2000/svg"
                  fill="none"
                  viewBox="0 0 24 24"
                  strokeWidth={1.5}
                  stroke="currentColor"
                  className="w-8 h-8 text-primary"
                >
                  <path
                    strokeLinecap="round"
                    strokeLinejoin="round"
                    d="M11.42 15.17 17.25 21A2.652 2.652 0 0 0 21 17.25l-5.877-5.877M11.42 15.17l2.496-3.03c.317-.384.74-.626 1.208-.766M11.42 15.17l-4.655 5.653a2.548 2.548 0 1 1-3.586-3.586l6.837-5.63m5.108-.233c.55-.164 1.163-.188 1.743-.14a4.5 4.5 0 0 0 4.486-6.336l-3.276 3.277a3.004 3.004 0 0 1-2.25-2.25l3.276-3.276a4.5 4.5 0 0 0-6.336 4.486c.091 1.076-.071 2.264-.904 2.95l-.102.085m-1.745 1.437L5.909 7.5H4.5L2.25 3.75l1.5-1.5L7.5 4.5v1.409l4.26 4.26m-1.745 1.437 1.745-1.437m6.615 8.206L15.75 15.75M4.867 19.125h.008v.008h-.008v-.008Z"
                  />
                </svg>
              </div>
              <h1 className="text-2xl font-bold mb-2">Coming soon</h1>
              <p className="text-base-content/60">
                The app is currently under development. Check back soon for
                SharePoint list management features.
              </p>
            </div>
            <a href="/" className="btn btn-outline">
              Back to home
            </a>
          </div>
        )}
      </div>
    </AppLayout>
  );
}

export default AppPage;
