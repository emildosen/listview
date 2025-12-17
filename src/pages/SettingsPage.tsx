import { useIsAuthenticated } from '@azure/msal-react';
import { useNavigate } from 'react-router-dom';
import { useEffect } from 'react';
import { useSettings } from '../contexts/SettingsContext';
import { SetupWizard } from '../components/SetupWizard';
import AppLayout from '../components/AppLayout';

function SettingsPage() {
  const isAuthenticated = useIsAuthenticated();
  const navigate = useNavigate();
  const {
    setupStatus,
    site,
    sitePath,
    isCustomSite,
    settingsList,
    clearSiteOverride,
    initialize,
  } = useSettings();

  useEffect(() => {
    if (!isAuthenticated) {
      navigate('/');
    }
  }, [isAuthenticated, navigate]);

  const handleResetSite = () => {
    clearSiteOverride();
    initialize();
  };

  if (!isAuthenticated) {
    return null;
  }

  const isReady = setupStatus === 'ready';

  return (
    <AppLayout>
      <div className="p-8">
        {/* Breadcrumb */}
        <div className="text-sm breadcrumbs mb-6">
          <ul>
            <li>
              <a href="/app">Home</a>
            </li>
            <li>Settings</li>
          </ul>
        </div>

        {!isReady ? (
          <SetupWizard />
        ) : (
          <div className="max-w-4xl">
            <h1 className="text-2xl font-bold mb-6">Settings</h1>

            {/* Site Configuration Card */}
            <div className="card bg-base-200 mb-6">
              <div className="card-body">
                <h2 className="card-title">
                  <svg
                    xmlns="http://www.w3.org/2000/svg"
                    fill="none"
                    viewBox="0 0 24 24"
                    strokeWidth={1.5}
                    stroke="currentColor"
                    className="w-5 h-5"
                  >
                    <path
                      strokeLinecap="round"
                      strokeLinejoin="round"
                      d="M2.25 12.75V12A2.25 2.25 0 0 1 4.5 9.75h15A2.25 2.25 0 0 1 21.75 12v.75m-8.69-6.44-2.12-2.12a1.5 1.5 0 0 0-1.061-.44H4.5A2.25 2.25 0 0 0 2.25 6v12a2.25 2.25 0 0 0 2.25 2.25h15A2.25 2.25 0 0 0 21.75 18V9a2.25 2.25 0 0 0-2.25-2.25h-5.379a1.5 1.5 0 0 1-1.06-.44Z"
                    />
                  </svg>
                  SharePoint Site
                </h2>

                <div className="grid gap-4 mt-4">
                  <div className="flex items-center justify-between p-3 bg-base-300 rounded-lg">
                    <div>
                      <p className="font-medium">Connected Site</p>
                      <p className="text-sm text-base-content/60">
                        {site?.displayName}
                      </p>
                    </div>
                    <a
                      href={site?.webUrl}
                      target="_blank"
                      rel="noopener noreferrer"
                      className="btn btn-ghost btn-sm"
                    >
                      Open in SharePoint
                      <svg
                        xmlns="http://www.w3.org/2000/svg"
                        fill="none"
                        viewBox="0 0 24 24"
                        strokeWidth={1.5}
                        stroke="currentColor"
                        className="w-4 h-4"
                      >
                        <path
                          strokeLinecap="round"
                          strokeLinejoin="round"
                          d="M13.5 6H5.25A2.25 2.25 0 0 0 3 8.25v10.5A2.25 2.25 0 0 0 5.25 21h10.5A2.25 2.25 0 0 0 18 18.75V10.5m-10.5 6L21 3m0 0h-5.25M21 3v5.25"
                        />
                      </svg>
                    </a>
                  </div>

                  <div className="flex items-center justify-between p-3 bg-base-300 rounded-lg">
                    <div>
                      <p className="font-medium">Site Path</p>
                      <p className="text-sm text-base-content/60">
                        <code>{sitePath}</code>
                      </p>
                    </div>
                    {isCustomSite ? (
                      <span className="badge badge-warning">Custom</span>
                    ) : (
                      <span className="badge badge-success">Standard</span>
                    )}
                  </div>

                  <div className="flex items-center justify-between p-3 bg-base-300 rounded-lg">
                    <div>
                      <p className="font-medium">Settings List</p>
                      <p className="text-sm text-base-content/60">
                        {settingsList?.displayName}
                      </p>
                    </div>
                    <span className="badge badge-ghost">Active</span>
                  </div>
                </div>

                {isCustomSite && (
                  <div className="card-actions mt-4">
                    <button
                      onClick={handleResetSite}
                      className="btn btn-outline btn-sm"
                    >
                      Reset to Standard Site
                    </button>
                  </div>
                )}
              </div>
            </div>

            {/* App Settings Card (placeholder for future settings) */}
            <div className="card bg-base-200">
              <div className="card-body">
                <h2 className="card-title">
                  <svg
                    xmlns="http://www.w3.org/2000/svg"
                    fill="none"
                    viewBox="0 0 24 24"
                    strokeWidth={1.5}
                    stroke="currentColor"
                    className="w-5 h-5"
                  >
                    <path
                      strokeLinecap="round"
                      strokeLinejoin="round"
                      d="M10.5 6h9.75M10.5 6a1.5 1.5 0 1 1-3 0m3 0a1.5 1.5 0 1 0-3 0M3.75 6H7.5m3 12h9.75m-9.75 0a1.5 1.5 0 0 1-3 0m3 0a1.5 1.5 0 0 0-3 0m-3.75 0H7.5m9-6h3.75m-3.75 0a1.5 1.5 0 0 1-3 0m3 0a1.5 1.5 0 0 0-3 0m-9.75 0h9.75"
                    />
                  </svg>
                  Application Settings
                </h2>
                <p className="text-base-content/60">
                  App-specific settings will appear here as features are added.
                </p>

                <div className="mt-4 p-8 border-2 border-dashed border-base-300 rounded-lg text-center">
                  <p className="text-base-content/40">No settings configured yet</p>
                </div>
              </div>
            </div>

            {/* Back to app */}
            <div className="mt-6">
              <a href="/app" className="btn btn-ghost">
                <svg
                  xmlns="http://www.w3.org/2000/svg"
                  fill="none"
                  viewBox="0 0 24 24"
                  strokeWidth={1.5}
                  stroke="currentColor"
                  className="w-4 h-4"
                >
                  <path
                    strokeLinecap="round"
                    strokeLinejoin="round"
                    d="M10.5 19.5 3 12m0 0 7.5-7.5M3 12h18"
                  />
                </svg>
                Back to App
              </a>
            </div>
          </div>
        )}
      </div>
    </AppLayout>
  );
}

export default SettingsPage;
