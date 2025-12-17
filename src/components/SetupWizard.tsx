import { useState } from 'react';
import { useSettings } from '../contexts/SettingsContext';
import { DEFAULT_SETTINGS_SITE_PATH } from '../services/sharepoint';

type SetupStep = 'choice' | 'standard-instructions' | 'custom-input';

export function SetupWizard() {
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
      <div className="flex flex-col items-center justify-center min-h-[400px]">
        <span className="loading loading-spinner loading-lg text-primary"></span>
        <p className="mt-4 text-base-content/60">Connecting to SharePoint...</p>
      </div>
    );
  }

  // Creating list state
  if (setupStatus === 'creating-list') {
    return (
      <div className="flex flex-col items-center justify-center min-h-[400px]">
        <span className="loading loading-spinner loading-lg text-primary"></span>
        <p className="mt-4 text-base-content/60">Creating settings list...</p>
      </div>
    );
  }

  // General error state
  if (setupStatus === 'error') {
    return (
      <div className="max-w-lg mx-auto">
        <div className="card bg-base-200">
          <div className="card-body text-center">
            <div className="w-16 h-16 rounded-full bg-error/10 flex items-center justify-center mx-auto">
              <svg
                xmlns="http://www.w3.org/2000/svg"
                fill="none"
                viewBox="0 0 24 24"
                strokeWidth={1.5}
                stroke="currentColor"
                className="w-8 h-8 text-error"
              >
                <path
                  strokeLinecap="round"
                  strokeLinejoin="round"
                  d="M12 9v3.75m-9.303 3.376c-.866 1.5.217 3.374 1.948 3.374h14.71c1.73 0 2.813-1.874 1.948-3.374L13.949 3.378c-.866-1.5-3.032-1.5-3.898 0L2.697 16.126ZM12 15.75h.007v.008H12v-.008Z"
                />
              </svg>
            </div>
            <h2 className="card-title justify-center mt-4">Connection Error</h2>
            <p className="text-base-content/60">
              Unable to connect to SharePoint. Please check your permissions and
              try again.
            </p>
            {error && (
              <p className="text-sm text-error mt-2 font-mono">{error}</p>
            )}
            <div className="card-actions justify-center mt-4">
              <button onClick={handleRetry} className="btn btn-primary">
                Retry
              </button>
            </div>
          </div>
        </div>
      </div>
    );
  }

  // List not found - prompt to create
  if (setupStatus === 'list-not-found') {
    return (
      <div className="max-w-lg mx-auto">
        <div className="card bg-base-200">
          <div className="card-body">
            <div className="flex items-center gap-3 mb-2">
              <div className="w-10 h-10 rounded-full bg-success/10 flex items-center justify-center">
                <svg
                  xmlns="http://www.w3.org/2000/svg"
                  fill="none"
                  viewBox="0 0 24 24"
                  strokeWidth={1.5}
                  stroke="currentColor"
                  className="w-5 h-5 text-success"
                >
                  <path
                    strokeLinecap="round"
                    strokeLinejoin="round"
                    d="M9 12.75 11.25 15 15 9.75M21 12a9 9 0 1 1-18 0 9 9 0 0 1 18 0Z"
                  />
                </svg>
              </div>
              <div>
                <h2 className="font-bold">Site Connected</h2>
                <p className="text-sm text-base-content/60">{site?.displayName}</p>
              </div>
            </div>

            <div className="divider"></div>

            <h3 className="font-semibold">Create Settings List</h3>
            <p className="text-base-content/60 text-sm">
              The site exists but doesn't have an LV-Settings list yet.
              Click below to create it.
            </p>

            <div className="card-actions justify-between mt-4">
              <button onClick={handleBackToChoice} className="btn btn-ghost">
                Back
              </button>
              <button onClick={handleCreateList} className="btn btn-primary">
                Create Settings List
              </button>
            </div>
          </div>
        </div>
      </div>
    );
  }

  // List creation failed
  if (setupStatus === 'list-creation-failed') {
    return (
      <div className="max-w-lg mx-auto">
        <div className="card bg-base-200">
          <div className="card-body">
            <div className="flex items-center gap-3 mb-2">
              <div className="w-10 h-10 rounded-full bg-error/10 flex items-center justify-center">
                <svg
                  xmlns="http://www.w3.org/2000/svg"
                  fill="none"
                  viewBox="0 0 24 24"
                  strokeWidth={1.5}
                  stroke="currentColor"
                  className="w-5 h-5 text-error"
                >
                  <path
                    strokeLinecap="round"
                    strokeLinejoin="round"
                    d="M12 9v3.75m-9.303 3.376c-.866 1.5.217 3.374 1.948 3.374h14.71c1.73 0 2.813-1.874 1.948-3.374L13.949 3.378c-.866-1.5-3.032-1.5-3.898 0L2.697 16.126ZM12 15.75h.007v.008H12v-.008Z"
                  />
                </svg>
              </div>
              <div>
                <h2 className="font-bold">Failed to Create List</h2>
                <p className="text-sm text-base-content/60">
                  Could not create the settings list on {site?.displayName}
                </p>
              </div>
            </div>

            <div className="alert alert-error mt-4">
              <svg
                xmlns="http://www.w3.org/2000/svg"
                className="h-5 w-5 shrink-0"
                fill="none"
                viewBox="0 0 24 24"
                stroke="currentColor"
              >
                <path
                  strokeLinecap="round"
                  strokeLinejoin="round"
                  strokeWidth="2"
                  d="M10 14l2-2m0 0l2-2m-2 2l-2-2m2 2l2 2m7-2a9 9 0 11-18 0 9 9 0 0118 0z"
                />
              </svg>
              <div>
                <p className="font-medium">Access Denied</p>
                <p className="text-sm">
                  {error || 'You may not have permission to create lists on this site.'}
                </p>
              </div>
            </div>

            <div className="mt-4 p-4 bg-base-300 rounded-lg">
              <p className="font-medium text-sm mb-2">To fix this:</p>
              <ol className="list-decimal list-inside text-sm text-base-content/70 space-y-1">
                <li>Ask a site owner to add you as a site member with Edit permissions</li>
                <li>Or create the "LV-Settings" list manually in SharePoint</li>
                <li>Then click Retry below</li>
              </ol>
            </div>

            <div className="card-actions justify-between mt-4">
              <button onClick={handleBackToChoice} className="btn btn-ghost">
                Use Different Site
              </button>
              <div className="flex gap-2">
                <button onClick={handleCreateList} className="btn btn-outline">
                  Retry
                </button>
                <a
                  href={site?.webUrl}
                  target="_blank"
                  rel="noopener noreferrer"
                  className="btn btn-primary"
                >
                  Open Site
                </a>
              </div>
            </div>
          </div>
        </div>
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
    <div className="max-w-2xl mx-auto">
      <div className="text-center mb-8">
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
              d="M9.594 3.94c.09-.542.56-.94 1.11-.94h2.593c.55 0 1.02.398 1.11.94l.213 1.281c.063.374.313.686.645.87.074.04.147.083.22.127.325.196.72.257 1.075.124l1.217-.456a1.125 1.125 0 0 1 1.37.49l1.296 2.247a1.125 1.125 0 0 1-.26 1.431l-1.003.827c-.293.241-.438.613-.43.992a7.723 7.723 0 0 1 0 .255c-.008.378.137.75.43.991l1.004.827c.424.35.534.955.26 1.43l-1.298 2.247a1.125 1.125 0 0 1-1.369.491l-1.217-.456c-.355-.133-.75-.072-1.076.124a6.47 6.47 0 0 1-.22.128c-.331.183-.581.495-.644.869l-.213 1.281c-.09.543-.56.94-1.11.94h-2.594c-.55 0-1.019-.398-1.11-.94l-.213-1.281c-.062-.374-.312-.686-.644-.87a6.52 6.52 0 0 1-.22-.127c-.325-.196-.72-.257-1.076-.124l-1.217.456a1.125 1.125 0 0 1-1.369-.49l-1.297-2.247a1.125 1.125 0 0 1 .26-1.431l1.004-.827c.292-.24.437-.613.43-.991a6.932 6.932 0 0 1 0-.255c.007-.38-.138-.751-.43-.992l-1.004-.827a1.125 1.125 0 0 1-.26-1.43l1.297-2.247a1.125 1.125 0 0 1 1.37-.491l1.216.456c.356.133.751.072 1.076-.124.072-.044.146-.086.22-.128.332-.183.582-.495.644-.869l.214-1.28Z"
            />
            <path
              strokeLinecap="round"
              strokeLinejoin="round"
              d="M15 12a3 3 0 1 1-6 0 3 3 0 0 1 6 0Z"
            />
          </svg>
        </div>
        <h1 className="text-2xl font-bold">Set Up ListView</h1>
        <p className="text-base-content/60 mt-2">
          ListView needs a SharePoint site to store app settings and data.
        </p>
      </div>

      {step === 'choice' && (
        <div className="grid md:grid-cols-2 gap-4">
          <div className="card bg-base-200 hover:bg-base-300 transition-colors cursor-pointer">
            <div className="card-body" onClick={handleUseStandard}>
              <h2 className="card-title">
                <svg
                  xmlns="http://www.w3.org/2000/svg"
                  fill="none"
                  viewBox="0 0 24 24"
                  strokeWidth={1.5}
                  stroke="currentColor"
                  className="w-5 h-5 text-primary"
                >
                  <path
                    strokeLinecap="round"
                    strokeLinejoin="round"
                    d="M9 12.75 11.25 15 15 9.75M21 12a9 9 0 1 1-18 0 9 9 0 0 1 18 0Z"
                  />
                </svg>
                Standard Setup
              </h2>
              <p className="text-base-content/60 text-sm">
                Use the default <code>/sites/ListView</code> site. This is
                shared across all users in your organization.
              </p>
              <div className="badge badge-primary badge-outline mt-2">
                Recommended
              </div>
            </div>
          </div>

          <div className="card bg-base-200 hover:bg-base-300 transition-colors cursor-pointer">
            <div className="card-body" onClick={handleUseCustom}>
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
                    d="m16.862 4.487 1.687-1.688a1.875 1.875 0 1 1 2.652 2.652L10.582 16.07a4.5 4.5 0 0 1-1.897 1.13L6 18l.8-2.685a4.5 4.5 0 0 1 1.13-1.897l8.932-8.931Zm0 0L19.5 7.125M18 14v4.75A2.25 2.25 0 0 1 15.75 21H5.25A2.25 2.25 0 0 1 3 18.75V8.25A2.25 2.25 0 0 1 5.25 6H10"
                  />
                </svg>
                Custom Site
              </h2>
              <p className="text-base-content/60 text-sm">
                Use a different SharePoint site. Each user must configure this
                manually.
              </p>
              <div className="badge badge-ghost mt-2">Advanced</div>
            </div>
          </div>
        </div>
      )}

      {step === 'standard-instructions' && (
        <div className="card bg-base-200">
          <div className="card-body">
            <h2 className="card-title">Create the ListView Site</h2>
            <p className="text-base-content/60">
              A SharePoint site needs to be created at{' '}
              <code>/sites/ListView</code> by a SharePoint admin.
            </p>

            <div className="divider"></div>

            <div className="space-y-4">
              <div className="flex gap-3">
                <div className="badge badge-primary badge-lg">1</div>
                <div>
                  <p className="font-medium">
                    Go to SharePoint Admin Center
                  </p>
                  <p className="text-sm text-base-content/60">
                    Navigate to{' '}
                    <a
                      href="https://admin.microsoft.com/sharepoint"
                      target="_blank"
                      rel="noopener noreferrer"
                      className="link link-primary"
                    >
                      admin.microsoft.com/sharepoint
                    </a>
                  </p>
                </div>
              </div>

              <div className="flex gap-3">
                <div className="badge badge-primary badge-lg">2</div>
                <div>
                  <p className="font-medium">Create a new Team site</p>
                  <p className="text-sm text-base-content/60">
                    Click "Create" and select "Team site"
                  </p>
                </div>
              </div>

              <div className="flex gap-3">
                <div className="badge badge-primary badge-lg">3</div>
                <div>
                  <p className="font-medium">
                    Set the site address to "ListView"
                  </p>
                  <p className="text-sm text-base-content/60">
                    The URL should be:{' '}
                    {sharePointUrl ? (
                      <code className="text-xs">{sharePointUrl}</code>
                    ) : (
                      <code className="text-xs">
                        https://[tenant].sharepoint.com/sites/ListView
                      </code>
                    )}
                  </p>
                </div>
              </div>

              <div className="flex gap-3">
                <div className="badge badge-primary badge-lg">4</div>
                <div>
                  <p className="font-medium">Grant access to users</p>
                  <p className="text-sm text-base-content/60">
                    Add users who need to use ListView as site members
                  </p>
                </div>
              </div>
            </div>

            {checkError && (
              <div className="alert alert-warning mt-4">
                <svg
                  xmlns="http://www.w3.org/2000/svg"
                  className="h-5 w-5 shrink-0"
                  fill="none"
                  viewBox="0 0 24 24"
                  stroke="currentColor"
                >
                  <path
                    strokeLinecap="round"
                    strokeLinejoin="round"
                    strokeWidth="2"
                    d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-3L13.732 4c-.77-1.333-2.694-1.333-3.464 0L3.34 16c-.77 1.333.192 3 1.732 3z"
                  />
                </svg>
                <span>{checkError}</span>
              </div>
            )}

            <div className="card-actions justify-between mt-4">
              <button
                onClick={() => setStep('choice')}
                className="btn btn-ghost"
              >
                Back
              </button>
              <button
                onClick={handleCheckStandardSite}
                className="btn btn-primary"
                disabled={isChecking}
              >
                {isChecking ? (
                  <>
                    <span className="loading loading-spinner loading-sm"></span>
                    Checking...
                  </>
                ) : (
                  "I've created the site"
                )}
              </button>
            </div>
          </div>
        </div>
      )}

      {step === 'custom-input' && (
        <div className="card bg-base-200">
          <div className="card-body">
            <h2 className="card-title">Use Custom Site</h2>
            <p className="text-base-content/60">
              Enter the path to an existing SharePoint site. This setting is
              stored locally and each user must configure it.
            </p>

            <div className="form-control mt-4">
              <label className="label">
                <span className="label-text">Site Path</span>
              </label>
              <div className="join w-full">
                <span className="join-item flex items-center px-3 bg-base-300 text-base-content/60 text-sm">
                  /sites/
                </span>
                <input
                  type="text"
                  placeholder="MySite"
                  className="input input-bordered join-item flex-1"
                  value={customPath}
                  onChange={(e) => setCustomPath(e.target.value)}
                  onKeyDown={(e) => {
                    if (e.key === 'Enter') {
                      handleCheckCustomSite();
                    }
                  }}
                />
              </div>
              <label className="label">
                <span className="label-text-alt text-base-content/50">
                  Example: "CRM" for /sites/CRM
                </span>
              </label>
            </div>

            {checkError && (
              <div className="alert alert-warning mt-2">
                <svg
                  xmlns="http://www.w3.org/2000/svg"
                  className="h-5 w-5 shrink-0"
                  fill="none"
                  viewBox="0 0 24 24"
                  stroke="currentColor"
                >
                  <path
                    strokeLinecap="round"
                    strokeLinejoin="round"
                    strokeWidth="2"
                    d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-3L13.732 4c-.77-1.333-2.694-1.333-3.464 0L3.34 16c-.77 1.333.192 3 1.732 3z"
                  />
                </svg>
                <span>{checkError}</span>
              </div>
            )}

            <div className="card-actions justify-between mt-4">
              <button
                onClick={() => {
                  setStep('choice');
                  setCheckError(null);
                }}
                className="btn btn-ghost"
              >
                Back
              </button>
              <button
                onClick={handleCheckCustomSite}
                className="btn btn-primary"
                disabled={isChecking || !customPath.trim()}
              >
                {isChecking ? (
                  <>
                    <span className="loading loading-spinner loading-sm"></span>
                    Checking...
                  </>
                ) : (
                  'Connect to Site'
                )}
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
