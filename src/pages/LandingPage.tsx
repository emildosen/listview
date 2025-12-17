import { useMsal, useIsAuthenticated } from '@azure/msal-react';
import { useNavigate } from 'react-router-dom';
import { useEffect } from 'react';
import { loginRequest } from '../auth/msalConfig';

function LandingPage() {
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
    <div className="min-h-screen bg-white" data-theme="light">
      {/* Hero Section */}
      <section className="min-h-screen flex flex-col justify-center px-4 bg-gradient-to-b from-[#f0fdf4] to-white">
        <div className="container mx-auto max-w-3xl text-center">
          <h1 className="text-5xl sm:text-6xl md:text-7xl font-bold tracking-tight text-gray-900 mb-6">
            SharePoint lists,
            <br />
            <span className="text-[#217346]">simplified</span>
          </h1>
          <p className="text-xl md:text-2xl text-gray-600 mb-12 max-w-xl mx-auto leading-relaxed">
            A lightweight tool to manage CRM and <br></br>knowledge data
            in SharePoint lists.
          </p>
          <div className="flex flex-col sm:flex-row gap-4 justify-center">
            <button
              onClick={handleSignIn}
              className="inline-flex items-center justify-center gap-2 px-8 py-4 bg-[#217346] hover:bg-[#1a5c38] text-white font-semibold rounded-lg transition-colors text-lg cursor-pointer"
            >
              <svg
                xmlns="http://www.w3.org/2000/svg"
                viewBox="0 0 23 23"
                className="w-5 h-5"
                fill="currentColor"
              >
                <path d="M0 0h11v11H0zM12 0h11v11H12zM0 12h11v11H0zM12 12h11v11H12z" />
              </svg>
              Sign in with Microsoft
            </button>
            <a
              href="https://github.com/emildosen/listview"
              target="_blank"
              rel="noopener noreferrer"
              className="inline-flex items-center justify-center gap-2 px-8 py-4 border-2 border-gray-300 hover:border-gray-400 text-gray-700 font-semibold rounded-lg transition-colors text-lg"
            >
              View on GitHub
            </a>
          </div>
        </div>
      </section>

      {/* Features Section */}
      <section className="py-24 px-4 bg-white">
        <div className="container mx-auto max-w-5xl">
          <h2 className="text-3xl md:text-4xl font-bold text-center text-gray-900 mb-4">
            Built for simplicity
          </h2>
          <p className="text-center text-gray-500 mb-16 max-w-xl mx-auto text-lg">
            Everything you need to manage your data, nothing you don't.
          </p>
          <div className="grid md:grid-cols-3 gap-8">
            <div className="p-8 rounded-2xl bg-gray-50 border border-gray-100">
              <div className="w-12 h-12 rounded-xl bg-[#217346]/10 flex items-center justify-center mb-5">
                <svg
                  xmlns="http://www.w3.org/2000/svg"
                  fill="none"
                  viewBox="0 0 24 24"
                  strokeWidth={1.5}
                  stroke="currentColor"
                  className="w-6 h-6 text-[#217346]"
                >
                  <path
                    strokeLinecap="round"
                    strokeLinejoin="round"
                    d="M16.5 10.5V6.75a4.5 4.5 0 1 0-9 0v3.75m-.75 11.25h10.5a2.25 2.25 0 0 0 2.25-2.25v-6.75a2.25 2.25 0 0 0-2.25-2.25H6.75a2.25 2.25 0 0 0-2.25 2.25v6.75a2.25 2.25 0 0 0 2.25 2.25Z"
                  />
                </svg>
              </div>
              <h3 className="font-semibold text-xl text-gray-900 mb-2">Secure by default</h3>
              <p className="text-gray-500 leading-relaxed">
                Uses your existing Microsoft 365 login. Your data stays in
                your SharePoint—we never store it.
              </p>
            </div>

            <div className="p-8 rounded-2xl bg-gray-50 border border-gray-100">
              <div className="w-12 h-12 rounded-xl bg-[#217346]/10 flex items-center justify-center mb-5">
                <svg
                  xmlns="http://www.w3.org/2000/svg"
                  fill="none"
                  viewBox="0 0 24 24"
                  strokeWidth={1.5}
                  stroke="currentColor"
                  className="w-6 h-6 text-[#217346]"
                >
                  <path
                    strokeLinecap="round"
                    strokeLinejoin="round"
                    d="M3.75 6A2.25 2.25 0 0 1 6 3.75h2.25A2.25 2.25 0 0 1 10.5 6v2.25a2.25 2.25 0 0 1-2.25 2.25H6a2.25 2.25 0 0 1-2.25-2.25V6ZM3.75 15.75A2.25 2.25 0 0 1 6 13.5h2.25a2.25 2.25 0 0 1 2.25 2.25V18a2.25 2.25 0 0 1-2.25 2.25H6A2.25 2.25 0 0 1 3.75 18v-2.25ZM13.5 6a2.25 2.25 0 0 1 2.25-2.25H18A2.25 2.25 0 0 1 20.25 6v2.25A2.25 2.25 0 0 1 18 10.5h-2.25a2.25 2.25 0 0 1-2.25-2.25V6ZM13.5 15.75a2.25 2.25 0 0 1 2.25-2.25H18a2.25 2.25 0 0 1 2.25 2.25V18A2.25 2.25 0 0 1 18 20.25h-2.25A2.25 2.25 0 0 1 13.5 18v-2.25Z"
                  />
                </svg>
              </div>
              <h3 className="font-semibold text-xl text-gray-900 mb-2">No code needed</h3>
              <p className="text-gray-500 leading-relaxed">
                Configure your lists and views without writing formulas or
                learning Power Apps.
              </p>
            </div>

            <div className="p-8 rounded-2xl bg-gray-50 border border-gray-100">
              <div className="w-12 h-12 rounded-xl bg-[#217346]/10 flex items-center justify-center mb-5">
                <svg
                  xmlns="http://www.w3.org/2000/svg"
                  fill="none"
                  viewBox="0 0 24 24"
                  strokeWidth={1.5}
                  stroke="currentColor"
                  className="w-6 h-6 text-[#217346]"
                >
                  <path
                    strokeLinecap="round"
                    strokeLinejoin="round"
                    d="M3 13.125C3 12.504 3.504 12 4.125 12h2.25c.621 0 1.125.504 1.125 1.125v6.75C7.5 20.496 6.996 21 6.375 21h-2.25A1.125 1.125 0 0 1 3 19.875v-6.75ZM9.75 8.625c0-.621.504-1.125 1.125-1.125h2.25c.621 0 1.125.504 1.125 1.125v11.25c0 .621-.504 1.125-1.125 1.125h-2.25a1.125 1.125 0 0 1-1.125-1.125V8.625ZM16.5 4.125c0-.621.504-1.125 1.125-1.125h2.25C20.496 3 21 3.504 21 4.125v15.75c0 .621-.504 1.125-1.125 1.125h-2.25a1.125 1.125 0 0 1-1.125-1.125V4.125Z"
                  />
                </svg>
              </div>
              <h3 className="font-semibold text-xl text-gray-900 mb-2">Custom reports</h3>
              <p className="text-gray-500 leading-relaxed">
                Create rollup views and reports across your lists to track
                what matters.
              </p>
            </div>
          </div>
        </div>
      </section>

      {/* How it works */}
      <section className="py-24 px-4 bg-[#f9fafb]">
        <div className="container mx-auto max-w-4xl">
          <h2 className="text-3xl md:text-4xl font-bold text-center text-gray-900 mb-16">
            How it works
          </h2>
          <div className="flex flex-col gap-10">
            <div className="flex items-start gap-6">
              <div className="flex-shrink-0 w-12 h-12 rounded-full bg-[#217346] text-white flex items-center justify-center font-bold text-lg">
                1
              </div>
              <div>
                <h3 className="font-semibold text-xl text-gray-900 mb-2">
                  Sign in with Microsoft
                </h3>
                <p className="text-gray-500 text-lg">
                  Use your existing M365 account. No new passwords to remember.
                </p>
              </div>
            </div>
            <div className="flex items-start gap-6">
              <div className="flex-shrink-0 w-12 h-12 rounded-full bg-[#217346] text-white flex items-center justify-center font-bold text-lg">
                2
              </div>
              <div>
                <h3 className="font-semibold text-xl text-gray-900 mb-2">
                  Connect your SharePoint lists
                </h3>
                <p className="text-gray-500 text-lg">
                  Point ListView at any SharePoint list you have access to.
                </p>
              </div>
            </div>
            <div className="flex items-start gap-6">
              <div className="flex-shrink-0 w-12 h-12 rounded-full bg-[#217346] text-white flex items-center justify-center font-bold text-lg">
                3
              </div>
              <div>
                <h3 className="font-semibold text-xl text-gray-900 mb-2">
                  Start managing your data
                </h3>
                <p className="text-gray-500 text-lg">
                  View, edit, and report on your data with a clean, simple
                  interface.
                </p>
              </div>
            </div>
          </div>
        </div>
      </section>

      {/* CTA Section */}
      <section className="py-24 px-4 bg-white">
        <div className="container mx-auto max-w-2xl text-center">
          <h2 className="text-3xl md:text-4xl font-bold text-gray-900 mb-4">
            Ready to simplify?
          </h2>
          <p className="text-gray-500 mb-10 text-lg">
            Free and open source. Start managing your SharePoint lists today.
          </p>
          <button
            onClick={handleSignIn}
            className="inline-flex items-center justify-center gap-2 px-8 py-4 bg-[#217346] hover:bg-[#1a5c38] text-white font-semibold rounded-lg transition-colors text-lg cursor-pointer"
          >
            <svg
              xmlns="http://www.w3.org/2000/svg"
              viewBox="0 0 23 23"
              className="w-5 h-5"
              fill="currentColor"
            >
              <path d="M0 0h11v11H0zM12 0h11v11H12zM0 12h11v11H0zM12 12h11v11H12z" />
            </svg>
            Sign in with Microsoft
          </button>
        </div>
      </section>

      {/* Footer */}
      <footer className="py-12 px-4 bg-[#f9fafb] border-t border-gray-200">
        <div className="container mx-auto max-w-5xl">
          <div className="flex flex-col md:flex-row items-center justify-between gap-6">
            <div>
              <p className="font-bold text-xl text-gray-900">ListView</p>
              <p className="text-gray-500">
                Open source SharePoint list management
              </p>
            </div>
            <div className="flex items-center gap-6">
              <a
                href="https://github.com/emildosen/listview"
                target="_blank"
                rel="noopener noreferrer"
                className="flex items-center gap-2 text-gray-600 hover:text-gray-900 transition-colors"
              >
                <svg
                  xmlns="http://www.w3.org/2000/svg"
                  viewBox="0 0 24 24"
                  fill="currentColor"
                  className="w-5 h-5"
                >
                  <path d="M12 0c-6.626 0-12 5.373-12 12 0 5.302 3.438 9.8 8.207 11.387.599.111.793-.261.793-.577v-2.234c-3.338.726-4.033-1.416-4.033-1.416-.546-1.387-1.333-1.756-1.333-1.756-1.089-.745.083-.729.083-.729 1.205.084 1.839 1.237 1.839 1.237 1.07 1.834 2.807 1.304 3.492.997.107-.775.418-1.305.762-1.604-2.665-.305-5.467-1.334-5.467-5.931 0-1.311.469-2.381 1.236-3.221-.124-.303-.535-1.524.117-3.176 0 0 1.008-.322 3.301 1.23.957-.266 1.983-.399 3.003-.404 1.02.005 2.047.138 3.006.404 2.291-1.552 3.297-1.23 3.297-1.23.653 1.653.242 2.874.118 3.176.77.84 1.235 1.911 1.235 3.221 0 4.609-2.807 5.624-5.479 5.921.43.372.823 1.102.823 2.222v3.293c0 .319.192.694.801.576 4.765-1.589 8.199-6.086 8.199-11.386 0-6.627-5.373-12-12-12z" />
                </svg>
                GitHub
              </a>
              <span className="text-gray-300">|</span>
              <span className="text-gray-400 text-sm">MIT License</span>
            </div>
          </div>
        </div>
      </footer>
    </div>
  );
}

export default LandingPage;
