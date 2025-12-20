import { createRoot } from 'react-dom/client';
import { PublicClientApplication, EventType } from '@azure/msal-browser';
import { MsalProvider } from '@azure/msal-react';
import './index.css';
import App from './App.tsx';
import { msalConfig } from './auth/msalConfig';
import { TokenRefreshProvider } from './auth/TokenRefreshProvider';
import { AuthErrorProvider } from './contexts/AuthErrorContext';
import { SessionExpiredModal } from './components/modals/SessionExpiredModal';

const msalInstance = new PublicClientApplication(msalConfig);

msalInstance.initialize().then(() => {
  const accounts = msalInstance.getAllAccounts();
  if (accounts.length > 0) {
    msalInstance.setActiveAccount(accounts[0]);
  }

  msalInstance.addEventCallback((event) => {
    if (event.eventType === EventType.LOGIN_SUCCESS && event.payload) {
      const payload = event.payload as { account: { username: string } };
      const account = payload.account;
      msalInstance.setActiveAccount(msalInstance.getAccountByUsername(account.username));
    }
  });

  createRoot(document.getElementById('root')!).render(
    <AuthErrorProvider>
      <MsalProvider instance={msalInstance}>
        <TokenRefreshProvider>
          <App />
          <SessionExpiredModal />
        </TokenRefreshProvider>
      </MsalProvider>
    </AuthErrorProvider>
  );
});
