import { BrowserRouter, Routes, Route } from 'react-router-dom';
import LandingPage from './pages/LandingPage';
import AppShell from './pages/AppShell';
import HomePage from './pages/HomePage';
import SettingsPage from './pages/SettingsPage';
import PagesPage from './pages/PagesPage';
import PageEditorPage from './pages/PageEditorPage';
import PageDisplayPage from './pages/PageDisplayPage';
import { SettingsProvider } from './contexts/SettingsContext';
import { ThemeProvider } from './contexts/ThemeContext';
import { FormConfigProvider } from './contexts/FormConfigContext';

function App() {
  return (
    <BrowserRouter>
      <ThemeProvider>
        <SettingsProvider>
          <FormConfigProvider>
            <Routes>
            <Route path="/" element={<LandingPage />} />
            <Route path="/app" element={<AppShell />}>
              <Route index element={<HomePage />} />
              <Route path="pages" element={<PagesPage />} />
              <Route path="pages/new" element={<PageEditorPage />} />
              <Route path="pages/:pageId" element={<PageDisplayPage />} />
              <Route path="pages/:pageId/edit" element={<PageEditorPage />} />
              <Route path="settings" element={<SettingsPage />} />
            </Route>
            </Routes>
          </FormConfigProvider>
        </SettingsProvider>
      </ThemeProvider>
    </BrowserRouter>
  );
}

export default App;
