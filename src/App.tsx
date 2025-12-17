import { BrowserRouter, Routes, Route } from 'react-router-dom';
import LandingPage from './pages/LandingPage';
import AppShell from './pages/AppShell';
import HomePage from './pages/HomePage';
import SettingsPage from './pages/SettingsPage';
import DataPage from './pages/DataPage';
import ListViewPage from './pages/ListViewPage';
import { SettingsProvider } from './contexts/SettingsContext';
import { ThemeProvider } from './contexts/ThemeContext';

function App() {
  return (
    <BrowserRouter>
      <ThemeProvider>
        <SettingsProvider>
          <Routes>
            <Route path="/" element={<LandingPage />} />
            <Route path="/app" element={<AppShell />}>
              <Route index element={<HomePage />} />
              <Route path="data" element={<DataPage />} />
              <Route path="lists/:siteId/:listId" element={<ListViewPage />} />
              <Route path="settings" element={<SettingsPage />} />
            </Route>
          </Routes>
        </SettingsProvider>
      </ThemeProvider>
    </BrowserRouter>
  );
}

export default App;
