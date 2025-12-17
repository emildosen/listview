import { BrowserRouter, Routes, Route } from 'react-router-dom';
import LandingPage from './pages/LandingPage';
import AppShell from './pages/AppShell';
import HomePage from './pages/HomePage';
import SettingsPage from './pages/SettingsPage';
import DataPage from './pages/DataPage';
import ListViewPage from './pages/ListViewPage';
import ViewsPage from './pages/ViewsPage';
import ViewEditorPage from './pages/ViewEditorPage';
import ViewDisplayPage from './pages/ViewDisplayPage';
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
              <Route path="lists" element={<DataPage />} />
              <Route path="lists/:siteId/:listId" element={<ListViewPage />} />
              <Route path="views" element={<ViewsPage />} />
              <Route path="views/new" element={<ViewEditorPage />} />
              <Route path="views/:viewId" element={<ViewDisplayPage />} />
              <Route path="views/:viewId/edit" element={<ViewEditorPage />} />
              <Route path="settings" element={<SettingsPage />} />
            </Route>
          </Routes>
        </SettingsProvider>
      </ThemeProvider>
    </BrowserRouter>
  );
}

export default App;
