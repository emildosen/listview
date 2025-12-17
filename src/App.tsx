import { BrowserRouter, Routes, Route } from 'react-router-dom';
import LandingPage from './pages/LandingPage';
import AppPage from './pages/AppPage';
import SettingsPage from './pages/SettingsPage';
import { SettingsProvider } from './contexts/SettingsContext';
import { ThemeProvider } from './contexts/ThemeContext';

function App() {
  return (
    <BrowserRouter>
      <ThemeProvider>
        <SettingsProvider>
          <Routes>
            <Route path="/" element={<LandingPage />} />
            <Route path="/app" element={<AppPage />} />
            <Route path="/settings" element={<SettingsPage />} />
          </Routes>
        </SettingsProvider>
      </ThemeProvider>
    </BrowserRouter>
  );
}

export default App;
