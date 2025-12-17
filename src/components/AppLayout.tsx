import type { ReactNode } from 'react';
import Sidebar from './Sidebar';

interface AppLayoutProps {
  children: ReactNode;
}

function AppLayout({ children }: AppLayoutProps) {
  return (
    <div className="min-h-screen bg-base-100">
      <Sidebar />
      <main className="ml-64">
        {children}
      </main>
    </div>
  );
}

export default AppLayout;
