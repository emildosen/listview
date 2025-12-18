import type { ReactNode } from 'react';
import { makeStyles, tokens } from '@fluentui/react-components';
import Sidebar from './Sidebar';

const useStyles = makeStyles({
  root: {
    minHeight: '100vh',
    backgroundColor: tokens.colorNeutralBackground1,
    display: 'flex',
    flexDirection: 'column',
  },
  main: {
    marginLeft: '256px',
    flex: 1,
    display: 'flex',
    flexDirection: 'column',
  },
});

interface AppLayoutProps {
  children: ReactNode;
}

function AppLayout({ children }: AppLayoutProps) {
  const styles = useStyles();

  return (
    <div className={styles.root}>
      <Sidebar />
      <main className={styles.main}>
        {children}
      </main>
    </div>
  );
}

export default AppLayout;
