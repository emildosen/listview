import { useNavigate } from 'react-router-dom';
import {
  makeStyles,
  mergeClasses,
  tokens,
  Button,
  Text,
  Title2,
} from '@fluentui/react-components';
import { WrenchScrewdriverRegular } from '@fluentui/react-icons';
import { useTheme } from '../contexts/ThemeContext';

const useStyles = makeStyles({
  container: {
    padding: '32px',
    flex: 1,
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
  },
  // Azure style: sharp edges, subtle shadow
  card: {
    backgroundColor: tokens.colorNeutralBackground1,
    borderRadius: '2px',
    border: '1px solid transparent',
    borderImage: 'linear-gradient(135deg, rgba(0,0,0,0.05) 0%, rgba(0,0,0,0.15) 100%) 1',
    boxShadow: '0 1px 2px rgba(0, 0, 0, 0.08), 0 2px 4px rgba(0, 0, 0, 0.04)',
    padding: '48px',
    maxWidth: '448px',
    textAlign: 'center',
  },
  cardDark: {
    backgroundColor: '#1a1a1a',
    borderImage: 'none',
    border: '1px solid #333333',
  },
  iconWrapper: {
    width: '64px',
    height: '64px',
    borderRadius: '50%',
    backgroundColor: tokens.colorBrandBackground2,
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    margin: '0 auto 16px',
  },
  icon: {
    color: tokens.colorBrandForeground1,
  },
  title: {
    marginBottom: '8px',
  },
  description: {
    color: tokens.colorNeutralForeground2,
    marginBottom: '32px',
    display: 'block',
  },
});

function HomePage() {
  const styles = useStyles();
  const { theme } = useTheme();
  const navigate = useNavigate();

  return (
    <div className={styles.container}>
      <div className={mergeClasses(styles.card, theme === 'dark' && styles.cardDark)}>
        <div className={styles.iconWrapper}>
          <WrenchScrewdriverRegular fontSize={32} className={styles.icon} />
        </div>
        <Title2 as="h1" className={styles.title}>Coming soon</Title2>
        <Text className={styles.description}>
          The app is currently under development. Check back soon for
          SharePoint list management features.
        </Text>
        <Button appearance="outline" onClick={() => navigate('/')}>
          Back to home
        </Button>
      </div>
    </div>
  );
}

export default HomePage;
