import { useNavigate } from 'react-router-dom';
import {
  makeStyles,
  tokens,
  Button,
  Text,
  Title2,
} from '@fluentui/react-components';
import { WrenchScrewdriverRegular } from '@fluentui/react-icons';

const useStyles = makeStyles({
  container: {
    padding: '32px',
    flex: 1,
  },
  content: {
    maxWidth: '448px',
    margin: '0 auto',
    textAlign: 'center',
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
  const navigate = useNavigate();

  return (
    <div className={styles.container}>
      <div className={styles.content}>
        <div>
          <div className={styles.iconWrapper}>
            <WrenchScrewdriverRegular fontSize={32} className={styles.icon} />
          </div>
          <Title2 as="h1" className={styles.title}>Coming soon</Title2>
          <Text className={styles.description}>
            The app is currently under development. Check back soon for
            SharePoint list management features.
          </Text>
        </div>
        <Button appearance="outline" onClick={() => navigate('/')}>
          Back to home
        </Button>
      </div>
    </div>
  );
}

export default HomePage;
