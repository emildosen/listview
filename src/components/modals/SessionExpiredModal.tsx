import {
  Dialog,
  DialogSurface,
  DialogBody,
  DialogTitle,
  DialogContent,
  DialogActions,
  Button,
  makeStyles,
  tokens,
} from '@fluentui/react-components';
import { ErrorCircle24Regular } from '@fluentui/react-icons';
import { useAuthError } from '../../contexts/AuthErrorContext';

const useStyles = makeStyles({
  surface: {
    maxWidth: '400px',
  },
  body: {
    gap: tokens.spacingVerticalL,
  },
  header: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalM,
  },
  icon: {
    color: tokens.colorPaletteRedForeground1,
    flexShrink: 0,
  },
  title: {
    color: tokens.colorNeutralForeground1,
  },
  content: {
    color: tokens.colorNeutralForeground2,
  },
});

/**
 * Modal displayed when the SharePoint/Graph token has expired.
 * Shows an error message and a reload button to refresh the app.
 */
export function SessionExpiredModal() {
  const styles = useStyles();
  const { isSessionExpired, errorMessage } = useAuthError();

  const handleReload = () => {
    window.location.reload();
  };

  return (
    <Dialog open={isSessionExpired} modalType="alert">
      <DialogSurface className={styles.surface}>
        <DialogBody className={styles.body}>
          <DialogTitle className={styles.header}>
            <ErrorCircle24Regular className={styles.icon} />
            <span className={styles.title}>Session Expired</span>
          </DialogTitle>
          <DialogContent className={styles.content}>
            {errorMessage || 'Your SharePoint session has expired. Please reload the app to reconnect.'}
          </DialogContent>
          <DialogActions>
            <Button appearance="primary" onClick={handleReload}>
              Reload App
            </Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}
