import {
  makeStyles,
  tokens,
  Text,
  Button,
  Badge,
  mergeClasses,
} from '@fluentui/react-components';
import { SettingsRegular, WarningRegular } from '@fluentui/react-icons';
import { useTheme } from '../../../contexts/ThemeContext';

const useStyles = makeStyles({
  header: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    padding: '12px 16px',
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    gap: '8px',
  },
  headerDark: {
    borderBottom: '1px solid #333333',
  },
  titleSection: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
    flex: 1,
    minWidth: 0,
  },
  title: {
    fontWeight: tokens.fontWeightSemibold,
    fontSize: tokens.fontSizeBase400,
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap',
  },
  notConfiguredBadge: {
    backgroundColor: tokens.colorPaletteYellowBackground2,
    color: tokens.colorPaletteYellowForeground2,
  },
  actions: {
    display: 'flex',
    alignItems: 'center',
    gap: '4px',
    flexShrink: 0,
  },
});

interface WebPartHeaderProps {
  title?: string;
  isConfigured: boolean;
  onSettingsClick: () => void;
}

export default function WebPartHeader({
  title,
  isConfigured,
  onSettingsClick,
}: WebPartHeaderProps) {
  const { theme } = useTheme();
  const styles = useStyles();

  return (
    <div className={mergeClasses(styles.header, theme === 'dark' && styles.headerDark)}>
      <div className={styles.titleSection}>
        {title && <Text className={styles.title}>{title}</Text>}
        {!isConfigured && (
          <Badge
            appearance="filled"
            className={styles.notConfiguredBadge}
            icon={<WarningRegular />}
            size="small"
          >
            Not configured
          </Badge>
        )}
      </div>
      <div className={styles.actions}>
        <Button
          appearance="subtle"
          icon={<SettingsRegular />}
          size="small"
          onClick={onSettingsClick}
          title="Configure web part"
        />
      </div>
    </div>
  );
}
