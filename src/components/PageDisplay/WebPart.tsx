import {
  makeStyles,
  tokens,
  Text,
  mergeClasses,
} from '@fluentui/react-components';
import { AppsAddInRegular } from '@fluentui/react-icons';
import { useTheme } from '../../contexts/ThemeContext';
import type { AnyWebPartConfig } from '../../types/page';
import ListItemsWebPart from './WebParts/ListItemsWebPart';
import ChartWebPart from './WebParts/ChartWebPart';

const useStyles = makeStyles({
  emptyContainer: {
    backgroundColor: tokens.colorNeutralBackground1,
    borderRadius: '2px',
    border: '1px solid transparent',
    borderImage: 'linear-gradient(135deg, rgba(0,0,0,0.05) 0%, rgba(0,0,0,0.15) 100%) 1',
    boxShadow: '0 1px 2px rgba(0, 0, 0, 0.08), 0 2px 4px rgba(0, 0, 0, 0.04)',
    padding: '48px 24px',
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    justifyContent: 'center',
    gap: '12px',
    height: '100%',
  },
  emptyContainerDark: {
    backgroundColor: '#1a1a1a',
    borderImage: 'none',
    border: '1px solid #333333',
  },
  emptyIcon: {
    color: tokens.colorNeutralForeground3,
    fontSize: '32px',
  },
  emptyText: {
    color: tokens.colorNeutralForeground3,
    textAlign: 'center',
  },
});

interface WebPartProps {
  config: AnyWebPartConfig | null;
  onConfigChange?: (config: AnyWebPartConfig) => void;
}

export default function WebPart({ config, onConfigChange }: WebPartProps) {
  const { theme } = useTheme();
  const styles = useStyles();

  // Empty state when no WebPart assigned
  if (!config) {
    return (
      <div className={mergeClasses(styles.emptyContainer, theme === 'dark' && styles.emptyContainerDark)}>
        <AppsAddInRegular className={styles.emptyIcon} />
        <Text className={styles.emptyText}>
          Click Customize to add a web part
        </Text>
      </div>
    );
  }

  // Render the appropriate WebPart based on type
  switch (config.type) {
    case 'list-items':
      return <ListItemsWebPart config={config} onConfigChange={onConfigChange} />;
    case 'chart':
      return <ChartWebPart config={config} onConfigChange={onConfigChange} />;
    default:
      return (
        <div className={mergeClasses(styles.emptyContainer, theme === 'dark' && styles.emptyContainerDark)}>
          <Text className={styles.emptyText}>
            Unknown web part type
          </Text>
        </div>
      );
  }
}
