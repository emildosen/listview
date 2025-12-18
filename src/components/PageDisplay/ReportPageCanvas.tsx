import { useCallback } from 'react';
import {
  makeStyles,
  tokens,
  Text,
  mergeClasses,
} from '@fluentui/react-components';
import { DocumentBulletListRegular } from '@fluentui/react-icons';
import { useTheme } from '../../contexts/ThemeContext';
import type { ReportLayoutConfig, AnyWebPartConfig } from '../../types/page';
import ReportSection from './ReportSection';

const useStyles = makeStyles({
  canvas: {
    display: 'flex',
    flexDirection: 'column',
    gap: '0',
  },
  emptyState: {
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    justifyContent: 'center',
    padding: '64px 24px',
    gap: '16px',
    backgroundColor: tokens.colorNeutralBackground1,
    borderRadius: '2px',
    border: '1px solid transparent',
    borderImage: 'linear-gradient(135deg, rgba(0,0,0,0.05) 0%, rgba(0,0,0,0.15) 100%) 1',
    boxShadow: '0 1px 2px rgba(0, 0, 0, 0.08), 0 2px 4px rgba(0, 0, 0, 0.04)',
  },
  emptyStateDark: {
    backgroundColor: '#1a1a1a',
    borderImage: 'none',
    border: '1px solid #333333',
  },
  emptyIcon: {
    color: tokens.colorNeutralForeground3,
    fontSize: '48px',
  },
  emptyText: {
    color: tokens.colorNeutralForeground3,
    textAlign: 'center',
  },
});

interface ReportPageCanvasProps {
  layout: ReportLayoutConfig;
  onWebPartConfigChange?: (sectionId: string, columnId: string, config: AnyWebPartConfig) => void;
}

export default function ReportPageCanvas({ layout, onWebPartConfigChange }: ReportPageCanvasProps) {
  const { theme } = useTheme();
  const styles = useStyles();

  const handleSectionConfigChange = useCallback(
    (sectionId: string) => (columnId: string, config: AnyWebPartConfig) => {
      onWebPartConfigChange?.(sectionId, columnId, config);
    },
    [onWebPartConfigChange]
  );

  // Empty state when no sections
  if (!layout.sections || layout.sections.length === 0) {
    return (
      <div className={mergeClasses(styles.emptyState, theme === 'dark' && styles.emptyStateDark)}>
        <DocumentBulletListRegular className={styles.emptyIcon} />
        <Text className={styles.emptyText}>
          Click Customize to start building your report page
        </Text>
      </div>
    );
  }

  return (
    <div className={styles.canvas}>
      {layout.sections.map((section) => (
        <ReportSection
          key={section.id}
          section={section}
          onWebPartConfigChange={handleSectionConfigChange(section.id)}
        />
      ))}
    </div>
  );
}
