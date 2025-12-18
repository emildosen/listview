import { useCallback } from 'react';
import { makeStyles } from '@fluentui/react-components';
import type { ReportSection as ReportSectionType, SectionLayout, AnyWebPartConfig } from '../../types/page';
import WebPart from './WebPart';

/**
 * Get CSS grid-template-columns value for each layout type
 */
function getGridColumns(layout: SectionLayout): string {
  switch (layout) {
    case 'one-column':
      return '1fr';
    case 'two-column':
      return '1fr 1fr';
    case 'three-column':
      return '1fr 1fr 1fr';
    case 'one-third-left':
      return '1fr 2fr';
    case 'one-third-right':
      return '2fr 1fr';
    default:
      return '1fr';
  }
}

const useStyles = makeStyles({
  section: {
    display: 'grid',
    gap: '16px',
    marginBottom: '16px',
  },
});

interface ReportSectionProps {
  section: ReportSectionType;
  onWebPartConfigChange?: (columnId: string, config: AnyWebPartConfig) => void;
}

export default function ReportSection({ section, onWebPartConfigChange }: ReportSectionProps) {
  const styles = useStyles();

  const handleConfigChange = useCallback(
    (columnId: string) => (config: AnyWebPartConfig) => {
      onWebPartConfigChange?.(columnId, config);
    },
    [onWebPartConfigChange]
  );

  return (
    <div
      className={styles.section}
      style={{ gridTemplateColumns: getGridColumns(section.layout) }}
    >
      {section.columns.map((column) => (
        <WebPart
          key={column.id}
          config={column.webPart}
          onConfigChange={handleConfigChange(column.id)}
        />
      ))}
    </div>
  );
}
