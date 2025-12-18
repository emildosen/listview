import { useCallback } from 'react';
import { makeStyles } from '@fluentui/react-components';
import type { ReportSection as ReportSectionType, SectionLayout, SectionHeight, AnyWebPartConfig } from '../../types/page';
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

/**
 * Get height value based on section height setting
 * Base height is 400px (100%)
 */
function getSectionHeight(height: SectionHeight | undefined): string {
  const baseHeight = 400; // Base height in pixels for 100%
  switch (height) {
    case 'half':
      return `${baseHeight * 0.5}px`; // 50% = 200px
    case 'medium':
      return `${baseHeight * 0.75}px`; // 75% = 300px
    case 'big':
      return `${baseHeight * 1.25}px`; // 125% = 500px
    case 'full':
    default:
      return `${baseHeight}px`; // 100% = 400px
  }
}

const useStyles = makeStyles({
  section: {
    display: 'grid',
    gap: '16px',
    marginBottom: '16px',
    alignItems: 'stretch',
  },
  webPartWrapper: {
    display: 'flex',
    flexDirection: 'column',
    overflow: 'hidden',
    height: '100%',
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

  const sectionHeight = getSectionHeight(section.height);

  return (
    <div
      className={styles.section}
      style={{
        gridTemplateColumns: getGridColumns(section.layout),
        height: sectionHeight,
      }}
    >
      {section.columns.map((column) => (
        <div key={column.id} className={styles.webPartWrapper}>
          <WebPart
            config={column.webPart}
            onConfigChange={handleConfigChange(column.id)}
          />
        </div>
      ))}
    </div>
  );
}
