import { useState } from 'react';
import { makeStyles, tokens, Link } from '@fluentui/react-components';

const useStyles = makeStyles({
  container: {
    display: 'flex',
    flexDirection: 'column',
    gap: '4px',
  },
  text: {
    whiteSpace: 'pre-wrap',
    wordBreak: 'break-word',
  },
  toggle: {
    fontSize: tokens.fontSizeBase200,
  },
});

interface TruncatedRichTextProps {
  value: string;
  maxLength?: number;
}

/**
 * Strips HTML tags from a string and decodes HTML entities.
 */
function stripHtml(html: string): string {
  if (!html) return '';

  // Create a temporary element to decode HTML entities and strip tags
  const doc = new DOMParser().parseFromString(html, 'text/html');
  return doc.body.textContent || '';
}

export function TruncatedRichText({ value, maxLength = 180 }: TruncatedRichTextProps) {
  const styles = useStyles();
  const [expanded, setExpanded] = useState(false);

  // Strip HTML and get plain text
  const plainText = stripHtml(value).trim();

  // Check if truncation is needed
  const needsTruncation = plainText.length > maxLength;

  // Get display text
  const displayText = expanded || !needsTruncation
    ? plainText
    : plainText.slice(0, maxLength).trimEnd() + '...';

  if (!plainText) {
    return <span>-</span>;
  }

  return (
    <div className={styles.container}>
      <span className={styles.text}>{displayText}</span>
      {needsTruncation && (
        <Link
          className={styles.toggle}
          onClick={(e) => {
            e.stopPropagation();
            setExpanded(!expanded);
          }}
        >
          {expanded ? 'Show less' : 'Show more'}
        </Link>
      )}
    </div>
  );
}
