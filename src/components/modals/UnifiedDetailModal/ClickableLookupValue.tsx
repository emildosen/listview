import { makeStyles, tokens, Link } from '@fluentui/react-components';
import { useModalNavigation, type NavigationEntry } from './ModalNavigationContext';

const useStyles = makeStyles({
  container: {
    display: 'inline',
  },
  link: {
    color: tokens.colorBrandForeground1,
    textDecorationLine: 'none',
    cursor: 'pointer',
    '&:hover': {
      textDecorationLine: 'underline',
    },
  },
  separator: {
    color: tokens.colorNeutralForeground1,
  },
});

interface LookupValue {
  LookupId: number;
  LookupValue: string;
}

interface ClickableLookupValueProps {
  value: unknown;
  targetListId: string;
  targetListName: string;
  siteId: string;
  siteUrl?: string;
  isMultiSelect?: boolean;
}

export function ClickableLookupValue({
  value,
  targetListId,
  targetListName,
  siteId,
  siteUrl,
  isMultiSelect = false,
}: ClickableLookupValueProps) {
  const styles = useStyles();
  const { navigateToItem } = useModalNavigation();

  const handleLookupClick = (lookupId: number, e: React.MouseEvent) => {
    e.preventDefault();
    e.stopPropagation();

    const entry: NavigationEntry = {
      listId: targetListId,
      siteId,
      siteUrl,
      itemId: String(lookupId),
      listName: targetListName,
    };

    navigateToItem(entry);
  };

  // Handle null/undefined
  if (value === null || value === undefined) {
    return <span>-</span>;
  }

  // Extract lookup values
  const lookupValues: LookupValue[] = [];

  if (isMultiSelect && Array.isArray(value)) {
    for (const v of value) {
      if (typeof v === 'object' && v !== null && 'LookupId' in v && 'LookupValue' in v) {
        lookupValues.push(v as LookupValue);
      }
    }
  } else if (typeof value === 'object' && value !== null && 'LookupId' in value && 'LookupValue' in value) {
    lookupValues.push(value as LookupValue);
  }

  // No valid lookup values
  if (lookupValues.length === 0) {
    // Try to render as string if possible
    if (typeof value === 'string') {
      return <span>{value}</span>;
    }
    return <span>-</span>;
  }

  return (
    <span className={styles.container}>
      {lookupValues.map((lv, index) => (
        <span key={lv.LookupId}>
          {index > 0 && <span className={styles.separator}>, </span>}
          <Link
            className={styles.link}
            onClick={(e) => handleLookupClick(lv.LookupId, e)}
            onMouseDown={(e) => e.stopPropagation()}
          >
            {lv.LookupValue}
          </Link>
        </span>
      ))}
    </span>
  );
}
