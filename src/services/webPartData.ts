import type { IPublicClientApplication, AccountInfo } from '@azure/msal-browser';
import {
  getListItems,
  getListColumns,
  type GraphListColumn,
  type GraphListItem,
} from '../auth/graphClient';
import type {
  ListItemsWebPartConfig,
  ChartWebPartConfig,
  WebPartFilter,
  WebPartJoin,
  ChartAggregation,
  JoinColumnConfig,
  JoinColumnAggregation,
} from '../types/page';

export interface WebPartDataResult {
  items: GraphListItem[];
  columns: GraphListColumn[];
  totalCount: number;
}

export interface ChartDataPoint {
  legend: string;
  data: number;
  color?: string;
}

// Default color palette for charts
const DEFAULT_COLORS = [
  '#0078d4', // Blue
  '#00bcf2', // Light Blue
  '#8764b8', // Purple
  '#e3008c', // Magenta
  '#d13438', // Red
  '#ff8c00', // Orange
  '#107c10', // Green
  '#038387', // Teal
  '#5c2d91', // Dark Purple
  '#ca5010', // Dark Orange
];

/**
 * Apply filter to a single item
 */
function matchesFilter(
  item: GraphListItem,
  filter: WebPartFilter,
  columns: GraphListColumn[]
): boolean {
  const fieldValue = item.fields[filter.column];
  const column = columns.find((c) => c.name === filter.column);

  // Handle lookup columns - get the display value
  let compareValue: unknown = fieldValue;
  if (column?.lookup && typeof fieldValue === 'object' && fieldValue !== null) {
    compareValue = (fieldValue as { LookupValue?: string }).LookupValue;
  }

  const filterValue = filter.value;
  const filterValueLower = String(filterValue).toLowerCase();

  // Helper for choice column comparison - handles both single and multi-value
  const choiceMatches = (operator: 'equals' | 'notEquals' | 'contains'): boolean => {
    // Choice columns can be arrays (multi-value) or strings (single-value)
    const values = Array.isArray(compareValue)
      ? compareValue.map((v) => String(v).toLowerCase())
      : [String(compareValue || '').toLowerCase()];

    switch (operator) {
      case 'equals':
        // For equals, check if any value in the array matches exactly
        return values.some((v) => v === filterValueLower);
      case 'notEquals':
        // For notEquals, ensure none of the values match
        return !values.some((v) => v === filterValueLower);
      case 'contains':
        // For contains, check if any value contains the filter
        return values.some((v) => v.includes(filterValueLower));
      default:
        return false;
    }
  };

  switch (filter.operator) {
    case 'equals':
      if (column?.boolean) {
        const boolValue =
          compareValue === true || compareValue === 'Yes' || compareValue === 'true';
        const filterBool = filterValue === 'Yes' || filterValue === true;
        return boolValue === filterBool;
      }
      // Handle choice columns (both single and multi-value)
      if (column?.choice || Array.isArray(compareValue)) {
        return choiceMatches('equals');
      }
      return String(compareValue || '').toLowerCase() === filterValueLower;

    case 'notEquals':
      if (column?.boolean) {
        const boolValue =
          compareValue === true || compareValue === 'Yes' || compareValue === 'true';
        const filterBool = filterValue === 'Yes' || filterValue === true;
        return boolValue !== filterBool;
      }
      // Handle choice columns (both single and multi-value)
      if (column?.choice || Array.isArray(compareValue)) {
        return choiceMatches('notEquals');
      }
      return String(compareValue || '').toLowerCase() !== filterValueLower;

    case 'contains':
      // Handle choice columns (both single and multi-value)
      if (column?.choice || Array.isArray(compareValue)) {
        return choiceMatches('contains');
      }
      return String(compareValue || '')
        .toLowerCase()
        .includes(filterValueLower);

    case 'startsWith':
      return String(compareValue || '')
        .toLowerCase()
        .startsWith(filterValueLower);

    case 'greaterThan':
      if (column?.dateTime) {
        return new Date(String(compareValue)) > new Date(String(filterValue));
      }
      return Number(compareValue) > Number(filterValue);

    case 'lessThan':
      if (column?.dateTime) {
        return new Date(String(compareValue)) < new Date(String(filterValue));
      }
      return Number(compareValue) < Number(filterValue);

    case 'isEmpty':
      return (
        compareValue === null ||
        compareValue === undefined ||
        compareValue === '' ||
        (Array.isArray(compareValue) && compareValue.length === 0)
      );

    case 'isNotEmpty':
      return (
        compareValue !== null &&
        compareValue !== undefined &&
        compareValue !== '' &&
        !(Array.isArray(compareValue) && compareValue.length === 0)
      );

    default:
      return true;
  }
}

/**
 * Apply all filters to items
 */
function applyFilters(
  items: GraphListItem[],
  filters: WebPartFilter[],
  columns: GraphListColumn[]
): GraphListItem[] {
  if (!filters || filters.length === 0) {
    return items;
  }

  return items.filter((item) => {
    let result = true;

    for (let i = 0; i < filters.length; i++) {
      const filter = filters[i];
      const matches = matchesFilter(item, filter, columns);

      if (i === 0) {
        result = matches;
      } else {
        if (filter.conjunction === 'and') {
          result = result && matches;
        } else {
          result = result || matches;
        }
      }
    }

    return result;
  });
}

/**
 * Sort items by a column
 */
function sortItems(
  items: GraphListItem[],
  sortColumn: string,
  sortDirection: 'asc' | 'desc',
  columns: GraphListColumn[]
): GraphListItem[] {
  const column = columns.find((c) => c.name === sortColumn);

  return [...items].sort((a, b) => {
    let aVal = a.fields[sortColumn];
    let bVal = b.fields[sortColumn];

    // Handle lookup columns
    if (column?.lookup) {
      aVal = (aVal as { LookupValue?: string })?.LookupValue;
      bVal = (bVal as { LookupValue?: string })?.LookupValue;
    }

    // Handle null/undefined
    if (aVal === null || aVal === undefined) aVal = '';
    if (bVal === null || bVal === undefined) bVal = '';

    // Compare
    let comparison: number;
    if (column?.number) {
      comparison = Number(aVal) - Number(bVal);
    } else if (column?.dateTime) {
      comparison = new Date(String(aVal)).getTime() - new Date(String(bVal)).getTime();
    } else {
      comparison = String(aVal).localeCompare(String(bVal));
    }

    return sortDirection === 'desc' ? -comparison : comparison;
  });
}

/**
 * Fetch data for a List Items web part
 */
export async function fetchListWebPartData(
  instance: IPublicClientApplication,
  account: AccountInfo,
  config: ListItemsWebPartConfig
): Promise<WebPartDataResult> {
  if (!config.dataSource?.siteId || !config.dataSource?.listId) {
    return { items: [], columns: [], totalCount: 0 };
  }

  const { siteId, listId } = config.dataSource;

  // Fetch items and columns
  const result = await getListItems(instance, account, siteId, listId);
  let { items } = result;
  let { columns } = result;

  // Apply filters
  if (config.filters && config.filters.length > 0) {
    items = applyFilters(items, config.filters, columns);
  }

  // Execute joins if configured
  if (config.joins && config.joins.length > 0) {
    const joinResult = await executeJoins(instance, account, items, config.joins, columns);
    items = joinResult.items;
    columns = joinResult.columns;
  }

  const totalCount = items.length;

  // Apply sort
  if (config.sort?.column) {
    items = sortItems(items, config.sort.column, config.sort.direction, columns);
  }

  // Apply max items limit
  const maxItems = config.maxItems || 50;
  items = items.slice(0, maxItems);

  return { items, columns, totalCount };
}

/**
 * Get display value from a field
 */
function getDisplayValue(item: GraphListItem, columnName: string, columns: GraphListColumn[]): string {
  const value = item.fields[columnName];
  const column = columns.find((c) => c.name === columnName);

  if (value === null || value === undefined) return '';

  // Handle lookup columns
  if (column?.lookup && typeof value === 'object') {
    return (value as { LookupValue?: string }).LookupValue || '';
  }

  // Handle boolean
  if (column?.boolean) {
    return value ? 'Yes' : 'No';
  }

  return String(value);
}

/**
 * Get numeric value from a field, handling boolean columns as 1/0
 */
function getNumericValue(item: GraphListItem, columnName: string, column?: GraphListColumn): number {
  const value = item.fields[columnName];
  if (value === null || value === undefined) return 0;

  // Handle boolean columns - Yes/true = 1, No/false = 0
  if (column?.boolean) {
    return value === true || value === 'Yes' || value === 'true' ? 1 : 0;
  }
  // Handle string "Yes"/"No" even without column metadata
  if (value === true || value === 'Yes' || value === 'true') return 1;
  if (value === false || value === 'No' || value === 'false') return 0;

  const num = Number(value);
  return isNaN(num) ? 0 : num;
}

/**
 * Aggregate values
 */
function aggregate(values: number[], aggregation: ChartAggregation): number {
  if (values.length === 0) return 0;

  switch (aggregation) {
    case 'count':
      return values.length;
    case 'sum':
      return values.reduce((a, b) => a + b, 0);
    case 'average':
      return values.reduce((a, b) => a + b, 0) / values.length;
    case 'min':
      return Math.min(...values);
    case 'max':
      return Math.max(...values);
    default:
      return values.length;
  }
}

/**
 * Fetch and aggregate data for a Chart web part
 */
export async function fetchChartWebPartData(
  instance: IPublicClientApplication,
  account: AccountInfo,
  config: ChartWebPartConfig
): Promise<ChartDataPoint[]> {
  if (!config.dataSource?.siteId || !config.dataSource?.listId) {
    return [];
  }

  if (!config.groupByColumn) {
    return [];
  }

  const { siteId, listId } = config.dataSource;

  // Fetch items and columns
  const result = await getListItems(instance, account, siteId, listId);
  let { items } = result;
  const { columns } = result;

  // Apply filters
  if (config.filters && config.filters.length > 0) {
    items = applyFilters(items, config.filters, columns);
  }

  // Group items by the groupByColumn
  const groups = new Map<string, number[]>();

  for (const item of items) {
    const groupKey = getDisplayValue(item, config.groupByColumn, columns) || '(Empty)';

    if (!groups.has(groupKey)) {
      groups.set(groupKey, []);
    }

    // For count, we just track presence; for other aggregations, track the value column
    if (config.aggregation === 'count') {
      groups.get(groupKey)!.push(1);
    } else if (config.valueColumn) {
      const valueCol = columns.find((c) => c.name === config.valueColumn);
      const numValue = getNumericValue(item, config.valueColumn, valueCol);
      groups.get(groupKey)!.push(numValue);
    }
  }

  // Convert to chart data points
  let dataPoints: ChartDataPoint[] = [];

  let colorIndex = 0;
  const colors = config.colorPalette || DEFAULT_COLORS;

  for (const [legend, values] of groups.entries()) {
    const aggregatedValue = aggregate(values, config.aggregation || 'count');
    dataPoints.push({
      legend,
      data: Math.round(aggregatedValue * 100) / 100, // Round to 2 decimal places
      color: colors[colorIndex % colors.length],
    });
    colorIndex++;
  }

  // Sort data points
  if (config.sortBy === 'value') {
    dataPoints.sort((a, b) =>
      config.sortDirection === 'desc' ? b.data - a.data : a.data - b.data
    );
  } else {
    // Sort by label
    dataPoints.sort((a, b) =>
      config.sortDirection === 'desc'
        ? b.legend.localeCompare(a.legend)
        : a.legend.localeCompare(b.legend)
    );
  }

  // Limit to max groups
  const maxGroups = config.maxGroups || 10;
  if (dataPoints.length > maxGroups) {
    const topGroups = dataPoints.slice(0, maxGroups - 1);
    const otherGroups = dataPoints.slice(maxGroups - 1);
    const otherTotal = otherGroups.reduce((sum, dp) => sum + dp.data, 0);
    topGroups.push({
      legend: 'Other',
      data: Math.round(otherTotal * 100) / 100,
      color: colors[(maxGroups - 1) % colors.length],
    });
    dataPoints = topGroups;
  }

  return dataPoints;
}

/**
 * Get column schema for a data source
 */
export async function getDataSourceColumns(
  instance: IPublicClientApplication,
  account: AccountInfo,
  siteId: string,
  listId: string
): Promise<GraphListColumn[]> {
  return getListColumns(instance, account, siteId, listId);
}

/**
 * Execute joins to merge data from related lists
 */
export async function executeJoins(
  instance: IPublicClientApplication,
  account: AccountInfo,
  primaryItems: GraphListItem[],
  joins: WebPartJoin[],
  primaryColumns: GraphListColumn[]
): Promise<{ items: GraphListItem[]; columns: GraphListColumn[] }> {
  if (!joins || joins.length === 0) {
    return { items: primaryItems, columns: primaryColumns };
  }

  let resultItems = [...primaryItems];
  const resultColumns = [...primaryColumns];

  for (const join of joins) {
    if (!join.targetSource?.siteId || !join.targetSource?.listId) {
      continue;
    }

    try {
      // Fetch target list data
      const targetResult = await getListItems(
        instance,
        account,
        join.targetSource.siteId,
        join.targetSource.listId
      );

      const targetItems = targetResult.items;
      const targetColumns = targetResult.columns;

      // Get the source column definition to determine if it's a lookup
      const sourceColumn = primaryColumns.find((c) => c.name === join.sourceColumn);
      // Get the target column definition to determine if it's a lookup (for reverse joins)
      const targetColumn = targetColumns.find((c) => c.name === join.targetColumn);

      // Build a map of target items by the join column
      // For reverse joins, we may have multiple target items per key (one-to-many)
      const targetMap = new Map<string, GraphListItem[]>();
      for (const targetItem of targetItems) {
        let keyValue: unknown;

        if (join.targetColumn === 'id') {
          keyValue = targetItem.id;
        } else if (targetColumn?.lookup) {
          // For lookup columns in target, get the LookupId
          const lookupIdField = `${join.targetColumn}LookupId`;
          keyValue = targetItem.fields[lookupIdField];
        } else {
          keyValue = targetItem.fields[join.targetColumn];
        }

        if (keyValue !== null && keyValue !== undefined) {
          const key = String(keyValue);
          if (!targetMap.has(key)) {
            targetMap.set(key, []);
          }
          targetMap.get(key)!.push(targetItem);
        }
      }

      // Build column configs - use columnConfigs if available, otherwise create from columnsToInclude
      const columnConfigs: JoinColumnConfig[] = join.columnConfigs?.length
        ? join.columnConfigs
        : join.columnsToInclude.map((name) => ({
            columnName: name,
            displayName: targetColumns.find((c) => c.name === name)?.displayName || name,
            aggregation: 'first' as JoinColumnAggregation,
          }));

      // Add target columns to result columns (with alias prefix if specified)
      for (const config of columnConfigs) {
        const col = targetColumns.find((c) => c.name === config.columnName);
        if (!col) continue;

        const aliasName = join.alias
          ? `${join.alias}${col.name}`
          : `${join.targetSource.listName}_${col.name}`;

        // Use custom display name from config, or generate default
        const aliasDisplayName = config.displayName
          || (join.alias ? `${join.alias}${col.displayName}` : `${join.targetSource.listName} - ${col.displayName}`);

        // When aggregation produces a number (not 'first'), override column type
        const isNumericAggregation = config.aggregation && config.aggregation !== 'first';

        resultColumns.push({
          ...col,
          name: aliasName,
          displayName: aliasDisplayName,
          // Clear boolean flag and set number flag when aggregating
          ...(isNumericAggregation ? { boolean: undefined, number: {} } : {}),
        });
      }

      // Helper to convert value to number, handling boolean/Yes/No columns
      const toNumericValue = (v: unknown, column?: GraphListColumn): number => {
        if (v === null || v === undefined) return 0;
        // Handle boolean columns - Yes/true = 1, No/false = 0
        if (column?.boolean) {
          return v === true || v === 'Yes' || v === 'true' ? 1 : 0;
        }
        // Handle string "Yes"/"No" even without column metadata
        if (v === true || v === 'Yes' || v === 'true') return 1;
        if (v === false || v === 'No' || v === 'false') return 0;
        const num = Number(v);
        return isNaN(num) ? 0 : num;
      };

      // Helper to apply aggregation to values
      const applyAggregation = (values: unknown[], aggregation: JoinColumnAggregation, column?: GraphListColumn): unknown => {
        if (values.length === 0) return null;

        switch (aggregation) {
          case 'first':
            return values[0];
          case 'count':
            return values.filter((v) => v !== null && v !== undefined).length;
          case 'sum': {
            const nums = values.map((v) => toNumericValue(v, column));
            return nums.reduce((a, b) => a + b, 0);
          }
          case 'avg': {
            const nums = values.map((v) => toNumericValue(v, column));
            return nums.length > 0 ? nums.reduce((a, b) => a + b, 0) / nums.length : 0;
          }
          case 'min': {
            const nums = values.map((v) => toNumericValue(v, column)).filter((n) => !isNaN(n));
            return nums.length > 0 ? Math.min(...nums) : null;
          }
          case 'max': {
            const nums = values.map((v) => toNumericValue(v, column)).filter((n) => !isNaN(n));
            return nums.length > 0 ? Math.max(...nums) : null;
          }
          default:
            return values[0];
        }
      };

      // Merge data
      const mergedItems: GraphListItem[] = [];

      for (const primaryItem of resultItems) {
        // Get the join key from the primary item
        let joinKey: string | null = null;

        if (join.sourceColumn === 'id') {
          // For reverse joins, use the primary item's ID
          joinKey = primaryItem.id;
        } else if (sourceColumn?.lookup) {
          // For lookup columns, get the LookupId
          const lookupIdField = `${join.sourceColumn}LookupId`;
          const lookupId = primaryItem.fields[lookupIdField];
          if (lookupId !== null && lookupId !== undefined) {
            joinKey = String(lookupId);
          }
        } else {
          // For non-lookup columns, use the value directly
          const value = primaryItem.fields[join.sourceColumn];
          if (value !== null && value !== undefined) {
            joinKey = String(value);
          }
        }

        // Find matching target items (may be multiple for reverse joins)
        const matchingTargets = joinKey ? targetMap.get(joinKey) : null;
        const hasMatches = matchingTargets && matchingTargets.length > 0;

        if (hasMatches || join.joinType === 'left') {
          // Merge the fields with aggregation
          const mergedFields = { ...primaryItem.fields };

          for (const config of columnConfigs) {
            const col = targetColumns.find((c) => c.name === config.columnName);
            if (!col) continue;

            const aliasName = join.alias
              ? `${join.alias}${col.name}`
              : `${join.targetSource.listName}_${col.name}`;

            if (hasMatches) {
              // Collect values from all matching items
              const values = matchingTargets!.map((item) => item.fields[config.columnName]);
              // Apply aggregation (pass column for boolean handling)
              mergedFields[aliasName] = applyAggregation(values, config.aggregation, col);
            } else {
              mergedFields[aliasName] = null;
            }
          }

          mergedItems.push({
            ...primaryItem,
            fields: mergedFields,
          });
        }
        // For inner join, items without a match are excluded
      }

      resultItems = mergedItems;
    } catch (err) {
      console.error(`Failed to execute join with ${join.targetSource.listName}:`, err);
      // Continue with other joins even if one fails
    }
  }

  return { items: resultItems, columns: resultColumns };
}
