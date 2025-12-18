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
  ChartAggregation,
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

  switch (filter.operator) {
    case 'equals':
      if (column?.boolean) {
        const boolValue =
          compareValue === true || compareValue === 'Yes' || compareValue === 'true';
        const filterBool = filterValue === 'Yes' || filterValue === true;
        return boolValue === filterBool;
      }
      return String(compareValue || '').toLowerCase() === String(filterValue).toLowerCase();

    case 'notEquals':
      if (column?.boolean) {
        const boolValue =
          compareValue === true || compareValue === 'Yes' || compareValue === 'true';
        const filterBool = filterValue === 'Yes' || filterValue === true;
        return boolValue !== filterBool;
      }
      return String(compareValue || '').toLowerCase() !== String(filterValue).toLowerCase();

    case 'contains':
      return String(compareValue || '')
        .toLowerCase()
        .includes(String(filterValue).toLowerCase());

    case 'startsWith':
      return String(compareValue || '')
        .toLowerCase()
        .startsWith(String(filterValue).toLowerCase());

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
  const { columns } = result;

  // Apply filters
  if (config.filters && config.filters.length > 0) {
    items = applyFilters(items, config.filters, columns);
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
 * Get numeric value from a field
 */
function getNumericValue(item: GraphListItem, columnName: string): number {
  const value = item.fields[columnName];
  if (value === null || value === undefined) return 0;
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
      const numValue = getNumericValue(item, config.valueColumn);
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
