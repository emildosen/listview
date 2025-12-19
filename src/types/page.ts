// Page type definitions for the Custom Pages feature

// Detail modal layout configuration (for customizable popup)
export interface DetailLayoutConfig {
  columnSettings: DetailColumnSetting[];
  // Section order: 'details', 'description', and linked list IDs (e.g., 'section-123')
  // Default order: ['details', 'description', ...linkedListIds]
  sectionOrder?: string[];
  /** @deprecated Use sectionOrder instead */
  relatedSectionOrder?: string[];  // Legacy: Array of RelatedSection.id values for ordering
}

export interface DetailColumnSetting {
  internalName: string;
  visible: boolean;
  displayStyle: 'stat' | 'list' | 'description';  // 'stat' = badge at top, 'list' = detail grid, 'description' = large text area
}

// Per-list detail popup configuration
// This is stored per list (by listId) so the same popup appears regardless of which page displays the list
export interface ListDetailConfig {
  listId: string;
  listName: string;
  siteId: string;
  siteUrl?: string;
  displayColumns: PageColumn[];        // Columns available in detail view
  detailLayout: DetailLayoutConfig;    // How columns are displayed (stat/list, order, visibility)
  relatedSections: RelatedSection[];   // Related lists to show in detail popup
}

export type PageType = 'lookup' | 'report';

export interface PageDefinition {
  id?: string;                        // SharePoint item ID
  name: string;                       // Display name (e.g., "Student Details")
  description?: string;               // Optional description
  pageType: PageType;                 // Type of page: lookup (data view) or report (dashboard)

  // Primary entity configuration (not used for report pages)
  primarySource: PageSource;          // Main list (e.g., Students)
  displayColumns: PageColumn[];       // Columns shown in detail view

  // Search/filter configuration
  searchConfig: SearchConfig;

  // Related lists configuration
  relatedSections: RelatedSection[];  // Related lists to show (e.g., Correspondence)

  // Detail modal layout configuration
  detailLayout?: DetailLayoutConfig;

  // Report page layout configuration (only used when pageType === 'report')
  reportLayout?: ReportLayoutConfig;

  createdAt?: string;
  updatedAt?: string;
}

export interface PageSource {
  siteId: string;
  siteUrl?: string;                   // Site URL for SP client creation
  listId: string;
  listName: string;                   // For display
}

export interface PageColumn {
  internalName: string;               // SharePoint column internal name
  displayName: string;                // Custom display name
  editable?: boolean;                 // Allow editing (default: true for non-readonly)
}

export interface SearchConfig {
  tableColumns?: PageColumn[];        // Columns shown in table view
  textSearchColumns: string[];        // Columns to search with text input
  filterColumns: FilterColumn[];      // Dropdown filter columns
}

export interface FilterColumn {
  internalName: string;
  displayName: string;
  type: 'choice' | 'lookup' | 'boolean';
  // For choice columns, options are auto-loaded from SharePoint column definition
  // For lookup columns, options are loaded from target list
}

export interface RelatedSection {
  id: string;                         // Unique section ID
  title: string;                      // Section header (e.g., "Correspondence")
  source: PageSource;                 // Related list details
  lookupColumn: string;               // Column in related list that links to primary
  displayColumns: PageColumn[];       // Columns to show in related items table
  defaultSort?: {
    column: string;
    direction: 'asc' | 'desc';
  };
}

// SharePoint list item representation
export interface PageItem {
  Id: number;
  Title: string;
  PageConfig: string;                 // JSON stringified PageDefinition
}

// ===== REPORT PAGE LAYOUT TYPES =====

/**
 * Layout options for sections - mirrors SharePoint's section layouts
 */
export type SectionLayout =
  | 'one-column'          // Full width
  | 'two-column'          // 50/50
  | 'three-column'        // 33/33/33
  | 'one-third-left'      // 1/3 + 2/3
  | 'one-third-right';    // 2/3 + 1/3

/**
 * Column width percentages for each layout type
 */
export const LAYOUT_COLUMN_WIDTHS: Record<SectionLayout, number[]> = {
  'one-column': [100],
  'two-column': [50, 50],
  'three-column': [33.33, 33.33, 33.33],
  'one-third-left': [33.33, 66.67],
  'one-third-right': [66.67, 33.33],
};

/**
 * Available WebPart types - extensible for future additions
 */
export type WebPartType = 'list-items' | 'chart';

/**
 * Base WebPart configuration
 */
export interface WebPartConfig {
  id: string;
  type: WebPartType;
  title?: string;
}

/**
 * Data source configuration for a web part
 */
export interface WebPartDataSource {
  siteId: string;
  siteUrl?: string;
  listId: string;
  listName: string;
}

/**
 * Filter operators for web part data
 */
export type WebPartFilterOperator =
  | 'equals'
  | 'notEquals'
  | 'contains'
  | 'startsWith'
  | 'greaterThan'
  | 'lessThan'
  | 'isEmpty'
  | 'isNotEmpty';

/**
 * Filter condition for web part data
 */
export interface WebPartFilter {
  id: string;
  column: string;
  operator: WebPartFilterOperator;
  value: string | number | boolean;
  conjunction: 'and' | 'or'; // How this filter combines with previous
}

/**
 * Aggregation options for joined columns
 */
export type JoinColumnAggregation = 'first' | 'count' | 'sum' | 'avg' | 'min' | 'max';

/**
 * Configuration for a column included in a join
 */
export interface JoinColumnConfig {
  columnName: string; // Internal column name from joined list
  displayName?: string; // Custom display name (optional)
  aggregation: JoinColumnAggregation; // How to aggregate when multiple matches
}

/**
 * Join configuration to link with another list
 */
export interface WebPartJoin {
  id: string;
  targetSource: WebPartDataSource; // The list to join with
  sourceColumn: string; // Column in primary list (lookup or ID)
  targetColumn: string; // Column in target list to match
  joinType: 'inner' | 'left'; // Inner = only matching, Left = all primary + matching
  columnsToInclude: string[]; // Legacy: simple column names (kept for backward compat)
  columnConfigs?: JoinColumnConfig[]; // Enhanced column configuration with aggregation
  alias?: string; // Prefix for joined columns (e.g., "Contact.")
}

/**
 * Sort configuration
 */
export interface WebPartSort {
  column: string;
  direction: 'asc' | 'desc';
}

/**
 * Display column configuration for web parts
 */
export interface WebPartDisplayColumn {
  internalName: string;
  displayName: string;
  width?: number; // Column width in px
  format?: 'text' | 'number' | 'date' | 'currency' | 'boolean' | 'lookup';
}

/**
 * Chart aggregation types
 */
export type ChartAggregation = 'count' | 'sum' | 'average' | 'min' | 'max';

/**
 * Legend visibility options for charts
 */
export type LegendPosition = 'on' | 'off';

/**
 * X-axis label style for bar charts
 */
export type XAxisLabelStyle = 'normal' | 'angled';

/**
 * List Items WebPart - displays a table of data
 */
export interface ListItemsWebPartConfig extends WebPartConfig {
  type: 'list-items';
  dataSource?: WebPartDataSource;
  displayColumns?: WebPartDisplayColumn[]; // Columns to show in table
  filters?: WebPartFilter[];
  joins?: WebPartJoin[];
  sort?: WebPartSort;
  maxItems?: number; // Limit rows (default: 50)
  showSearch?: boolean; // Enable text search
  searchColumns?: string[]; // Columns to search
}

/**
 * Chart WebPart - displays a visualization
 */
/**
 * Available chart types
 */
export type ChartType = 'bar' | 'donut' | 'line' | 'horizontal-bar' | 'area' | 'gauge' | 'heatmap' | 'scatter' | 'gantt';

export interface ChartWebPartConfig extends WebPartConfig {
  type: 'chart';
  chartType?: ChartType;
  dataSource?: WebPartDataSource;
  filters?: WebPartFilter[];
  joins?: WebPartJoin[];

  // Chart-specific settings
  groupByColumn?: string; // Column to group by (X-axis / segments)
  valueColumn?: string; // Column to aggregate (Y-axis / values)
  aggregation?: ChartAggregation; // How to aggregate values

  // Display options
  legendPosition?: LegendPosition;
  legendLabel?: string; // Custom legend label (e.g., "Attendance Count")
  showLabels?: boolean;
  maxGroups?: number; // Limit number of groups (default: 10)
  showOther?: boolean; // Combine excess groups into "Other" (default: true)
  includeNull?: boolean; // Fill gaps with 0 values (e.g., missing dates in a range)
  sortBy?: 'label' | 'value';
  sortDirection?: 'asc' | 'desc';
  colorPalette?: string[]; // Custom colors
  xAxisLabelStyle?: XAxisLabelStyle; // Bar chart: 'normal' or 'angled' (shows all labels)

  // Gauge chart options
  gaugeMinValue?: number; // Minimum value (default: 0)
  gaugeMaxValue?: number; // Maximum value (auto-calculated if not set)

  // Heatmap chart options
  secondaryGroupByColumn?: string; // Y-axis grouping for heatmap

  // Gantt chart options
  ganttStartColumn?: string; // Column with start date
  ganttEndColumn?: string; // Column with end date
  ganttLabelColumn?: string; // Column for bar labels
}

/**
 * Union type for all WebPart configurations
 */
export type AnyWebPartConfig = ListItemsWebPartConfig | ChartWebPartConfig;

/**
 * A column within a section - contains a single WebPart (or empty)
 */
export interface ReportColumn {
  id: string;
  webPart: AnyWebPartConfig | null;
}

/**
 * Section height options
 */
export type SectionHeight = 'half' | 'medium' | 'full' | 'big';

/**
 * A horizontal section with a specific layout
 */
export interface ReportSection {
  id: string;
  layout: SectionLayout;
  columns: ReportColumn[];
  height?: SectionHeight; // Default is 'full' (100%)
}

/**
 * Complete report page layout configuration
 */
export interface ReportLayoutConfig {
  sections: ReportSection[];
}
