// Page type definitions for the Custom Pages feature

// Detail modal layout configuration (for customizable popup)
export interface DetailLayoutConfig {
  columnSettings: DetailColumnSetting[];
  relatedSectionOrder?: string[];  // Array of RelatedSection.id values for ordering
}

export interface DetailColumnSetting {
  internalName: string;
  visible: boolean;
  displayStyle: 'stat' | 'list';  // 'stat' = badge at top, 'list' = detail grid
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

export type DisplayMode = 'inline' | 'table';

export interface SearchConfig {
  displayMode: DisplayMode;           // 'inline' = side panel, 'table' = full table view
  titleColumn: string;                // Main column to display in search results (inline mode)
  subtitleColumns?: string[];         // Secondary columns shown below title (inline mode)
  tableColumns?: PageColumn[];        // Columns shown in table view (table mode)
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
  allowCreate: boolean;               // Can add new items
  allowEdit: boolean;                 // Can edit existing items
  allowDelete: boolean;               // Can delete items
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
 * Join configuration to link with another list
 */
export interface WebPartJoin {
  id: string;
  targetSource: WebPartDataSource; // The list to join with
  sourceColumn: string; // Column in primary list (lookup or ID)
  targetColumn: string; // Column in target list to match
  joinType: 'inner' | 'left'; // Inner = only matching, Left = all primary + matching
  columnsToInclude: string[]; // Which columns to pull from joined list
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
export interface ChartWebPartConfig extends WebPartConfig {
  type: 'chart';
  chartType?: 'bar' | 'donut' | 'line' | 'horizontal-bar';
  dataSource?: WebPartDataSource;
  filters?: WebPartFilter[];
  joins?: WebPartJoin[];

  // Chart-specific settings
  groupByColumn?: string; // Column to group by (X-axis / segments)
  valueColumn?: string; // Column to aggregate (Y-axis / values)
  aggregation?: ChartAggregation; // How to aggregate values

  // Display options
  showLegend?: boolean;
  showLabels?: boolean;
  maxGroups?: number; // Limit number of groups (default: 10)
  sortBy?: 'label' | 'value';
  sortDirection?: 'asc' | 'desc';
  colorPalette?: string[]; // Custom colors
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
 * A horizontal section with a specific layout
 */
export interface ReportSection {
  id: string;
  layout: SectionLayout;
  columns: ReportColumn[];
}

/**
 * Complete report page layout configuration
 */
export interface ReportLayoutConfig {
  sections: ReportSection[];
}
