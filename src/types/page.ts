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

export interface PageDefinition {
  id?: string;                        // SharePoint item ID
  name: string;                       // Display name (e.g., "Student Details")
  description?: string;               // Optional description

  // Primary entity configuration
  primarySource: PageSource;          // Main list (e.g., Students)
  displayColumns: PageColumn[];       // Columns shown in detail view

  // Search/filter configuration
  searchConfig: SearchConfig;

  // Related lists configuration
  relatedSections: RelatedSection[];  // Related lists to show (e.g., Correspondence)

  // Detail modal layout configuration
  detailLayout?: DetailLayoutConfig;

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
