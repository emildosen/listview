// View type definitions for the Views feature

export interface ViewDefinition {
  id?: string;                    // SharePoint item ID
  name: string;                   // Display name
  description?: string;
  mode: 'union' | 'aggregate';    // Rollup mode
  sources: ViewSource[];          // Lists to pull from
  columns: ViewColumn[];          // Columns to display
  filters?: ViewFilter[];         // Optional filters
  sorting?: ViewSorting;          // Optional sorting
  createdAt?: string;
  updatedAt?: string;
}

export interface ViewSource {
  siteId: string;
  listId: string;
  listName: string;               // For display
}

export interface ViewColumn {
  sourceListId: string;           // Which list this column comes from
  internalName: string;           // SharePoint column internal name
  displayName: string;            // Custom display name
  aggregation?: AggregationType;  // For aggregate mode
}

export type AggregationType = 'count' | 'sum' | 'avg' | 'min' | 'max';

export type FilterOperator = 'eq' | 'ne' | 'gt' | 'lt' | 'contains';

export interface ViewFilter {
  column: string;
  operator: FilterOperator;
  value: string;
}

export interface ViewSorting {
  column: string;
  direction: 'asc' | 'desc';
}

// SharePoint list item representation
export interface ViewItem {
  Id: number;
  Title: string;
  ViewConfig: string;  // JSON stringified ViewDefinition
}
