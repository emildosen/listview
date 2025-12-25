import {
  PersonRegular,
  PeopleRegular,
  DocumentRegular,
  FolderRegular,
  CalendarRegular,
  ClockRegular,
  TaskListAddRegular,
  CheckmarkRegular,
  DataPieRegular,
  ChartMultipleRegular,
  BuildingRegular,
  HomeRegular,
  MailRegular,
  ChatRegular,
  MoneyRegular,
  ReceiptRegular,
  BookRegular,
  HatGraduationRegular,
  HeartRegular,
  StarRegular,
  TagRegular,
  BookmarkRegular,
  GlobeRegular,
  LightbulbRegular,
  PhoneRegular,
  VideoRegular,
  RocketRegular,
  SettingsRegular,
  ShieldRegular,
  KeyRegular,
  type FluentIcon,
} from '@fluentui/react-icons';

// Curated list of icons for pages
export const PAGE_ICONS: Record<string, FluentIcon> = {
  PersonRegular,
  PeopleRegular,
  DocumentRegular,
  FolderRegular,
  CalendarRegular,
  ClockRegular,
  TaskListAddRegular,
  CheckmarkRegular,
  DataPieRegular,
  ChartMultipleRegular,
  BuildingRegular,
  HomeRegular,
  MailRegular,
  ChatRegular,
  MoneyRegular,
  ReceiptRegular,
  BookRegular,
  HatGraduationRegular,
  HeartRegular,
  StarRegular,
  TagRegular,
  BookmarkRegular,
  GlobeRegular,
  LightbulbRegular,
  PhoneRegular,
  VideoRegular,
  RocketRegular,
  SettingsRegular,
  ShieldRegular,
  KeyRegular,
};

// Get icon component by name, with fallback to DocumentRegular
export function getPageIcon(iconName: string | undefined): FluentIcon {
  if (!iconName) return DocumentRegular;
  return PAGE_ICONS[iconName] || DocumentRegular;
}

// List of available icon names for pickers
export const PAGE_ICON_OPTIONS = Object.keys(PAGE_ICONS);

// Get display name from icon name (e.g., "PersonRegular" -> "Person")
export function getIconDisplayName(iconName: string): string {
  return iconName.replace(/Regular$/, '');
}

// Default icons based on page type
export const DEFAULT_PAGE_ICONS: Record<string, string> = {
  lookup: 'DocumentRegular',
  report: 'DataPieRegular',
};
