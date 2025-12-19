// Unified Detail Modal - Notion-style modal with inline editing and navigation
export { UnifiedDetailModal } from './UnifiedDetailModal';
export { ModalNavigationProvider, useModalNavigation, type NavigationEntry } from './ModalNavigationContext';
export { useAutoSave } from './useAutoSave';
export { useFieldEdit, type FieldEditState } from './useFieldEdit';

// Inline edit components
export { InlineEditField } from './InlineEditField';
export { InlineEditText } from './InlineEditText';
export { InlineEditChoice } from './InlineEditChoice';
export { InlineEditLookup } from './InlineEditLookup';
export { InlineEditNumber } from './InlineEditNumber';
export { InlineEditDate, formatDateForInput, formatDateTimeForInput, formatDateForDisplay, formatDateTimeForDisplay } from './InlineEditDate';
export { InlineEditBoolean } from './InlineEditBoolean';
export { DescriptionField } from './DescriptionField';
export { RelatedSectionView } from './RelatedSectionView';
