import { useState, useEffect, useCallback, useMemo, useId } from 'react';
import { useMsal } from '@azure/msal-react';
import type { SPFI } from '@pnp/sp';
import {
  makeStyles,
  tokens,
  Text,
  Button,
  Spinner,
  MessageBar,
  MessageBarBody,
  Dialog,
  DialogSurface,
  DialogBody,
  DialogTitle,
  mergeClasses,
} from '@fluentui/react-components';
import {
  DismissRegular,
  SettingsRegular,
  DeleteRegular,
  ArrowLeftRegular,
  ArrowRightRegular,
  OpenRegular,
} from '@fluentui/react-icons';
import { getListItems, isSharePointUrl, type GraphListColumn, type GraphListItem } from '../../../auth/graphClient';
import { updateListItem, deleteListItem, createSPClient } from '../../../services/sharepoint';
import type { PageDefinition, PageColumn, DetailLayoutConfig, ListDetailConfig, RelatedSection } from '../../../types/page';
import { useSettings } from '../../../contexts/SettingsContext';
import { useTheme } from '../../../contexts/ThemeContext';
import { ModalNavigationProvider, useModalNavigation, type NavigationEntry } from './ModalNavigationContext';
import { InlineEditField } from './InlineEditField';
import { InlineEditText } from './InlineEditText';
import { InlineEditChoice } from './InlineEditChoice';
import { InlineEditLookup } from './InlineEditLookup';
import { InlineEditNumber } from './InlineEditNumber';
import { InlineEditDate, formatDateForInput, formatDateTimeForInput, formatDateForDisplay, formatDateTimeForDisplay } from './InlineEditDate';
import { InlineEditBoolean } from './InlineEditBoolean';
import { DescriptionField } from './DescriptionField';
import { RelatedSectionView } from './RelatedSectionView';
import StatBox from '../../PageDisplay/StatBox';
import DetailCustomizeDrawer from '../../PageDisplay/DetailCustomizeDrawer';
import { SharePointLink } from '../../common/SharePointLink';
import { useListFormConfig } from '../../../hooks/useListFormConfig';
import type { LookupOption } from '../../../contexts/FormConfigContext';

const useStyles = makeStyles({
  surface: {
    maxWidth: '1000px',
    width: '95vw',
    maxHeight: '90vh',
  },
  dialogTitle: {
    display: 'flex',
    flexDirection: 'row',
    alignItems: 'center',
    justifyContent: 'space-between',
    gap: '16px',
  },
  navButtons: {
    display: 'flex',
    alignItems: 'center',
    gap: '2px',
    marginRight: '8px',
  },
  titleText: {
    fontSize: tokens.fontSizeBase500,
    fontWeight: tokens.fontWeightSemibold,
    lineHeight: tokens.lineHeightBase500,
    flex: 1,
    minWidth: 0,
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap',
  },
  headerActions: {
    display: 'flex',
    alignItems: 'center',
    gap: '4px',
    flexShrink: 0,
  },
  body: {
    display: 'block',
    overflowY: 'auto',
    maxHeight: 'calc(90vh - 80px)',
    '& > *': {
      marginBottom: '24px',
    },
    '& > *:last-child': {
      marginBottom: 0,
    },
  },
  statBoxContainer: {
    display: 'flex',
    flexWrap: 'wrap',
    gap: '8px',
    marginTop: '10px',
  },
  detailsCard: {
    padding: '16px',
    backgroundColor: tokens.colorNeutralBackground2,
    borderRadius: tokens.borderRadiusMedium,
  },
  detailsCardDark: {
    backgroundColor: '#1a1a1a',
    border: '1px solid #333333',
  },
  sectionTitle: {
    fontWeight: tokens.fontWeightSemibold,
    fontSize: tokens.fontSizeBase200,
    textTransform: 'uppercase',
    letterSpacing: '0.02em',
    color: tokens.colorNeutralForeground2,
    marginBottom: '12px',
  },
  detailsGrid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(2, 1fr)',
    gap: '8px',
  },
  loadingContainer: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    padding: '48px',
  },
  relatedSection: {
    marginTop: '24px',
  },
  relatedSectionHeader: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    marginBottom: '12px',
  },
});

interface UnifiedDetailModalProps {
  listId: string;
  listName: string;
  siteId: string;
  siteUrl?: string;
  columns: GraphListColumn[];
  item: GraphListItem;
  page?: PageDefinition;
  titleColumnOverride?: string;
  onClose: () => void;
  onItemUpdated?: () => void;
  onItemDeleted?: () => void;
}

export function UnifiedDetailModal(props: UnifiedDetailModalProps) {
  const modalId = useId();
  const initialEntry: NavigationEntry = {
    listId: props.listId,
    siteId: props.siteId,
    siteUrl: props.siteUrl,
    itemId: props.item.id,
    listName: props.listName,
  };

  return (
    <ModalNavigationProvider modalId={modalId} initialEntry={initialEntry}>
      <UnifiedDetailModalContent {...props} />
    </ModalNavigationProvider>
  );
}

function UnifiedDetailModalContent({
  listId: initialListId,
  listName: initialListName,
  siteId: initialSiteId,
  siteUrl: initialSiteUrl,
  columns: initialColumns,
  item: initialItem,
  page,
  titleColumnOverride,
  onClose,
  onItemUpdated,
  onItemDeleted,
}: UnifiedDetailModalProps) {
  const styles = useStyles();
  const { theme } = useTheme();
  const { instance, accounts } = useMsal();
  const account = accounts[0];
  const { getListDetailConfig, saveListDetailConfig } = useSettings();
  const { currentEntry, canGoBack, canGoForward, goBack, goForward, isNavigating, setIsNavigating } = useModalNavigation();

  // Current display state - may differ from initial props when navigating
  const [currentListId, setCurrentListId] = useState(initialListId);
  const [currentListName, setCurrentListName] = useState(initialListName);
  const [currentSiteId, setCurrentSiteId] = useState(initialSiteId);
  const [currentSiteUrl, setCurrentSiteUrl] = useState(initialSiteUrl);
  const [currentColumns, setCurrentColumns] = useState(initialColumns);
  const [currentItem, setCurrentItem] = useState(initialItem);
  const [listDetailConfig, setListDetailConfig] = useState<ListDetailConfig | null>(null);

  // UI state
  const [customizeOpen, setCustomizeOpen] = useState(false);
  const [deleting, setDeleting] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [spClient, setSpClient] = useState<SPFI | null>(null);

  // Field edit state
  const [editingField, setEditingField] = useState<string | null>(null);
  const [hoveredField, setHoveredField] = useState<string | null>(null);
  const [savingFields, setSavingFields] = useState<Set<string>>(new Set());
  const [fieldErrors, setFieldErrors] = useState<Record<string, string>>({});

  // Form config for field metadata and lookup options
  const { fields: formFields, getLookupOptions } = useListFormConfig(currentSiteId, currentListId);
  const [lookupOptions, setLookupOptions] = useState<Record<string, LookupOption[]>>({});

  // Initialize SP client
  useEffect(() => {
    if (!currentSiteUrl || !account) return;
    createSPClient(instance, account, currentSiteUrl)
      .then(setSpClient)
      .catch(console.error);
  }, [instance, account, currentSiteUrl]);

  // Load list detail config
  useEffect(() => {
    const config = getListDetailConfig(currentListId);
    if (config) {
      setListDetailConfig(config);
    } else {
      // Create default config from columns
      const defaultConfig = createDefaultListDetailConfig(
        currentListId,
        currentListName,
        currentSiteId,
        currentSiteUrl,
        currentColumns
      );
      setListDetailConfig(defaultConfig);
    }
  }, [currentListId, currentListName, currentSiteId, currentSiteUrl, currentColumns, getListDetailConfig]);

  // Handle navigation changes
  useEffect(() => {
    if (!currentEntry) return;

    // Check if we navigated to a different item
    if (currentEntry.listId !== currentListId || currentEntry.itemId !== currentItem.id) {
      loadNavigatedItem(currentEntry);
    }
  }, [currentEntry]);

  const loadNavigatedItem = async (entry: NavigationEntry) => {
    setIsNavigating(true);
    setError(null);

    try {
      // Fetch item data
      const { columns, items } = await getListItems(instance, account, entry.siteId, entry.listId);
      const item = items.find(i => i.id === entry.itemId);

      if (!item) {
        throw new Error('Item not found');
      }

      // Update state
      setCurrentListId(entry.listId);
      setCurrentListName(entry.listName);
      setCurrentSiteId(entry.siteId);
      setCurrentSiteUrl(entry.siteUrl);
      setCurrentColumns(columns);
      setCurrentItem(item);

      // Load config for new list
      const config = getListDetailConfig(entry.listId);
      if (config) {
        setListDetailConfig(config);
      } else {
        const defaultConfig = createDefaultListDetailConfig(
          entry.listId,
          entry.listName,
          entry.siteId,
          entry.siteUrl,
          columns
        );
        setListDetailConfig(defaultConfig);
      }

      // Create new SP client if site changed
      if (entry.siteUrl && entry.siteUrl !== currentSiteUrl) {
        const client = await createSPClient(instance, account, entry.siteUrl);
        setSpClient(client);
      }
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Failed to load item');
    } finally {
      setIsNavigating(false);
    }
  };

  // Get layout config
  const layoutConfig = useMemo(() => {
    if (!listDetailConfig) return null;
    return getEffectiveLayoutConfig(
      listDetailConfig.detailLayout,
      listDetailConfig.displayColumns,
      listDetailConfig.relatedSections
    );
  }, [listDetailConfig]);

  // Determine title column
  const titleColumn = useMemo(() => {
    return (
      titleColumnOverride ??
      page?.searchConfig?.tableColumns?.[0]?.internalName ??
      listDetailConfig?.displayColumns[0]?.internalName ??
      'Title'
    );
  }, [titleColumnOverride, page, listDetailConfig]);

  // Separate columns by display style
  const { statColumns, listColumns, descriptionColumn } = useMemo(() => {
    if (!layoutConfig) {
      return { statColumns: [], listColumns: [], descriptionColumn: null };
    }

    const visible = layoutConfig.columnSettings.filter(c => c.visible && c.internalName !== titleColumn);
    return {
      statColumns: visible.filter(c => c.displayStyle === 'stat'),
      listColumns: visible.filter(c => c.displayStyle === 'list'),
      descriptionColumn: visible.find(c => c.displayStyle === 'description') ?? null,
    };
  }, [layoutConfig, titleColumn]);

  // Get column metadata
  const getColumnMetadata = useCallback((internalName: string) => {
    return currentColumns.find(c => c.name === internalName);
  }, [currentColumns]);

  const getFormField = useCallback((internalName: string) => {
    return formFields.find(f => f.name === internalName);
  }, [formFields]);

  const getDisplayName = useCallback((internalName: string) => {
    const pageCol = listDetailConfig?.displayColumns.find(c => c.internalName === internalName);
    if (pageCol) return pageCol.displayName;
    const col = getColumnMetadata(internalName);
    return col?.displayName ?? internalName;
  }, [listDetailConfig, getColumnMetadata]);

  // Auto-save handler
  const handleSaveField = useCallback(async (fieldName: string, value: unknown) => {
    if (!spClient) return;

    setSavingFields(prev => new Set(prev).add(fieldName));
    setFieldErrors(prev => {
      const next = { ...prev };
      delete next[fieldName];
      return next;
    });

    try {
      const formField = getFormField(fieldName);
      const payload: Record<string, unknown> = {};

      // Handle lookup fields specially
      if (formField?.lookup) {
        payload[`${fieldName}Id`] = value;
      } else {
        payload[fieldName] = value;
      }

      await updateListItem(spClient, currentListId, parseInt(currentItem.id, 10), payload);

      // Update local item state
      setCurrentItem(prev => ({
        ...prev,
        fields: { ...prev.fields, [fieldName]: value },
      }));

      onItemUpdated?.();
    } catch (err) {
      const message = err instanceof Error ? err.message : 'Failed to save';
      setFieldErrors(prev => ({ ...prev, [fieldName]: message }));
      throw err;
    } finally {
      setSavingFields(prev => {
        const next = new Set(prev);
        next.delete(fieldName);
        return next;
      });
    }
  }, [spClient, currentListId, currentItem.id, getFormField, onItemUpdated]);

  // Handle delete
  const handleDelete = async () => {
    if (!spClient) return;

    setDeleting(true);
    try {
      await deleteListItem(spClient, currentListId, parseInt(currentItem.id, 10));
      onItemDeleted?.();
      onClose();
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Failed to delete');
    } finally {
      setDeleting(false);
    }
  };

  // Handle customize save
  const handleSaveConfig = async (config: DetailLayoutConfig, relatedSections?: RelatedSection[]) => {
    if (!listDetailConfig) return;

    const updatedConfig: ListDetailConfig = {
      ...listDetailConfig,
      detailLayout: config,
      relatedSections: relatedSections ?? listDetailConfig.relatedSections,
    };

    await saveListDetailConfig(updatedConfig);
    setListDetailConfig(updatedConfig);
    setCustomizeOpen(false);
  };

  // Get stat value as string (for StatBox)
  const getStatValue = (fieldName: string, value: unknown): string => {
    if (value === null || value === undefined) return '-';

    const formField = getFormField(fieldName);
    const colMeta = getColumnMetadata(fieldName);

    // Boolean
    if (formField?.boolean || colMeta?.boolean) {
      return value ? 'Yes' : 'No';
    }

    // Date
    if (formField?.dateTime || colMeta?.dateTime) {
      const isDateOnly = (formField?.dateTime?.format ?? colMeta?.dateTime?.format) === 'dateOnly';
      return isDateOnly ? formatDateForDisplay(value) : formatDateTimeForDisplay(value);
    }

    // Lookup
    if (formField?.lookup || colMeta?.lookup) {
      if (typeof value === 'object' && value !== null && 'LookupValue' in value) {
        return (value as { LookupValue: string }).LookupValue;
      }
      if (Array.isArray(value)) {
        return value
          .map(v => (typeof v === 'object' && v !== null && 'LookupValue' in v ? v.LookupValue : String(v)))
          .join(', ');
      }
    }

    return String(value);
  };

  // Render field value (view mode) - can return ReactNode for links
  const renderFieldValue = (fieldName: string, value: unknown): React.ReactNode => {
    if (value === null || value === undefined) return '-';

    // URL - render as link
    if (typeof value === 'string' && isSharePointUrl(value)) {
      return <SharePointLink url={value} />;
    }

    return getStatValue(fieldName, value);
  };

  // Get title value
  const titleValue = currentItem.fields[titleColumn];
  const titleDisplay = typeof titleValue === 'object' && titleValue !== null && 'LookupValue' in titleValue
    ? (titleValue as { LookupValue: string }).LookupValue
    : String(titleValue ?? currentListName);

  if (!listDetailConfig || !layoutConfig) {
    return (
      <Dialog open onOpenChange={(_, data) => !data.open && onClose()}>
        <DialogSurface className={styles.surface}>
          <div className={styles.loadingContainer}>
            <Spinner size="medium" />
          </div>
        </DialogSurface>
      </Dialog>
    );
  }

  return (
    <>
      <Dialog open onOpenChange={(_, data) => !data.open && onClose()}>
        <DialogSurface className={styles.surface}>
          <DialogTitle className={styles.dialogTitle}>
            {/* Navigation buttons */}
            <div className={styles.navButtons}>
              <Button
                appearance="subtle"
                size="small"
                icon={<ArrowLeftRegular />}
                disabled={!canGoBack || isNavigating}
                onClick={goBack}
                title="Back"
              />
              <Button
                appearance="subtle"
                size="small"
                icon={<ArrowRightRegular />}
                disabled={!canGoForward || isNavigating}
                onClick={goForward}
                title="Forward"
              />
            </div>

            <Text className={styles.titleText}>{titleDisplay}</Text>

            <div className={styles.headerActions}>
              {currentSiteUrl && (
                <Button
                  appearance="subtle"
                  size="small"
                  icon={<OpenRegular />}
                  as="a"
                  href={`${currentSiteUrl}/Lists/${encodeURIComponent(currentListName)}/DispForm.aspx?ID=${currentItem.id}`}
                  target="_blank"
                  title="Open in SharePoint"
                />
              )}
              <Button
                appearance="subtle"
                size="small"
                icon={<DeleteRegular />}
                onClick={handleDelete}
                disabled={deleting}
                title="Delete"
              />
              <Button
                appearance="subtle"
                size="small"
                icon={<SettingsRegular />}
                onClick={() => setCustomizeOpen(true)}
                title="Customize"
              />
              <Button
                appearance="subtle"
                size="small"
                icon={<DismissRegular />}
                onClick={onClose}
                title="Close"
              />
            </div>
          </DialogTitle>

          <DialogBody className={styles.body}>
            {error && (
              <MessageBar intent="error">
                <MessageBarBody>{error}</MessageBarBody>
              </MessageBar>
            )}

            {isNavigating ? (
              <div className={styles.loadingContainer}>
                <Spinner size="medium" label="Loading..." />
              </div>
            ) : (
              <>
                {/* Stat boxes */}
                {statColumns.length > 0 && (
                  <div className={styles.statBoxContainer}>
                    {statColumns.map(col => {
                      const value = currentItem.fields[col.internalName];
                      return (
                        <StatBox
                          key={col.internalName}
                          label={getDisplayName(col.internalName)}
                          value={getStatValue(col.internalName, value)}
                        />
                      );
                    })}
                  </div>
                )}

                {/* Description field */}
                {descriptionColumn && (
                  <DescriptionField
                    value={String(currentItem.fields[descriptionColumn.internalName] ?? '')}
                    isRichText={getColumnMetadata(descriptionColumn.internalName)?.text?.textType === 'richText'}
                    isSaving={savingFields.has(descriptionColumn.internalName)}
                    readOnly={getColumnMetadata(descriptionColumn.internalName)?.readOnly}
                    onSave={(value) => handleSaveField(descriptionColumn.internalName, value)}
                  />
                )}

                {/* Detail fields */}
                {listColumns.length > 0 && (
                  <div className={mergeClasses(styles.detailsCard, theme === 'dark' && styles.detailsCardDark)}>
                    <Text className={styles.sectionTitle}>Details</Text>
                    <div className={styles.detailsGrid}>
                      {listColumns.map(col => (
                        <DetailFieldEdit
                          key={col.internalName}
                          fieldName={col.internalName}
                          label={getDisplayName(col.internalName)}
                          value={currentItem.fields[col.internalName]}
                          formField={getFormField(col.internalName)}
                          columnMetadata={getColumnMetadata(col.internalName)}
                          isEditing={editingField === col.internalName}
                          isHovered={hoveredField === col.internalName}
                          isSaving={savingFields.has(col.internalName)}
                          error={fieldErrors[col.internalName] ?? null}
                          siteId={currentSiteId}
                          getLookupOptions={getLookupOptions}
                          lookupOptions={lookupOptions}
                          setLookupOptions={setLookupOptions}
                          onStartEdit={() => setEditingField(col.internalName)}
                          onCancelEdit={() => setEditingField(null)}
                          onSave={handleSaveField}
                          onMouseEnter={() => setHoveredField(col.internalName)}
                          onMouseLeave={() => setHoveredField(null)}
                          onClearError={() => setFieldErrors(prev => {
                            const next = { ...prev };
                            delete next[col.internalName];
                            return next;
                          })}
                          renderValue={renderFieldValue}
                        />
                      ))}
                    </div>
                  </div>
                )}

                {/* Related sections */}
                {listDetailConfig.relatedSections.length > 0 && (
                  <>
                    {(layoutConfig.relatedSectionOrder ?? listDetailConfig.relatedSections.map(s => s.id)).map(sectionId => {
                      const section = listDetailConfig.relatedSections.find(s => s.id === sectionId);
                      if (!section) return null;
                      return (
                        <RelatedSectionView
                          key={section.id}
                          section={section}
                          parentItem={currentItem}
                        />
                      );
                    })}
                  </>
                )}
              </>
            )}
          </DialogBody>
        </DialogSurface>
      </Dialog>

      {/* Customize drawer */}
      <DetailCustomizeDrawer
        listDetailConfig={listDetailConfig}
        columnMetadata={currentColumns}
        titleColumn={titleColumn}
        open={customizeOpen}
        onClose={() => setCustomizeOpen(false)}
        onSave={handleSaveConfig}
      />
    </>
  );
}

// Helper component for inline editing a single field
interface DetailFieldEditProps {
  fieldName: string;
  label: string;
  value: unknown;
  formField: ReturnType<typeof import('../../../hooks/useListFormConfig').useListFormConfig>['fields'][0] | undefined;
  columnMetadata: GraphListColumn | undefined;
  isEditing: boolean;
  isHovered: boolean;
  isSaving: boolean;
  error: string | null;
  siteId: string;
  getLookupOptions: (siteId: string, listId: string, columnName: string) => Promise<LookupOption[]>;
  lookupOptions: Record<string, LookupOption[]>;
  setLookupOptions: React.Dispatch<React.SetStateAction<Record<string, LookupOption[]>>>;
  onStartEdit: () => void;
  onCancelEdit: () => void;
  onSave: (fieldName: string, value: unknown) => Promise<void>;
  onMouseEnter: () => void;
  onMouseLeave: () => void;
  onClearError: () => void;
  renderValue: (fieldName: string, value: unknown) => React.ReactNode;
}

function DetailFieldEdit({
  fieldName,
  label,
  value,
  formField,
  columnMetadata,
  isEditing,
  isHovered,
  isSaving,
  error,
  siteId,
  getLookupOptions,
  lookupOptions,
  setLookupOptions,
  onStartEdit,
  onCancelEdit,
  onSave,
  onMouseEnter,
  onMouseLeave,
  onClearError,
  renderValue,
}: DetailFieldEditProps) {
  const [editValue, setEditValue] = useState<unknown>(value);
  const [lookupLoading, setLookupLoading] = useState(false);

  // Sync edit value with prop
  useEffect(() => {
    if (!isEditing) {
      setEditValue(value);
    }
  }, [value, isEditing]);

  // Load lookup options when entering edit mode
  useEffect(() => {
    if (!isEditing || !formField?.lookup) return;
    if (lookupOptions[fieldName]) return;

    const loadOptions = async () => {
      setLookupLoading(true);
      try {
        const options = await getLookupOptions(siteId, formField.lookup!.listId, formField.lookup!.columnName);
        setLookupOptions(prev => ({ ...prev, [fieldName]: options }));
      } catch {
        setLookupOptions(prev => ({ ...prev, [fieldName]: [] }));
      } finally {
        setLookupLoading(false);
      }
    };
    loadOptions();
  }, [isEditing, formField, fieldName, siteId, getLookupOptions, lookupOptions, setLookupOptions]);

  const handleCommit = async () => {
    try {
      await onSave(fieldName, editValue);
      onCancelEdit();
    } catch {
      // Error is handled in parent
    }
  };

  const isReadOnly = formField?.readOnly || columnMetadata?.readOnly;

  // Render appropriate edit component based on field type
  const renderEditComponent = () => {
    // Choice field
    if (formField?.choice?.choices) {
      return (
        <InlineEditChoice
          value={String(editValue ?? '')}
          choices={formField.choice.choices}
          onChange={(v) => setEditValue(v)}
          onCommit={handleCommit}
          onCancel={onCancelEdit}
        />
      );
    }

    // Lookup field
    if (formField?.lookup) {
      const extractId = (v: unknown): number | null => {
        if (typeof v === 'number') return v;
        if (typeof v === 'object' && v !== null && 'LookupId' in v) {
          return (v as { LookupId: number }).LookupId;
        }
        return null;
      };

      const currentId = formField.lookup.allowMultipleValues
        ? (Array.isArray(editValue) ? editValue.map(extractId).filter((id): id is number => id !== null) : [])
        : extractId(editValue);

      return (
        <InlineEditLookup
          value={currentId}
          options={lookupOptions[fieldName] ?? []}
          isLoading={lookupLoading}
          isMultiSelect={formField.lookup.allowMultipleValues ?? false}
          onChange={(v) => setEditValue(v)}
          onCommit={handleCommit}
          onCancel={onCancelEdit}
        />
      );
    }

    // Boolean field
    if (formField?.boolean || columnMetadata?.boolean) {
      return (
        <InlineEditBoolean
          value={Boolean(editValue)}
          onChange={(v) => setEditValue(v)}
          onCommit={handleCommit}
          onCancel={onCancelEdit}
        />
      );
    }

    // Number field
    if (formField?.number || columnMetadata?.number) {
      return (
        <InlineEditNumber
          value={typeof editValue === 'number' ? editValue : null}
          onChange={(v) => setEditValue(v)}
          onCommit={handleCommit}
          onCancel={onCancelEdit}
        />
      );
    }

    // DateTime field
    if (formField?.dateTime || columnMetadata?.dateTime) {
      const isDateOnly = (formField?.dateTime?.format ?? columnMetadata?.dateTime?.format) === 'dateOnly';
      const formattedValue = isDateOnly
        ? formatDateForInput(editValue)
        : formatDateTimeForInput(editValue);

      return (
        <InlineEditDate
          value={formattedValue}
          dateOnly={isDateOnly}
          onChange={(v) => setEditValue(v)}
          onCommit={handleCommit}
          onCancel={onCancelEdit}
        />
      );
    }

    // Multiline text
    if (formField?.text?.allowMultipleLines || columnMetadata?.text?.allowMultipleLines) {
      return (
        <InlineEditText
          value={String(editValue ?? '')}
          multiline
          onChange={(v) => setEditValue(v)}
          onCommit={handleCommit}
          onCancel={onCancelEdit}
        />
      );
    }

    // Default: single line text
    return (
      <InlineEditText
        value={String(editValue ?? '')}
        onChange={(v) => setEditValue(v)}
        onCommit={handleCommit}
        onCancel={onCancelEdit}
      />
    );
  };

  return (
    <InlineEditField
      label={label}
      isEditing={isEditing}
      isHovered={isHovered}
      isSaving={isSaving}
      error={error}
      readOnly={isReadOnly}
      onStartEdit={onStartEdit}
      onMouseEnter={onMouseEnter}
      onMouseLeave={onMouseLeave}
      onClearError={onClearError}
      editComponent={renderEditComponent()}
    >
      {renderValue(fieldName, value)}
    </InlineEditField>
  );
}

// Helper functions

function createDefaultListDetailConfig(
  listId: string,
  listName: string,
  siteId: string,
  siteUrl: string | undefined,
  columns: GraphListColumn[]
): ListDetailConfig {
  const displayColumns: PageColumn[] = columns
    .filter(c => !c.hidden && !c.name.startsWith('_'))
    .map(c => ({
      internalName: c.name,
      displayName: c.displayName,
      editable: !c.readOnly,
    }));

  return {
    listId,
    listName,
    siteId,
    siteUrl,
    displayColumns,
    detailLayout: {
      columnSettings: displayColumns.map(c => ({
        internalName: c.internalName,
        visible: true,
        displayStyle: 'list' as const,
      })),
    },
    relatedSections: [],
  };
}

function getEffectiveLayoutConfig(
  existingLayout: DetailLayoutConfig | undefined,
  displayColumns: PageColumn[],
  relatedSections: RelatedSection[]
): DetailLayoutConfig {
  const defaultSettings = displayColumns.map(col => ({
    internalName: col.internalName,
    visible: true,
    displayStyle: 'list' as const,
  }));

  if (!existingLayout) {
    return {
      columnSettings: defaultSettings,
      relatedSectionOrder: relatedSections.map(s => s.id),
    };
  }

  const validColumnNames = new Set(displayColumns.map(c => c.internalName));
  const existingNames = new Set(existingLayout.columnSettings.map(s => s.internalName));

  const existingSettings = existingLayout.columnSettings.filter(s =>
    validColumnNames.has(s.internalName)
  );

  const newColumns = displayColumns
    .filter(col => !existingNames.has(col.internalName))
    .map(col => ({
      internalName: col.internalName,
      visible: true,
      displayStyle: 'list' as const,
    }));

  return {
    columnSettings: [...existingSettings, ...newColumns],
    relatedSectionOrder: existingLayout.relatedSectionOrder ?? relatedSections.map(s => s.id),
  };
}
