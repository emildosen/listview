import { useState, useEffect, useRef, useCallback } from 'react';
import { useFormConfigContext, type LookupOption } from '../contexts/FormConfigContext';
import type { FormFieldConfig } from '../auth/graphClient';

interface UseListFormConfigResult {
  fields: FormFieldConfig[];
  loading: boolean;
  error: string | null;
  getLookupOptions: (targetSiteId: string, targetListId: string, columnName: string) => Promise<LookupOption[]>;
}

/**
 * Hook to get form field configuration for a list.
 * Fetches from cache or Graph API, handles loading/error states.
 * Fields are returned in the order they appear in SharePoint's default form.
 */
export function useListFormConfig(
  siteId: string | undefined,
  listId: string | undefined
): UseListFormConfigResult {
  const { getFormConfig, getLookupOptions: contextGetLookupOptions } = useFormConfigContext();
  const [fields, setFields] = useState<FormFieldConfig[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  // Use ref to avoid re-running effect when getFormConfig changes reference
  const getFormConfigRef = useRef(getFormConfig);
  getFormConfigRef.current = getFormConfig;

  // Stable ref for getLookupOptions
  const getLookupOptionsRef = useRef(contextGetLookupOptions);
  getLookupOptionsRef.current = contextGetLookupOptions;

  // Stable wrapper for getLookupOptions
  const getLookupOptions = useCallback(
    (targetSiteId: string, targetListId: string, columnName: string) => {
      return getLookupOptionsRef.current(targetSiteId, targetListId, columnName);
    },
    []
  );

  // Track if we've already fetched to prevent re-fetching
  const fetchedRef = useRef<string | null>(null);

  useEffect(() => {
    if (!siteId || !listId) {
      setLoading(false);
      setFields([]);
      return;
    }

    // Skip if we already fetched for this siteId/listId combination
    const key = `${siteId}:${listId}`;
    if (fetchedRef.current === key) {
      return;
    }

    let cancelled = false;

    const fetchConfig = async () => {
      setLoading(true);
      setError(null);

      try {
        const config = await getFormConfigRef.current(siteId, listId);
        if (!cancelled) {
          setFields(config.fields);
          fetchedRef.current = key;
        }
      } catch (err) {
        if (!cancelled) {
          setError(err instanceof Error ? err.message : 'Failed to load form config');
        }
      } finally {
        if (!cancelled) {
          setLoading(false);
        }
      }
    };

    fetchConfig();

    return () => {
      cancelled = true;
    };
  }, [siteId, listId]); // Removed getFormConfig - it's stable and accessed via ref

  return { fields, loading, error, getLookupOptions };
}
