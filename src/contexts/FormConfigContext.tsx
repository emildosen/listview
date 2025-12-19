import { createContext, useContext, useCallback, useRef, useMemo, type ReactNode } from 'react';
import { useMsal } from '@azure/msal-react';
import { getFormFieldConfig, getListItems, type FormFieldConfig } from '../auth/graphClient';

const CACHE_TTL_MS = 5 * 60 * 1000; // 5 minutes

export interface FormConfig {
  listId: string;
  siteId: string;
  contentTypeId: string;
  fields: FormFieldConfig[];
  fetchedAt: number;
}

export interface LookupOption {
  id: number;
  value: string;
}

interface LookupOptionsCache {
  options: LookupOption[];
  fetchedAt: number;
}

interface FormConfigContextValue {
  getFormConfig: (siteId: string, listId: string) => Promise<FormConfig>;
  getLookupOptions: (siteId: string, listId: string, columnName: string) => Promise<LookupOption[]>;
  invalidateCache: (listId: string) => void;
}

const FormConfigContext = createContext<FormConfigContextValue | null>(null);

export function FormConfigProvider({ children }: { children: ReactNode }) {
  const { instance, accounts } = useMsal();
  const cacheRef = useRef<Record<string, FormConfig>>({});
  const pendingRef = useRef<Record<string, Promise<FormConfig>>>({});
  const lookupCacheRef = useRef<Record<string, LookupOptionsCache>>({});
  const lookupPendingRef = useRef<Record<string, Promise<LookupOption[]>>>({});

  // Use refs to avoid re-creating getFormConfig when accounts changes
  const instanceRef = useRef(instance);
  const accountsRef = useRef(accounts);
  instanceRef.current = instance;
  accountsRef.current = accounts;

  const getFormConfig = useCallback(
    async (siteId: string, listId: string): Promise<FormConfig> => {
      const cacheKey = `${siteId}:${listId}`;
      const cached = cacheRef.current[cacheKey];
      const now = Date.now();

      // Return cached if still valid
      if (cached && now - cached.fetchedAt < CACHE_TTL_MS) {
        return cached;
      }

      // Deduplicate concurrent requests
      if (cacheKey in pendingRef.current) {
        return pendingRef.current[cacheKey];
      }

      const account = accountsRef.current[0];
      if (!account) {
        throw new Error('No authenticated account');
      }

      // Fetch and cache
      const fetchPromise = (async () => {
        try {
          const { contentTypeId, fields } = await getFormFieldConfig(
            instanceRef.current,
            account,
            siteId,
            listId
          );

          const config: FormConfig = {
            listId,
            siteId,
            contentTypeId,
            fields,
            fetchedAt: now,
          };

          cacheRef.current[cacheKey] = config;
          return config;
        } finally {
          delete pendingRef.current[cacheKey];
        }
      })();

      pendingRef.current[cacheKey] = fetchPromise;
      return fetchPromise;
    },
    [] // No dependencies - uses refs for mutable values
  );

  const getLookupOptions = useCallback(
    async (siteId: string, listId: string, columnName: string): Promise<LookupOption[]> => {
      const cacheKey = `${siteId}:${listId}:${columnName}`;
      const cached = lookupCacheRef.current[cacheKey];
      const now = Date.now();

      // Return cached if still valid
      if (cached && now - cached.fetchedAt < CACHE_TTL_MS) {
        return cached.options;
      }

      // Deduplicate concurrent requests
      if (cacheKey in lookupPendingRef.current) {
        return lookupPendingRef.current[cacheKey];
      }

      const account = accountsRef.current[0];
      if (!account) {
        throw new Error('No authenticated account');
      }

      // Fetch and cache
      const fetchPromise = (async () => {
        try {
          const result = await getListItems(
            instanceRef.current,
            account,
            siteId,
            listId
          );

          // Extract id and the display column value
          const options: LookupOption[] = result.items.map((item) => ({
            id: parseInt(item.id, 10),
            value: String(item.fields[columnName] ?? item.fields['Title'] ?? item.id),
          }));

          // Sort alphabetically by value
          options.sort((a, b) => a.value.localeCompare(b.value));

          lookupCacheRef.current[cacheKey] = { options, fetchedAt: now };
          return options;
        } finally {
          delete lookupPendingRef.current[cacheKey];
        }
      })();

      lookupPendingRef.current[cacheKey] = fetchPromise;
      return fetchPromise;
    },
    [] // No dependencies - uses refs for mutable values
  );

  const invalidateCache = useCallback((listId: string) => {
    // Remove all cache entries for this listId (any siteId)
    for (const key of Object.keys(cacheRef.current)) {
      if (key.endsWith(`:${listId}`)) {
        delete cacheRef.current[key];
      }
    }
    // Also invalidate lookup cache for this list
    for (const key of Object.keys(lookupCacheRef.current)) {
      if (key.includes(`:${listId}:`)) {
        delete lookupCacheRef.current[key];
      }
    }
  }, []);

  // Memoize context value to prevent unnecessary re-renders of consumers
  const contextValue = useMemo(
    () => ({ getFormConfig, getLookupOptions, invalidateCache }),
    [getFormConfig, getLookupOptions, invalidateCache]
  );

  return (
    <FormConfigContext.Provider value={contextValue}>
      {children}
    </FormConfigContext.Provider>
  );
}

export function useFormConfigContext() {
  const context = useContext(FormConfigContext);
  if (!context) {
    throw new Error('useFormConfigContext must be used within a FormConfigProvider');
  }
  return context;
}
