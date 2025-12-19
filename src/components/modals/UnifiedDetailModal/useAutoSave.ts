import { useRef, useCallback, useEffect } from 'react';

interface PendingSave {
  fieldName: string;
  value: unknown;
  timeoutId: ReturnType<typeof setTimeout>;
}

interface UseAutoSaveOptions {
  debounceMs?: number;
  onSave: (fieldName: string, value: unknown) => Promise<void>;
  onSaveStart?: (fieldName: string) => void;
  onSaveComplete?: (fieldName: string) => void;
  onSaveError?: (fieldName: string, error: Error) => void;
}

interface UseAutoSaveReturn {
  queueSave: (fieldName: string, value: unknown) => void;
  cancelSave: (fieldName: string) => void;
  flushAll: () => Promise<void>;
  isPending: (fieldName: string) => boolean;
}

const DEFAULT_DEBOUNCE_MS = 800;

export function useAutoSave({
  debounceMs = DEFAULT_DEBOUNCE_MS,
  onSave,
  onSaveStart,
  onSaveComplete,
  onSaveError,
}: UseAutoSaveOptions): UseAutoSaveReturn {
  const pendingSaves = useRef<Map<string, PendingSave>>(new Map());
  const savingFields = useRef<Set<string>>(new Set());

  // Execute save for a specific field
  const executeSave = useCallback(async (fieldName: string, value: unknown) => {
    // Remove from pending
    pendingSaves.current.delete(fieldName);

    // Mark as saving
    savingFields.current.add(fieldName);
    onSaveStart?.(fieldName);

    try {
      await onSave(fieldName, value);
      onSaveComplete?.(fieldName);
    } catch (error) {
      onSaveError?.(fieldName, error instanceof Error ? error : new Error(String(error)));
    } finally {
      savingFields.current.delete(fieldName);
    }
  }, [onSave, onSaveStart, onSaveComplete, onSaveError]);

  // Queue a save with debounce
  const queueSave = useCallback((fieldName: string, value: unknown) => {
    // Cancel existing pending save for this field
    const existing = pendingSaves.current.get(fieldName);
    if (existing) {
      clearTimeout(existing.timeoutId);
    }

    // Schedule new save
    const timeoutId = setTimeout(() => {
      executeSave(fieldName, value);
    }, debounceMs);

    pendingSaves.current.set(fieldName, {
      fieldName,
      value,
      timeoutId,
    });
  }, [debounceMs, executeSave]);

  // Cancel a pending save
  const cancelSave = useCallback((fieldName: string) => {
    const pending = pendingSaves.current.get(fieldName);
    if (pending) {
      clearTimeout(pending.timeoutId);
      pendingSaves.current.delete(fieldName);
    }
  }, []);

  // Flush all pending saves immediately
  const flushAll = useCallback(async () => {
    const savePromises: Promise<void>[] = [];

    pendingSaves.current.forEach((pending) => {
      clearTimeout(pending.timeoutId);
      savePromises.push(executeSave(pending.fieldName, pending.value));
    });

    pendingSaves.current.clear();
    await Promise.all(savePromises);
  }, [executeSave]);

  // Check if a field has a pending save
  const isPending = useCallback((fieldName: string) => {
    return pendingSaves.current.has(fieldName) || savingFields.current.has(fieldName);
  }, []);

  // Cleanup on unmount
  useEffect(() => {
    return () => {
      pendingSaves.current.forEach((pending) => {
        clearTimeout(pending.timeoutId);
      });
      pendingSaves.current.clear();
    };
  }, []);

  return {
    queueSave,
    cancelSave,
    flushAll,
    isPending,
  };
}
