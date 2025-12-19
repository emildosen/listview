import { useState, useCallback, useRef, useEffect } from 'react';

export interface FieldEditState {
  isEditing: boolean;
  isHovered: boolean;
  isSaving: boolean;
  error: string | null;
}

interface UseFieldEditOptions {
  initialValue: unknown;
  onSave: (value: unknown) => Promise<void>;
  onCancel?: () => void;
}

interface UseFieldEditReturn {
  state: FieldEditState;
  value: unknown;
  startEdit: () => void;
  cancelEdit: () => void;
  setValue: (value: unknown) => void;
  commitEdit: () => Promise<void>;
  setHovered: (hovered: boolean) => void;
  clearError: () => void;
  inputRef: React.RefObject<HTMLInputElement | HTMLTextAreaElement | HTMLSelectElement>;
}

export function useFieldEdit({
  initialValue,
  onSave,
  onCancel,
}: UseFieldEditOptions): UseFieldEditReturn {
  const [state, setState] = useState<FieldEditState>({
    isEditing: false,
    isHovered: false,
    isSaving: false,
    error: null,
  });

  const [value, setValueState] = useState<unknown>(initialValue);
  const originalValue = useRef<unknown>(initialValue);
  const inputRef = useRef<HTMLInputElement | HTMLTextAreaElement | HTMLSelectElement>(null);

  // Update value when initialValue changes (e.g., after successful save)
  // Valid prop-to-state sync pattern for controlled editing
  useEffect(() => {
    if (!state.isEditing) {
      // eslint-disable-next-line react-hooks/set-state-in-effect
      setValueState(initialValue);
      originalValue.current = initialValue;
    }
  }, [initialValue, state.isEditing]);

  const startEdit = useCallback(() => {
    originalValue.current = value;
    setState(prev => ({ ...prev, isEditing: true, error: null }));
    // Focus input after state update
    setTimeout(() => {
      inputRef.current?.focus();
      if (inputRef.current && 'select' in inputRef.current) {
        inputRef.current.select();
      }
    }, 0);
  }, [value]);

  const cancelEdit = useCallback(() => {
    setValueState(originalValue.current);
    setState(prev => ({ ...prev, isEditing: false, error: null }));
    onCancel?.();
  }, [onCancel]);

  const setValue = useCallback((newValue: unknown) => {
    setValueState(newValue);
  }, []);

  const commitEdit = useCallback(async () => {
    // Don't save if value hasn't changed
    if (value === originalValue.current) {
      setState(prev => ({ ...prev, isEditing: false }));
      return;
    }

    setState(prev => ({ ...prev, isSaving: true, error: null }));

    try {
      await onSave(value);
      originalValue.current = value;
      setState(prev => ({ ...prev, isEditing: false, isSaving: false }));
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Failed to save';
      setState(prev => ({ ...prev, isSaving: false, error: message }));
    }
  }, [value, onSave]);

  const setHovered = useCallback((hovered: boolean) => {
    setState(prev => ({ ...prev, isHovered: hovered }));
  }, []);

  const clearError = useCallback(() => {
    setState(prev => ({ ...prev, error: null }));
  }, []);

  return {
    state,
    value,
    startEdit,
    cancelEdit,
    setValue,
    commitEdit,
    setHovered,
    clearError,
    inputRef: inputRef as React.RefObject<HTMLInputElement | HTMLTextAreaElement | HTMLSelectElement>,
  };
}
