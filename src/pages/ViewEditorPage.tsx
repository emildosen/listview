import { useMemo } from 'react';
import { useParams, useNavigate, Link } from 'react-router-dom';
import { useSettings } from '../contexts/SettingsContext';
import ViewEditor from '../components/ViewEditor';
import type { ViewDefinition } from '../types/view';

function ViewEditorPage() {
  const { viewId } = useParams<{ viewId?: string }>();
  const navigate = useNavigate();
  const { views, saveView } = useSettings();

  const isEditMode = !!viewId;

  // Find view for editing
  const initialView = useMemo((): ViewDefinition | undefined => {
    if (!viewId) return undefined;
    return views.find((v) => v.id === viewId);
  }, [viewId, views]);

  const loading = isEditMode && views.length === 0;

  const handleSave = async (view: ViewDefinition) => {
    await saveView(view);
    navigate('/app/views');
  };

  const handleCancel = () => {
    navigate('/app/views');
  };

  if (loading) {
    return (
      <div className="p-8">
        <div className="flex items-center justify-center py-12">
          <span className="loading loading-spinner loading-lg text-primary" />
        </div>
      </div>
    );
  }

  if (isEditMode && !initialView) {
    return (
      <div className="p-8">
        <div className="text-sm breadcrumbs mb-6">
          <ul>
            <li>
              <Link to="/app">Home</Link>
            </li>
            <li>
              <Link to="/app/views">Views</Link>
            </li>
            <li>Not Found</li>
          </ul>
        </div>
        <div className="alert alert-error">
          <span>View not found</span>
        </div>
        <div className="mt-4">
          <Link to="/app/views" className="btn btn-ghost">
            Back to Views
          </Link>
        </div>
      </div>
    );
  }

  return (
    <div className="p-8">
      {/* Breadcrumb */}
      <div className="text-sm breadcrumbs mb-6">
        <ul>
          <li>
            <Link to="/app">Home</Link>
          </li>
          <li>
            <Link to="/app/views">Views</Link>
          </li>
          <li>{isEditMode ? `Edit: ${initialView?.name}` : 'Create View'}</li>
        </ul>
      </div>

      <div className="max-w-4xl">
        <div className="mb-6">
          <h1 className="text-2xl font-bold mb-1">
            {isEditMode ? 'Edit View' : 'Create View'}
          </h1>
          <p className="text-base-content/60">
            {isEditMode
              ? 'Modify the view configuration below.'
              : 'Configure a new view to combine and display data from multiple lists.'}
          </p>
        </div>

        <ViewEditor
          initialView={initialView}
          onSave={handleSave}
          onCancel={handleCancel}
        />
      </div>
    </div>
  );
}

export default ViewEditorPage;
