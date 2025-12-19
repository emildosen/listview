import { useRef, useCallback } from 'react';
import { Editor } from '@tinymce/tinymce-react';
import { makeStyles, tokens, mergeClasses } from '@fluentui/react-components';
import { useTheme } from '../../contexts/ThemeContext';

// Import TinyMCE core and required modules for self-hosted bundling
import 'tinymce/tinymce';
import 'tinymce/themes/silver';
import 'tinymce/icons/default';
import 'tinymce/models/dom';

// Import plugins
import 'tinymce/plugins/lists';
import 'tinymce/plugins/link';
import 'tinymce/plugins/autolink';
import 'tinymce/plugins/table';
import 'tinymce/plugins/autoresize';
import 'tinymce/plugins/image';
import 'tinymce/plugins/media';
import 'tinymce/plugins/code';
import 'tinymce/plugins/charmap';
import 'tinymce/plugins/emoticons';
import 'tinymce/plugins/emoticons/js/emojis';

// Import skins
import 'tinymce/skins/ui/oxide/skin.min.css';
import 'tinymce/skins/ui/oxide-dark/skin.min.css';

import type { Editor as TinyMCEEditor } from 'tinymce';

const useStyles = makeStyles({
  container: {
    position: 'relative',
    minHeight: '80px',
    borderRadius: tokens.borderRadiusMedium,
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    backgroundColor: tokens.colorNeutralBackground2,
    transitionProperty: 'border-color, background-color',
    transitionDuration: '0.15s',
    transitionTimingFunction: 'ease',
    overflow: 'hidden',
    ':hover': {
      border: `1px solid ${tokens.colorNeutralStroke1Hover}`,
    },
  },
  containerFocused: {
    border: `1px solid ${tokens.colorBrandStroke1}`,
    backgroundColor: tokens.colorNeutralBackground1,
  },
  containerDark: {
    backgroundColor: '#1a1a1a',
    border: '1px solid #333333',
    ':hover': {
      border: '1px solid #444444',
    },
  },
  containerDarkFocused: {
    border: `1px solid ${tokens.colorBrandStroke1}`,
    backgroundColor: '#1a1a1a',
  },
  placeholder: {
    position: 'absolute',
    top: '12px',
    left: '16px',
    color: tokens.colorNeutralForeground4,
    fontStyle: 'italic',
    pointerEvents: 'none',
    zIndex: 1,
  },
});

interface RichTextEditorProps {
  value: string;
  onChange: (value: string) => void;
  onBlur?: () => void;
  placeholder?: string;
  readOnly?: boolean;
  minHeight?: number;
}

export function RichTextEditor({
  value,
  onChange,
  onBlur,
  placeholder = 'Start typing...',
  readOnly = false,
  minHeight = 80,
}: RichTextEditorProps) {
  const styles = useStyles();
  const { theme } = useTheme();
  const editorRef = useRef<TinyMCEEditor | null>(null);
  const isDark = theme === 'dark';

  const handleEditorChange = useCallback((content: string) => {
    onChange(content);
  }, [onChange]);

  const handleBlur = useCallback(() => {
    onBlur?.();
  }, [onBlur]);

  // Check if content is empty (TinyMCE may have empty paragraphs)
  const isEmpty = !value || value === '<p></p>' || value === '<p><br></p>';

  return (
    <div className={mergeClasses(styles.container, isDark && styles.containerDark)}>
      {isEmpty && !readOnly && (
        <span className={styles.placeholder}>{placeholder}</span>
      )}
      <Editor
        licenseKey="gpl"
        onInit={(_evt, editor) => {
          editorRef.current = editor;
        }}
        value={value}
        onEditorChange={handleEditorChange}
        onBlur={handleBlur}
        disabled={readOnly}
        init={{
          // Appearance
          skin: isDark ? 'oxide-dark' : 'oxide',
          content_css: isDark ? 'dark' : 'default',

          // Inline mode for seamless editing
          inline: false,
          menubar: false,
          statusbar: false,

          // Size
          min_height: minHeight,
          max_height: 500,
          autoresize_bottom_margin: 0,

          // Toolbar - all features (trim based on SP compatibility)
          toolbar: 'bold italic underline strikethrough | forecolor backcolor | bullist numlist | link image table | emoticons charmap | removeformat code',
          toolbar_mode: 'sliding',

          // Plugins
          plugins: 'lists link autolink table autoresize image media code charmap emoticons',

          // Link settings
          link_default_target: '_blank',
          link_assume_external_targets: true,

          // Table settings
          table_responsive_width: true,
          table_default_attributes: {
            border: '1',
          },

          // Image settings (for pasting)
          images_upload_handler: () => Promise.reject('Image upload not supported'),
          paste_data_images: false,

          // Content styling to match Fluent UI
          content_style: `
            body {
              font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif;
              font-size: 14px;
              line-height: 1.5;
              margin: 12px 16px;
              padding: 0;
              color: ${isDark ? '#ffffff' : '#242424'};
              background: transparent;
            }
            p { margin: 0 0 8px 0; }
            p:last-child { margin-bottom: 0; }
            ul, ol { margin: 0 0 8px 0; padding-left: 24px; }
            a { color: #0078d4; text-decoration: none; }
            a:hover { text-decoration: underline; }
            table { border-collapse: collapse; width: 100%; margin: 8px 0; }
            th, td { border: 1px solid ${isDark ? '#444' : '#d1d1d1'}; padding: 8px; }
            th { background: ${isDark ? '#2a2a2a' : '#f5f5f5'}; font-weight: 600; }
          `,

          // Keyboard shortcuts
          setup: (editor) => {
            // Save on Ctrl/Cmd + Enter
            editor.addShortcut('meta+enter', 'Blur editor', () => {
              editor.fire('blur');
            });
            editor.addShortcut('ctrl+enter', 'Blur editor', () => {
              editor.fire('blur');
            });
          },
        }}
      />
    </div>
  );
}
