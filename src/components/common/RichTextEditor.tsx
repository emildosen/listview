import { useRef, useCallback } from 'react';
import { Editor } from '@tinymce/tinymce-react';
import { makeStyles, tokens, mergeClasses } from '@fluentui/react-components';
import { useTheme } from '../../contexts/ThemeContext';

// Import TinyMCE core and required modules for self-hosted bundling
import 'tinymce/tinymce';
import 'tinymce/themes/silver';
import 'tinymce/icons/default';
import 'tinymce/models/dom';

// Import plugins (no table/image for SP compatibility)
import 'tinymce/plugins/lists';
import 'tinymce/plugins/link';
import 'tinymce/plugins/autolink';
import 'tinymce/plugins/autoresize';
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
  showToolbar?: boolean;
}

export function RichTextEditor({
  value,
  onChange,
  onBlur,
  placeholder = 'Start typing...',
  readOnly = false,
  minHeight = 80,
  showToolbar = true,
}: RichTextEditorProps) {
  const styles = useStyles();
  const { theme } = useTheme();
  const editorRef = useRef<TinyMCEEditor | null>(null);
  const isDark = theme === 'dark';

  const handleEditorChange = useCallback((content: string) => {
    if (showToolbar) {
      // Rich text mode - keep HTML
      onChange(content);
    } else {
      // Plain text mode - strip HTML, convert <br> to newlines
      const plainText = content
        .replace(/<br\s*\/?>/gi, '\n')
        .replace(/<[^>]*>/g, '')
        .replace(/&nbsp;/g, ' ')
        .replace(/&amp;/g, '&')
        .replace(/&lt;/g, '<')
        .replace(/&gt;/g, '>')
        .replace(/&quot;/g, '"');
      onChange(plainText);
    }
  }, [onChange, showToolbar]);

  const handleBlur = useCallback(() => {
    onBlur?.();
  }, [onBlur]);

  // Check if content is empty (TinyMCE may have empty paragraphs)
  const isEmpty = !value || value === '<p></p>' || value === '<p><br></p>';

  // For plain text mode, convert newlines to <br> for display in editor
  const displayValue = showToolbar ? value : value.replace(/\n/g, '<br>');

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
        value={displayValue}
        onEditorChange={handleEditorChange}
        onBlur={handleBlur}
        disabled={readOnly}
        init={{
          // Appearance - use oxide-dark skin for dark mode
          skin: isDark ? 'oxide-dark' : 'oxide',
          content_css: false, // Don't load external CSS, use content_style instead

          // Inline mode for seamless editing
          inline: false,
          menubar: false,
          statusbar: false,

          // Size
          min_height: minHeight,
          max_height: 600,
          autoresize_bottom_margin: 0,

          // Toolbar - SP-compatible features (no tables/images)
          toolbar: showToolbar
            ? 'bold italic underline strikethrough | forecolor backcolor | bullist numlist | link | emoticons charmap | removeformat code'
            : false,
          toolbar_mode: 'sliding',

          // Plugins (no table/image for SP compatibility)
          plugins: showToolbar
            ? 'lists link autolink autoresize code charmap emoticons'
            : 'autoresize',

          // For plain text mode (no toolbar), don't wrap in <p> tags
          forced_root_block: showToolbar ? 'p' : '',
          newline_behavior: showToolbar ? 'default' : 'linebreak',

          // Link settings
          link_default_target: '_blank',
          link_assume_external_targets: true,

          // Content styling to match Fluent UI theme
          content_style: `
            body {
              font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif;
              font-size: 14px;
              line-height: 1.5;
              margin: 12px 16px;
              padding: 0;
              color: ${isDark ? '#ffffff' : '#242424'};
              background-color: ${isDark ? '#1a1a1a' : '#ffffff'};
            }
            p { margin: 0 0 8px 0; }
            p:last-child { margin-bottom: 0; }
            ul, ol { margin: 0 0 8px 0; padding-left: 24px; }
            a { color: #0078d4; text-decoration: none; }
            a:hover { text-decoration: underline; }
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
