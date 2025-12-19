import { useRef, useCallback, useEffect, useLayoutEffect } from 'react';
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
    // Style TinyMCE to match container
    '& .tox-tinymce': {
      border: 'none !important',
      borderRadius: `${tokens.borderRadiusMedium} !important`,
    },
    '& .tox-editor-header': {
      borderBottom: 'none !important',
    },
  },
  containerDark: {
    backgroundColor: '#1a1a1a',
    border: '1px solid #333333',
    ':hover': {
      border: '1px solid #444444',
    },
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

// Convert HTML to plain text
function htmlToPlainText(html: string): string {
  return html
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<[^>]*>/g, '')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"');
}

// Convert plain text to HTML for display
function plainTextToHtml(text: string): string {
  return text.replace(/\n/g, '<br>');
}

// Inject global styles for TinyMCE floating elements (color pickers, etc.)
// These need high z-index to appear above Fluent UI Dialogs
const TINYMCE_GLOBAL_STYLES_ID = 'tinymce-global-styles';

function ensureTinyMCEGlobalStyles() {
  if (document.getElementById(TINYMCE_GLOBAL_STYLES_ID)) return;

  const style = document.createElement('style');
  style.id = TINYMCE_GLOBAL_STYLES_ID;
  style.textContent = `
    /* TinyMCE floating elements (color pickers, menus, dialogs) need high z-index */
    .tox-tinymce-aux {
      z-index: 1000001 !important;
    }

    /* Compact toolbar - reduce height only */
    .tox .tox-toolbar,
    .tox .tox-toolbar__overflow,
    .tox .tox-toolbar__primary {
      background: none !important;
      padding: 0 4px !important;
    }

    .tox .tox-editor-header {
      padding: 0 !important;
    }

    .tox .tox-toolbar-overlord {
      padding: 0 !important;
    }
  `;
  document.head.appendChild(style);
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
  const lastExternalValue = useRef(value);

  // Ensure global styles are injected for TinyMCE floating elements
  useLayoutEffect(() => {
    ensureTinyMCEGlobalStyles();
  }, []);

  // Update editor content when external value changes (e.g., from parent reset)
  useEffect(() => {
    if (editorRef.current && value !== lastExternalValue.current) {
      const displayValue = showToolbar ? value : plainTextToHtml(value);
      editorRef.current.setContent(displayValue);
      lastExternalValue.current = value;
    }
  }, [value, showToolbar]);

  const handleBlur = useCallback(() => {
    if (!editorRef.current) return;

    const content = editorRef.current.getContent();
    const outputValue = showToolbar ? content : htmlToPlainText(content);

    // Only trigger onChange if value actually changed
    if (outputValue !== lastExternalValue.current) {
      lastExternalValue.current = outputValue;
      onChange(outputValue);
    }

    onBlur?.();
  }, [onChange, onBlur, showToolbar]);

  // Check if content is empty
  const isEmpty = !value || value === '<p></p>' || value === '<p><br></p>';

  // Initial value for editor (only used on mount)
  const initialValue = showToolbar ? value : plainTextToHtml(value);

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
        initialValue={initialValue}
        onBlur={handleBlur}
        disabled={readOnly}
        init={{
          // Appearance - use oxide-dark skin for dark mode
          skin: isDark ? 'oxide-dark' : 'oxide',
          content_css: false,

          // Editor mode
          inline: false,
          menubar: false,
          statusbar: false,

          // Size
          min_height: minHeight,
          max_height: 600,
          autoresize_bottom_margin: 0,

          // Toolbar - SP-compatible features (no tables/images)
          toolbar: showToolbar
            ? 'bold italic underline strikethrough | forecolor backcolor | bullist numlist | link | emoticons charmap | removeformat inlinecode'
            : false,
          toolbar_mode: 'sliding',

          // Plugins
          plugins: showToolbar
            ? 'lists link autolink autoresize charmap emoticons'
            : 'autoresize',

          // For plain text mode, don't wrap in <p> tags
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
            code {
              font-family: 'Consolas', 'Monaco', monospace;
              font-size: 0.9em;
              background-color: ${isDark ? '#2d2d2d' : '#f0f0f0'};
              padding: 2px 6px;
              border-radius: 3px;
            }
          `,

          // Keyboard shortcuts and custom buttons
          setup: (editor) => {
            editor.addShortcut('meta+enter', 'Blur editor', () => {
              editor.fire('blur');
            });
            editor.addShortcut('ctrl+enter', 'Blur editor', () => {
              editor.fire('blur');
            });

            // Custom inline code button (toggles <code> tag)
            editor.ui.registry.addToggleButton('inlinecode', {
              icon: 'sourcecode',
              tooltip: 'Inline code',
              onAction: () => {
                editor.execCommand('mceToggleFormat', false, 'code');
              },
              onSetup: (api) => {
                const changed = editor.formatter.formatChanged('code', (state) => {
                  api.setActive(state);
                });
                return () => changed.unbind();
              },
            });
          },
        }}
      />
    </div>
  );
}
