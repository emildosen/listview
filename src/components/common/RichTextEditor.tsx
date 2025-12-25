import { useRef, useCallback, useEffect, useState, useMemo } from 'react';
import { useEditor, EditorContent } from '@tiptap/react';
import { BubbleMenu } from '@tiptap/react/menus';
import StarterKit from '@tiptap/starter-kit';
import Placeholder from '@tiptap/extension-placeholder';
import Link from '@tiptap/extension-link';
import { TextStyle } from '@tiptap/extension-text-style';
import { Color } from '@tiptap/extension-color';
import Highlight from '@tiptap/extension-highlight';
import Underline from '@tiptap/extension-underline';
import { Table } from '@tiptap/extension-table';
import { TableRow } from '@tiptap/extension-table-row';
import { TableCell } from '@tiptap/extension-table-cell';
import { TableHeader } from '@tiptap/extension-table-header';
import { makeStyles, tokens, mergeClasses, Tooltip } from '@fluentui/react-components';
import {
  TextBoldRegular,
  TextItalicRegular,
  TextUnderlineRegular,
  TextStrikethroughRegular,
  LinkRegular,
  TextHeader1Regular,
  TextHeader2Regular,
  TextBulletListLtrRegular,
  TextNumberListLtrRegular,
  HighlightRegular,
  DismissRegular,
  TableRegular,
} from '@fluentui/react-icons';
import { useTheme } from '../../contexts/ThemeContext';

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
  containerDark: {
    backgroundColor: '#1a1a1a',
    border: '1px solid #333333',
    ':hover': {
      border: '1px solid #444444',
    },
  },
  editorContent: {
    fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
    fontSize: '15px',
    lineHeight: '1.5',
    padding: '12px 16px',
    outline: 'none',
    minHeight: 'var(--editor-min-height, 80px)',
    maxHeight: '600px',
    overflowY: 'auto',
    color: tokens.colorNeutralForeground1,
    '& .ProseMirror': {
      outline: 'none',
      minHeight: 'inherit',
    },
    '& .ProseMirror p': {
      margin: '0 0 8px 0',
    },
    '& .ProseMirror p:last-child': {
      marginBottom: 0,
    },
    '& .ProseMirror ul, & .ProseMirror ol': {
      margin: '0 0 8px 0',
      paddingLeft: '24px',
    },
    '& .ProseMirror a': {
      color: '#0078d4',
      textDecoration: 'none',
      ':hover': {
        textDecoration: 'underline',
      },
    },
    '& .ProseMirror h1': {
      fontSize: '1.5em',
      fontWeight: 600,
      margin: '0 0 12px 0',
    },
    '& .ProseMirror h2': {
      fontSize: '1.25em',
      fontWeight: 600,
      margin: '0 0 10px 0',
    },
    '& .ProseMirror mark': {
      backgroundColor: '#fff3bf',
      borderRadius: '2px',
      padding: '0 2px',
    },
    '& .ProseMirror p.is-editor-empty:first-child::before': {
      content: 'attr(data-placeholder)',
      float: 'left',
      color: tokens.colorNeutralForeground4,
      fontStyle: 'italic',
      pointerEvents: 'none',
      height: 0,
    },
    '& .ProseMirror table': {
      borderCollapse: 'collapse',
      margin: '8px 0',
      width: '100%',
      tableLayout: 'fixed',
    },
    '& .ProseMirror th, & .ProseMirror td': {
      border: '1px solid #d0d0d0',
      padding: '6px 10px',
      verticalAlign: 'top',
      minWidth: '50px',
      position: 'relative',
    },
    '& .ProseMirror th': {
      backgroundColor: tokens.colorNeutralBackground3,
      fontWeight: tokens.fontWeightSemibold,
      textAlign: 'left',
    },
    '& .ProseMirror .selectedCell::after': {
      backgroundColor: 'rgba(0, 120, 212, 0.15)',
      content: '""',
      left: 0,
      right: 0,
      top: 0,
      bottom: 0,
      pointerEvents: 'none',
      position: 'absolute',
    },
  },
  editorContentDark: {
    color: '#ffffff',
    '& .ProseMirror th, & .ProseMirror td': {
      border: '1px solid #444444',
    },
    '& .ProseMirror th': {
      backgroundColor: '#2a2a2a',
    },
  },
  bubbleMenu: {
    display: 'flex',
    alignItems: 'center',
    gap: '2px',
    padding: '4px',
    backgroundColor: tokens.colorNeutralBackground1,
    borderRadius: tokens.borderRadiusMedium,
    boxShadow: tokens.shadow16,
    border: `1px solid ${tokens.colorNeutralStroke1}`,
  },
  bubbleMenuDark: {
    backgroundColor: '#2a2a2a',
    border: '1px solid #444444',
  },
  bubbleButton: {
    minWidth: '28px',
    height: '28px',
    padding: '4px',
    borderRadius: tokens.borderRadiusSmall,
    backgroundColor: 'transparent',
    border: 'none',
    cursor: 'pointer',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    color: tokens.colorNeutralForeground1,
    ':hover': {
      backgroundColor: tokens.colorNeutralBackground3,
    },
  },
  bubbleButtonActive: {
    backgroundColor: tokens.colorNeutralBackground3,
    color: tokens.colorBrandForeground1,
  },
  bubbleButtonDark: {
    color: '#ffffff',
    ':hover': {
      backgroundColor: '#3a3a3a',
    },
  },
  bubbleButtonActiveDark: {
    backgroundColor: '#3a3a3a',
    color: '#60cdff',
  },
  colorPicker: {
    display: 'flex',
    gap: '4px',
    padding: '8px',
    backgroundColor: tokens.colorNeutralBackground1,
    borderRadius: tokens.borderRadiusMedium,
    boxShadow: tokens.shadow16,
    border: `1px solid ${tokens.colorNeutralStroke1}`,
  },
  colorPickerDark: {
    backgroundColor: '#2a2a2a',
    border: '1px solid #444444',
  },
  colorSwatch: {
    width: '20px',
    height: '20px',
    borderRadius: '4px',
    cursor: 'pointer',
    border: '1px solid rgba(0,0,0,0.1)',
    ':hover': {
      transform: 'scale(1.1)',
    },
  },
  toolbar: {
    display: 'flex',
    alignItems: 'center',
    gap: '1px',
    padding: '4px 8px',
    borderBottom: `1px solid ${tokens.colorNeutralStroke1}`,
    backgroundColor: tokens.colorNeutralBackground3,
  },
  toolbarDark: {
    backgroundColor: '#252525',
    borderBottom: '1px solid #333333',
  },
  toolbarDivider: {
    width: '1px',
    height: '16px',
    backgroundColor: tokens.colorNeutralStroke2,
    margin: '0 4px',
  },
  toolbarDividerDark: {
    backgroundColor: '#444444',
  },
  toolbarButton: {
    minWidth: '24px',
    height: '24px',
    padding: '2px',
    borderRadius: tokens.borderRadiusSmall,
    backgroundColor: 'transparent',
    border: 'none',
    cursor: 'pointer',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    color: tokens.colorNeutralForeground2,
    fontSize: '16px',
    ':hover': {
      backgroundColor: tokens.colorNeutralBackground1Hover,
      color: tokens.colorNeutralForeground1,
    },
  },
  toolbarButtonDark: {
    color: '#aaaaaa',
    ':hover': {
      backgroundColor: '#3a3a3a',
      color: '#ffffff',
    },
  },
  toolbarButtonActive: {
    backgroundColor: tokens.colorNeutralBackground1,
    color: tokens.colorBrandForeground1,
  },
  toolbarButtonActiveDark: {
    backgroundColor: '#3a3a3a',
    color: '#60cdff',
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
    .replace(/<\/p>\s*<p[^>]*>/gi, '\n')  // Paragraph breaks to newlines
    .replace(/<br\s*\/?>/gi, '\n')         // BR tags to newlines
    .replace(/<[^>]*>/g, '')               // Strip remaining HTML tags
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

// Highlight colors - light (saved to SharePoint) and dark (display only)
const highlightColorsLight = [
  '#fff3bf', '#ffd6d6', '#ffe0c2', '#d4edda', '#cce5ff',
  '#e2d9f3', '#f8d7da', '#d1ecf1', '#fff3cd', '#e2e3e5',
];

const highlightColorsDark = [
  '#5c4800', '#5c2828', '#5c3a1a', '#1e4a28', '#1e3a5c',
  '#3d2e5c', '#4d2828', '#1e4a5c', '#4d4200', '#3a3a3a',
];

// Create bidirectional color maps
const lightToDarkMap = new Map<string, string>();
const darkToLightMap = new Map<string, string>();
highlightColorsLight.forEach((light, i) => {
  const dark = highlightColorsDark[i];
  lightToDarkMap.set(light.toLowerCase(), dark);
  darkToLightMap.set(dark.toLowerCase(), light);
});

// Convert RGB to hex
function rgbToHex(r: number, g: number, b: number): string {
  return '#' + [r, g, b].map(x => x.toString(16).padStart(2, '0')).join('');
}

// Translate highlight colors in HTML content
function translateHighlightColors(html: string, toDark: boolean): string {
  const map = toDark ? lightToDarkMap : darkToLightMap;
  return html
    // Handle hex background-color
    .replace(/background-color:\s*(#[a-fA-F0-9]{6})/gi, (match, color) => {
      const translated = map.get(color.toLowerCase());
      return translated ? `background-color: ${translated}` : match;
    })
    // Handle rgb() background-color
    .replace(/background-color:\s*rgb\((\d+),\s*(\d+),\s*(\d+)\)/gi, (match, r, g, b) => {
      const hex = rgbToHex(parseInt(r), parseInt(g), parseInt(b));
      const translated = map.get(hex.toLowerCase());
      return translated ? `background-color: ${translated}` : match;
    })
    // Handle data-color attribute
    .replace(/data-color="(#[a-fA-F0-9]{6})"/gi, (match, color) => {
      const translated = map.get(color.toLowerCase());
      return translated ? `data-color="${translated}"` : match;
    });
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
  const isDark = theme === 'dark';
  const lastExternalValue = useRef(value);
  const prevThemeRef = useRef(theme);
  const [showHighlightPicker, setShowHighlightPicker] = useState(false);

  // Get display content (translate to dark colors if in dark mode)
  const getDisplayContent = useCallback((html: string) => {
    if (!showToolbar) return plainTextToHtml(html);
    return isDark ? translateHighlightColors(html, true) : html;
  }, [showToolbar, isDark]);


  // Configure extensions based on mode - memoized to prevent duplicate extension warnings
  const extensions = useMemo(() => [
    StarterKit.configure({
      heading: showToolbar ? { levels: [1, 2] } : false,
      bulletList: showToolbar ? {} : false,
      orderedList: showToolbar ? {} : false,
    }),
    Placeholder.configure({
      placeholder,
    }),
    Link.configure({
      openOnClick: false,
      HTMLAttributes: {
        target: '_blank',
        rel: 'noopener noreferrer',
      },
    }),
    TextStyle,
    Color,
    Highlight.configure({
      multicolor: true,
    }),
    Underline,
    ...(showToolbar ? [
      Table.configure({
        resizable: false,
      }),
      TableRow,
      TableHeader,
      TableCell,
    ] : []),
  ], [showToolbar, placeholder]);

  const editor = useEditor({
    extensions,
    content: getDisplayContent(value),
    editable: !readOnly,
    editorProps: {
      attributes: {
        class: 'tiptap-editor',
      },
    },
  });

  // Handle blur separately to avoid stale closure issues with useEditor
  useEffect(() => {
    if (!editor) return;

    const handleBlur = () => {
      const html = editor.getHTML();
      // Always translate dark colors back to light for saving
      let outputValue: string;
      if (!showToolbar) {
        outputValue = htmlToPlainText(html);
      } else {
        // Translate any dark colors to light for saving to SharePoint
        outputValue = translateHighlightColors(html, false);
      }

      if (outputValue !== lastExternalValue.current) {
        lastExternalValue.current = outputValue;
        onChange(outputValue);
      }

      onBlur?.();
    };

    editor.on('blur', handleBlur);
    return () => {
      editor.off('blur', handleBlur);
    };
  }, [editor, showToolbar, onChange, onBlur]);

  // Add keyboard shortcut for blur
  useEffect(() => {
    if (!editor) return;

    const handleKeyDown = (event: KeyboardEvent) => {
      if ((event.metaKey || event.ctrlKey) && event.key === 'Enter') {
        editor.commands.blur();
      }
    };

    const editorElement = editor.view.dom;
    editorElement.addEventListener('keydown', handleKeyDown);

    return () => {
      editorElement.removeEventListener('keydown', handleKeyDown);
    };
  }, [editor]);

  // Reset formatting on Enter (new line)
  useEffect(() => {
    if (!editor) return;

    const handleKeyUp = (event: KeyboardEvent) => {
      // On Enter, clear all marks (bold, italic, underline, highlight, etc.)
      if (event.key === 'Enter' && !event.shiftKey) {
        // Use setTimeout to ensure the new paragraph is created first
        setTimeout(() => {
          editor.commands.unsetAllMarks();
        }, 0);
      }
    };

    const editorElement = editor.view.dom;
    editorElement.addEventListener('keyup', handleKeyUp);

    return () => {
      editorElement.removeEventListener('keyup', handleKeyUp);
    };
  }, [editor]);

  // Update editor content when external value changes
  useEffect(() => {
    if (editor && value !== lastExternalValue.current) {
      editor.commands.setContent(getDisplayContent(value));
      lastExternalValue.current = value;
    }
  }, [value, editor, getDisplayContent]);

  // Re-render content when theme changes to translate highlight colors
  useEffect(() => {
    if (!editor || !showToolbar) return;
    if (prevThemeRef.current !== theme) {
      const html = editor.getHTML();
      // Translate colors: if switching to dark, convert light->dark; if switching to light, convert dark->light
      const toDark = theme === 'dark';
      const translatedContent = translateHighlightColors(html, toDark);
      editor.commands.setContent(translatedContent);
      prevThemeRef.current = theme;
    }
  }, [theme, editor, showToolbar]);

  const handleBubbleButtonClick = useCallback(
    (action: () => void) => (e: React.MouseEvent) => {
      e.preventDefault();
      action();
    },
    []
  );

  const handleSetLink = useCallback(() => {
    if (!editor) return;

    const previousUrl = editor.getAttributes('link').href;
    const url = window.prompt('Enter URL', previousUrl);

    if (url === null) return;

    if (url === '') {
      editor.chain().focus().extendMarkRange('link').unsetLink().run();
    } else {
      editor.chain().focus().extendMarkRange('link').setLink({ href: url }).run();
    }
  }, [editor]);

  const handleHighlightSelect = useCallback(
    (color: string) => {
      if (!editor) return;
      editor.chain().focus().toggleHighlight({ color }).run();
      setShowHighlightPicker(false);
    },
    [editor]
  );

  // Insert a 3x3 table
  const handleInsertTable = useCallback(() => {
    if (!editor) return;
    editor.chain().focus().insertTable({ rows: 3, cols: 3, withHeaderRow: true }).run();
  }, [editor]);

  if (!editor) {
    return null;
  }

  return (
    <div
      className={mergeClasses(styles.container, isDark && styles.containerDark)}
      tabIndex={0}
      onFocus={() => {
        // When container gets focus, forward it to the editor
        if (!editor.isFocused) {
          editor.commands.focus();
        }
      }}
    >
      {/* Command strip toolbar at top - only in rich text mode */}
      {showToolbar && (
        <div className={mergeClasses(styles.toolbar, isDark && styles.toolbarDark)}>
          <Tooltip content="Bold (⌘B)" relationship="label">
            <button
              type="button"
              onClick={() => editor.chain().focus().toggleBold().run()}
              className={mergeClasses(
                styles.toolbarButton,
                isDark && styles.toolbarButtonDark,
                editor.isActive('bold') && styles.toolbarButtonActive,
                editor.isActive('bold') && isDark && styles.toolbarButtonActiveDark
              )}
            >
              <TextBoldRegular />
            </button>
          </Tooltip>
          <Tooltip content="Italic (⌘I)" relationship="label">
            <button
              type="button"
              onClick={() => editor.chain().focus().toggleItalic().run()}
              className={mergeClasses(
                styles.toolbarButton,
                isDark && styles.toolbarButtonDark,
                editor.isActive('italic') && styles.toolbarButtonActive,
                editor.isActive('italic') && isDark && styles.toolbarButtonActiveDark
              )}
            >
              <TextItalicRegular />
            </button>
          </Tooltip>
          <Tooltip content="Underline (⌘U)" relationship="label">
            <button
              type="button"
              onClick={() => editor.chain().focus().toggleUnderline().run()}
              className={mergeClasses(
                styles.toolbarButton,
                isDark && styles.toolbarButtonDark,
                editor.isActive('underline') && styles.toolbarButtonActive,
                editor.isActive('underline') && isDark && styles.toolbarButtonActiveDark
              )}
            >
              <TextUnderlineRegular />
            </button>
          </Tooltip>

          <div className={mergeClasses(styles.toolbarDivider, isDark && styles.toolbarDividerDark)} />

          <Tooltip content="Heading 1" relationship="label">
            <button
              type="button"
              onClick={() => editor.chain().focus().toggleHeading({ level: 1 }).run()}
              className={mergeClasses(
                styles.toolbarButton,
                isDark && styles.toolbarButtonDark,
                editor.isActive('heading', { level: 1 }) && styles.toolbarButtonActive,
                editor.isActive('heading', { level: 1 }) && isDark && styles.toolbarButtonActiveDark
              )}
            >
              <TextHeader1Regular />
            </button>
          </Tooltip>
          <Tooltip content="Heading 2" relationship="label">
            <button
              type="button"
              onClick={() => editor.chain().focus().toggleHeading({ level: 2 }).run()}
              className={mergeClasses(
                styles.toolbarButton,
                isDark && styles.toolbarButtonDark,
                editor.isActive('heading', { level: 2 }) && styles.toolbarButtonActive,
                editor.isActive('heading', { level: 2 }) && isDark && styles.toolbarButtonActiveDark
              )}
            >
              <TextHeader2Regular />
            </button>
          </Tooltip>

          <div className={mergeClasses(styles.toolbarDivider, isDark && styles.toolbarDividerDark)} />

          <Tooltip content="Bullet list" relationship="label">
            <button
              type="button"
              onClick={() => editor.chain().focus().toggleBulletList().run()}
              className={mergeClasses(
                styles.toolbarButton,
                isDark && styles.toolbarButtonDark,
                editor.isActive('bulletList') && styles.toolbarButtonActive,
                editor.isActive('bulletList') && isDark && styles.toolbarButtonActiveDark
              )}
            >
              <TextBulletListLtrRegular />
            </button>
          </Tooltip>
          <Tooltip content="Numbered list" relationship="label">
            <button
              type="button"
              onClick={() => editor.chain().focus().toggleOrderedList().run()}
              className={mergeClasses(
                styles.toolbarButton,
                isDark && styles.toolbarButtonDark,
                editor.isActive('orderedList') && styles.toolbarButtonActive,
                editor.isActive('orderedList') && isDark && styles.toolbarButtonActiveDark
              )}
            >
              <TextNumberListLtrRegular />
            </button>
          </Tooltip>

          <div className={mergeClasses(styles.toolbarDivider, isDark && styles.toolbarDividerDark)} />

          <Tooltip content="Insert table" relationship="label">
            <button
              type="button"
              onClick={handleInsertTable}
              className={mergeClasses(
                styles.toolbarButton,
                isDark && styles.toolbarButtonDark
              )}
            >
              <TableRegular />
            </button>
          </Tooltip>
        </div>
      )}

      {/* Bubble menu for text selection - only in rich text mode */}
      {showToolbar && (
        <BubbleMenu
          editor={editor}
          className={mergeClasses(styles.bubbleMenu, isDark && styles.bubbleMenuDark)}
        >
          <Tooltip content="Bold" relationship="label">
            <button
              type="button"
              onClick={handleBubbleButtonClick(() => editor.chain().focus().toggleBold().run())}
              className={mergeClasses(
                styles.bubbleButton,
                isDark && styles.bubbleButtonDark,
                editor.isActive('bold') && styles.bubbleButtonActive,
                editor.isActive('bold') && isDark && styles.bubbleButtonActiveDark
              )}
            >
              <TextBoldRegular />
            </button>
          </Tooltip>
          <Tooltip content="Italic" relationship="label">
            <button
              type="button"
              onClick={handleBubbleButtonClick(() => editor.chain().focus().toggleItalic().run())}
              className={mergeClasses(
                styles.bubbleButton,
                isDark && styles.bubbleButtonDark,
                editor.isActive('italic') && styles.bubbleButtonActive,
                editor.isActive('italic') && isDark && styles.bubbleButtonActiveDark
              )}
            >
              <TextItalicRegular />
            </button>
          </Tooltip>
          <Tooltip content="Underline" relationship="label">
            <button
              type="button"
              onClick={handleBubbleButtonClick(() => editor.chain().focus().toggleUnderline().run())}
              className={mergeClasses(
                styles.bubbleButton,
                isDark && styles.bubbleButtonDark,
                editor.isActive('underline') && styles.bubbleButtonActive,
                editor.isActive('underline') && isDark && styles.bubbleButtonActiveDark
              )}
            >
              <TextUnderlineRegular />
            </button>
          </Tooltip>
          <Tooltip content="Strikethrough" relationship="label">
            <button
              type="button"
              onClick={handleBubbleButtonClick(() => editor.chain().focus().toggleStrike().run())}
              className={mergeClasses(
                styles.bubbleButton,
                isDark && styles.bubbleButtonDark,
                editor.isActive('strike') && styles.bubbleButtonActive,
                editor.isActive('strike') && isDark && styles.bubbleButtonActiveDark
              )}
            >
              <TextStrikethroughRegular />
            </button>
          </Tooltip>
          <Tooltip content="Highlight" relationship="label">
            <button
              type="button"
              onClick={handleBubbleButtonClick(() => setShowHighlightPicker(!showHighlightPicker))}
              className={mergeClasses(
                styles.bubbleButton,
                isDark && styles.bubbleButtonDark,
                editor.isActive('highlight') && styles.bubbleButtonActive,
                editor.isActive('highlight') && isDark && styles.bubbleButtonActiveDark
              )}
            >
              <HighlightRegular />
            </button>
          </Tooltip>
          <Tooltip content="Link" relationship="label">
            <button
              type="button"
              onClick={handleBubbleButtonClick(handleSetLink)}
              className={mergeClasses(
                styles.bubbleButton,
                isDark && styles.bubbleButtonDark,
                editor.isActive('link') && styles.bubbleButtonActive,
                editor.isActive('link') && isDark && styles.bubbleButtonActiveDark
              )}
            >
              <LinkRegular />
            </button>
          </Tooltip>
        </BubbleMenu>
      )}

      {/* Highlight color picker popup */}
      {showHighlightPicker && (
        <div
          className={mergeClasses(styles.colorPicker, isDark && styles.colorPickerDark)}
          style={{ position: 'fixed', zIndex: 1000002 }}
        >
          {(isDark ? highlightColorsDark : highlightColorsLight).map((color) => (
            <div
              key={color}
              className={styles.colorSwatch}
              style={{ backgroundColor: color }}
              onClick={() => handleHighlightSelect(color)}
            />
          ))}
          <button
            type="button"
            className={styles.bubbleButton}
            onClick={() => {
              editor.chain().focus().unsetHighlight().run();
              setShowHighlightPicker(false);
            }}
          >
            <DismissRegular />
          </button>
        </div>
      )}

      {/* Editor content */}
      <EditorContent
        editor={editor}
        className={mergeClasses(
          styles.editorContent,
          isDark && styles.editorContentDark
        )}
        style={{ '--editor-min-height': `${minHeight}px` } as React.CSSProperties}
      />
    </div>
  );
}
