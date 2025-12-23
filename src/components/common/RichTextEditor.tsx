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
import { Extension } from '@tiptap/core';
import Suggestion from '@tiptap/suggestion';
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
  TextParagraphRegular,
  HighlightRegular,
  DismissRegular,
  TableRegular,
} from '@fluentui/react-icons';
import { useTheme } from '../../contexts/ThemeContext';
import type { Editor, Range } from '@tiptap/core';
import type { ReactNode } from 'react';

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
    fontSize: '14px',
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
    '& .ProseMirror mark': {
      backgroundColor: '#5c4800',
    },
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

// Slash command items
interface SlashCommandItem {
  title: string;
  description: string;
  icon: ReactNode;
  command: (editor: Editor) => void;
}

const slashCommands: SlashCommandItem[] = [
  {
    title: 'Text',
    description: 'Plain text paragraph',
    icon: <TextParagraphRegular />,
    command: (editor) => editor.chain().focus().setParagraph().run(),
  },
  {
    title: 'Heading 1',
    description: 'Large heading',
    icon: <TextHeader1Regular />,
    command: (editor) => editor.chain().focus().toggleHeading({ level: 1 }).run(),
  },
  {
    title: 'Heading 2',
    description: 'Medium heading',
    icon: <TextHeader2Regular />,
    command: (editor) => editor.chain().focus().toggleHeading({ level: 2 }).run(),
  },
  {
    title: 'Bullet List',
    description: 'Unordered list',
    icon: <TextBulletListLtrRegular />,
    command: (editor) => editor.chain().focus().toggleBulletList().run(),
  },
  {
    title: 'Numbered List',
    description: 'Ordered list',
    icon: <TextNumberListLtrRegular />,
    command: (editor) => editor.chain().focus().toggleOrderedList().run(),
  },
];

// Create slash command extension
function createSlashCommandExtension(isDark: boolean) {
  return Extension.create({
    name: 'slashCommand',
    addOptions() {
      return {
        suggestion: {
          char: '/',
          command: ({ editor, range, props }: { editor: Editor; range: Range; props: SlashCommandItem }) => {
            props.command(editor);
            editor.chain().focus().deleteRange(range).run();
          },
          items: ({ query }: { query: string }) => {
            return slashCommands.filter((item) =>
              item.title.toLowerCase().includes(query.toLowerCase())
            );
          },
          render: () => {
            let component: HTMLDivElement | null = null;
            let selectedIndex = 0;
            let items: SlashCommandItem[] = [];

            const updateMenu = () => {
              if (!component) return;

              // Re-render the menu
              const menuHtml = items.map((item, index) => `
                <div class="slash-menu-item ${index === selectedIndex ? 'active' : ''}" data-index="${index}">
                  <span class="slash-menu-icon"></span>
                  <div class="slash-menu-text">
                    <span class="slash-menu-title">${item.title}</span>
                    <span class="slash-menu-desc">${item.description}</span>
                  </div>
                </div>
              `).join('');
              component.innerHTML = menuHtml;
            };

            return {
              onStart: (props: { clientRect: () => DOMRect | null; items: SlashCommandItem[]; command: (item: SlashCommandItem) => void }) => {
                items = props.items;
                selectedIndex = 0;

                component = document.createElement('div');
                component.className = `slash-command-menu ${isDark ? 'dark' : ''}`;
                component.style.cssText = `
                  position: fixed;
                  z-index: 1000001;
                  background: ${isDark ? '#1a1a1a' : '#ffffff'};
                  border: 1px solid ${isDark ? '#333333' : '#e0e0e0'};
                  border-radius: 6px;
                  box-shadow: 0 4px 16px rgba(0,0,0,0.16);
                  min-width: 200px;
                  max-height: 300px;
                  overflow-y: auto;
                  padding: 4px;
                `;

                const rect = props.clientRect?.();
                if (rect) {
                  component.style.left = `${rect.left}px`;
                  component.style.top = `${rect.bottom + 4}px`;
                }

                // Add styles for menu items
                const style = document.createElement('style');
                style.textContent = `
                  .slash-command-menu .slash-menu-item {
                    display: flex;
                    align-items: center;
                    gap: 12px;
                    padding: 8px 12px;
                    cursor: pointer;
                    border-radius: 4px;
                  }
                  .slash-command-menu .slash-menu-item:hover,
                  .slash-command-menu .slash-menu-item.active {
                    background: ${isDark ? '#252525' : '#f5f5f5'};
                  }
                  .slash-command-menu .slash-menu-text {
                    display: flex;
                    flex-direction: column;
                  }
                  .slash-command-menu .slash-menu-title {
                    font-size: 14px;
                    font-weight: 600;
                    color: ${isDark ? '#ffffff' : '#242424'};
                  }
                  .slash-command-menu .slash-menu-desc {
                    font-size: 12px;
                    color: ${isDark ? '#888888' : '#666666'};
                  }
                `;
                document.head.appendChild(style);

                updateMenu();

                // Add click handlers
                component.addEventListener('click', (e) => {
                  const target = (e.target as HTMLElement).closest('.slash-menu-item');
                  if (target) {
                    const index = parseInt(target.getAttribute('data-index') || '0');
                    props.command(items[index]);
                  }
                });

                document.body.appendChild(component);
              },
              onUpdate: (props: { clientRect: () => DOMRect | null; items: SlashCommandItem[] }) => {
                items = props.items;
                selectedIndex = 0;
                updateMenu();

                const rect = props.clientRect?.();
                if (rect && component) {
                  component.style.left = `${rect.left}px`;
                  component.style.top = `${rect.bottom + 4}px`;
                }
              },
              onKeyDown: (props: { event: KeyboardEvent }) => {
                if (props.event.key === 'ArrowUp') {
                  selectedIndex = (selectedIndex - 1 + items.length) % items.length;
                  updateMenu();
                  return true;
                }
                if (props.event.key === 'ArrowDown') {
                  selectedIndex = (selectedIndex + 1) % items.length;
                  updateMenu();
                  return true;
                }
                if (props.event.key === 'Enter') {
                  if (items[selectedIndex]) {
                    // We need to trigger the command through the suggestion plugin
                    return true;
                  }
                  return false;
                }
                if (props.event.key === 'Escape') {
                  return true;
                }
                return false;
              },
              onExit: () => {
                if (component) {
                  component.remove();
                  component = null;
                }
              },
            };
          },
        },
      };
    },
    addProseMirrorPlugins() {
      return [
        Suggestion({
          editor: this.editor,
          ...this.options.suggestion,
        }),
      ];
    },
  });
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

// Highlight colors
const highlightColors = [
  '#fff3bf', '#ffd6d6', '#ffe0c2', '#d4edda', '#cce5ff',
  '#e2d9f3', '#f8d7da', '#d1ecf1', '#fff3cd', '#e2e3e5',
];

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
  const [showHighlightPicker, setShowHighlightPicker] = useState(false);

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
      createSlashCommandExtension(isDark),
    ] : []),
  ], [showToolbar, placeholder, isDark]);

  const editor = useEditor({
    extensions,
    content: showToolbar ? value : plainTextToHtml(value),
    editable: !readOnly,
    editorProps: {
      attributes: {
        class: 'tiptap-editor',
      },
    },
    onBlur: ({ editor }) => {
      const html = editor.getHTML();
      const outputValue = showToolbar ? html : htmlToPlainText(html);

      if (outputValue !== lastExternalValue.current) {
        lastExternalValue.current = outputValue;
        onChange(outputValue);
      }

      onBlur?.();
    },
  });

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

  // Update editor content when external value changes
  useEffect(() => {
    if (editor && value !== lastExternalValue.current) {
      const displayValue = showToolbar ? value : plainTextToHtml(value);
      editor.commands.setContent(displayValue);
      lastExternalValue.current = value;
    }
  }, [value, editor, showToolbar]);

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
          {highlightColors.map((color) => (
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
