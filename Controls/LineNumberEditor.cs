using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace MiniSolidworkAutomator.Controls
{
    /// <summary>
    /// A code editor with line numbers, bookmarks, and advanced features
    /// </summary>
    public class LineNumberEditor : UserControl
    {
        private RichTextBox editor = null!;
        private Panel lineNumberPanel = null!;
        private Panel bookmarkPanel = null!;
        
        private List<int> bookmarkedLines = new List<int>();
        private int lineHeight = 16;
        private Font codeFont = null!;
        
        // Theme colors
        private static readonly Color DarkBackground = Color.FromArgb(30, 30, 30);
        private static readonly Color LineNumberBg = Color.FromArgb(40, 40, 40);
        private static readonly Color LineNumberFg = Color.FromArgb(130, 130, 130);
        private static readonly Color CurrentLineBg = Color.FromArgb(50, 50, 50);
        private static readonly Color BookmarkColor = Color.FromArgb(33, 150, 243);
        private static readonly Color TextWhite = Color.FromArgb(212, 212, 212);

        public new event EventHandler? TextChanged;
        public event EventHandler? BookmarksChanged;

        [DllImport("uxtheme.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        private static extern int SetWindowTheme(IntPtr hwnd, string? pszSubAppName, string? pszSubIdList);

        private static void ApplyDarkScrollbar(Control control)
        {
            if (control.IsHandleCreated)
            {
                SetWindowTheme(control.Handle, "DarkMode_Explorer", null);
            }
            control.HandleCreated += (s, e) => SetWindowTheme(control.Handle, "DarkMode_Explorer", null);
        }

        public new string Text
        {
            get => editor.Text;
            set => editor.Text = value;
        }

        public bool ReadOnly
        {
            get => editor.ReadOnly;
            set => editor.ReadOnly = value;
        }

        public RichTextBox EditorControl => editor;

        public List<int> Bookmarks => bookmarkedLines;

        public int CurrentLine
        {
            get
            {
                int index = editor.SelectionStart;
                return editor.GetLineFromCharIndex(index) + 1;
            }
        }

        public LineNumberEditor()
        {
            InitializeComponents();
        }

        private void InitializeComponents()
        {
            this.BackColor = DarkBackground;
            this.Padding = new Padding(0);

            codeFont = new Font("Cascadia Code", 11);
            lineHeight = (int)(codeFont.GetHeight() * 1.2f);

            // Bookmark panel (leftmost)
            bookmarkPanel = new Panel
            {
                Dock = DockStyle.Left,
                Width = 16,
                BackColor = LineNumberBg
            };
            bookmarkPanel.Paint += BookmarkPanel_Paint;
            bookmarkPanel.MouseClick += BookmarkPanel_MouseClick;

            // Line number panel
            lineNumberPanel = new Panel
            {
                Dock = DockStyle.Left,
                Width = 50,
                BackColor = LineNumberBg
            };
            lineNumberPanel.Paint += LineNumberPanel_Paint;

            // Main editor
            editor = new RichTextBox
            {
                Dock = DockStyle.Fill,
                Font = codeFont,
                BackColor = DarkBackground,
                ForeColor = TextWhite,
                BorderStyle = BorderStyle.None,
                AcceptsTab = true,
                WordWrap = false,
                ScrollBars = RichTextBoxScrollBars.Both,
                DetectUrls = false
            };
            
            editor.TextChanged += (s, e) => 
            {
                RefreshLineNumbers();
                TextChanged?.Invoke(this, e);
            };
            editor.VScroll += (s, e) => RefreshLineNumbers();
            editor.Resize += (s, e) => RefreshLineNumbers();
            editor.SelectionChanged += (s, e) => RefreshLineNumbers();
            editor.KeyDown += Editor_KeyDown;
            ApplyDarkScrollbar(editor);

            // Add controls in order
            this.Controls.Add(editor);
            this.Controls.Add(lineNumberPanel);
            this.Controls.Add(bookmarkPanel);
        }

        private void Editor_KeyDown(object? sender, KeyEventArgs e)
        {
            // Ctrl+B to toggle bookmark
            if (e.Control && e.KeyCode == Keys.B)
            {
                ToggleBookmark(CurrentLine);
                e.Handled = true;
            }
            // F2 to go to next bookmark
            else if (e.KeyCode == Keys.F2 && !e.Control)
            {
                GoToNextBookmark();
                e.Handled = true;
            }
            // Shift+F2 to go to previous bookmark
            else if (e.KeyCode == Keys.F2 && e.Shift)
            {
                GoToPreviousBookmark();
                e.Handled = true;
            }
        }

        public void ToggleBookmark(int line)
        {
            if (bookmarkedLines.Contains(line))
                bookmarkedLines.Remove(line);
            else
                bookmarkedLines.Add(line);
            
            bookmarkedLines.Sort();
            RefreshLineNumbers();
            BookmarksChanged?.Invoke(this, EventArgs.Empty);
        }

        public void GoToNextBookmark()
        {
            if (bookmarkedLines.Count == 0) return;
            int current = CurrentLine;
            int next = bookmarkedLines.FirstOrDefault(b => b > current);
            if (next == 0) next = bookmarkedLines[0]; // Wrap around
            GoToLine(next);
        }

        public void GoToPreviousBookmark()
        {
            if (bookmarkedLines.Count == 0) return;
            int current = CurrentLine;
            int prev = bookmarkedLines.LastOrDefault(b => b < current);
            if (prev == 0) prev = bookmarkedLines[^1]; // Wrap around
            GoToLine(prev);
        }

        public void GoToLine(int lineNumber)
        {
            if (lineNumber < 1) lineNumber = 1;
            int lineCount = editor.Lines.Length;
            if (lineNumber > lineCount) lineNumber = lineCount;

            int charIndex = editor.GetFirstCharIndexFromLine(lineNumber - 1);
            if (charIndex >= 0)
            {
                editor.SelectionStart = charIndex;
                editor.SelectionLength = 0;
                editor.ScrollToCaret();
                editor.Focus();
            }
        }

        public void ClearBookmarks()
        {
            bookmarkedLines.Clear();
            RefreshLineNumbers();
            BookmarksChanged?.Invoke(this, EventArgs.Empty);
        }

        private void RefreshLineNumbers()
        {
            lineNumberPanel.Invalidate();
            bookmarkPanel.Invalidate();
        }

        private void LineNumberPanel_Paint(object? sender, PaintEventArgs e)
        {
            e.Graphics.Clear(LineNumberBg);
            
            if (editor.Lines.Length == 0) return;

            int firstVisibleChar = editor.GetCharIndexFromPosition(new Point(0, 0));
            int firstVisibleLine = editor.GetLineFromCharIndex(firstVisibleChar) + 1;
            
            int lastVisibleChar = editor.GetCharIndexFromPosition(new Point(0, editor.ClientSize.Height));
            int lastVisibleLine = editor.GetLineFromCharIndex(lastVisibleChar) + 1;
            
            int currentLine = CurrentLine;

            using var font = new Font("Consolas", 9);
            using var brush = new SolidBrush(LineNumberFg);
            using var currentBrush = new SolidBrush(TextWhite);
            using var highlightBrush = new SolidBrush(CurrentLineBg);
            
            for (int line = firstVisibleLine; line <= lastVisibleLine && line <= editor.Lines.Length; line++)
            {
                int charIndex = editor.GetFirstCharIndexFromLine(line - 1);
                if (charIndex < 0) continue;
                
                Point pos = editor.GetPositionFromCharIndex(charIndex);
                int y = pos.Y;

                // Highlight current line
                if (line == currentLine)
                {
                    e.Graphics.FillRectangle(highlightBrush, 0, y, lineNumberPanel.Width, lineHeight);
                }

                string lineNum = line.ToString();
                var textBrush = line == currentLine ? currentBrush : brush;
                
                SizeF size = e.Graphics.MeasureString(lineNum, font);
                float x = lineNumberPanel.Width - size.Width - 5;
                e.Graphics.DrawString(lineNum, font, textBrush, x, y + 2);
            }
        }

        private void BookmarkPanel_Paint(object? sender, PaintEventArgs e)
        {
            e.Graphics.Clear(LineNumberBg);
            
            if (editor.Lines.Length == 0 || bookmarkedLines.Count == 0) return;

            using var brush = new SolidBrush(BookmarkColor);
            
            foreach (int line in bookmarkedLines)
            {
                if (line > editor.Lines.Length) continue;
                
                int charIndex = editor.GetFirstCharIndexFromLine(line - 1);
                if (charIndex < 0) continue;
                
                Point pos = editor.GetPositionFromCharIndex(charIndex);
                int y = pos.Y + 3;

                // Draw bookmark indicator (circle)
                e.Graphics.FillEllipse(brush, 3, y, 10, 10);
            }
        }

        private void BookmarkPanel_MouseClick(object? sender, MouseEventArgs e)
        {
            // Find which line was clicked
            int charIndex = editor.GetCharIndexFromPosition(new Point(0, e.Y));
            int line = editor.GetLineFromCharIndex(charIndex) + 1;
            ToggleBookmark(line);
        }

        // Search functionality
        public int FindText(string searchText, bool matchCase, bool searchUp)
        {
            if (string.IsNullOrEmpty(searchText)) return -1;

            RichTextBoxFinds options = RichTextBoxFinds.None;
            if (matchCase) options |= RichTextBoxFinds.MatchCase;
            if (searchUp) options |= RichTextBoxFinds.Reverse;

            int start = searchUp ? 0 : editor.SelectionStart + editor.SelectionLength;
            int end = searchUp ? editor.SelectionStart : editor.TextLength;

            int index = editor.Find(searchText, start, end, options);
            
            // Wrap around
            if (index < 0)
            {
                start = searchUp ? editor.SelectionStart : 0;
                end = searchUp ? editor.TextLength : editor.SelectionStart;
                index = editor.Find(searchText, start, end, options);
            }

            if (index >= 0)
            {
                editor.SelectionStart = index;
                editor.SelectionLength = searchText.Length;
                editor.ScrollToCaret();
                editor.Focus();
            }

            return index;
        }

        public int ReplaceText(string searchText, string replaceText, bool matchCase)
        {
            if (editor.SelectedText.Equals(searchText, matchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase))
            {
                editor.SelectedText = replaceText;
                return 1;
            }
            return FindText(searchText, matchCase, false) >= 0 ? 0 : -1;
        }

        public int ReplaceAll(string searchText, string replaceText, bool matchCase)
        {
            if (string.IsNullOrEmpty(searchText)) return 0;

            var comparison = matchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
            string text = editor.Text;
            int count = 0;
            int index = 0;

            var newText = new System.Text.StringBuilder();
            int lastIndex = 0;

            while ((index = text.IndexOf(searchText, lastIndex, comparison)) >= 0)
            {
                newText.Append(text.Substring(lastIndex, index - lastIndex));
                newText.Append(replaceText);
                lastIndex = index + searchText.Length;
                count++;
            }
            newText.Append(text.Substring(lastIndex));

            if (count > 0)
            {
                editor.Text = newText.ToString();
            }

            return count;
        }
    }
}
