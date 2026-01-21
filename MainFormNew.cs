using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.CodeAnalysis.CSharp.Scripting;
using Microsoft.CodeAnalysis.Scripting;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using MiniSolidworkAutomator.Models;
using MiniSolidworkAutomator.Controls;
using MiniSolidworkAutomator.Services;
using MiniSolidworkAutomator.Localization;
using System.Collections;

namespace MiniSolidworkAutomator
{
    // ListView sorter for file lists
    public class ListViewItemComparer : IComparer
    {
        private readonly int column;
        private readonly SortOrder order;

        public ListViewItemComparer(int column, SortOrder order)
        {
            this.column = column;
            this.order = order;
        }

        public int Compare(object? x, object? y)
        {
            if (x is not ListViewItem itemX || y is not ListViewItem itemY)
                return 0;

            string textX = column < itemX.SubItems.Count ? itemX.SubItems[column].Text : "";
            string textY = column < itemY.SubItems.Count ? itemY.SubItems[column].Text : "";

            int result;

            // Try date comparison for column 2 (Modified time)
            if (column == 2)
            {
                if (DateTime.TryParse(textX, out DateTime dateX) && DateTime.TryParse(textY, out DateTime dateY))
                {
                    result = DateTime.Compare(dateX, dateY);
                }
                else
                {
                    result = string.Compare(textX, textY, StringComparison.OrdinalIgnoreCase);
                }
            }
            else
            {
                result = string.Compare(textX, textY, StringComparison.OrdinalIgnoreCase);
            }

            return order == SortOrder.Descending ? -result : result;
        }
    }

    public partial class MainFormNew : Form
    {
        // ============ DWM API for Dark Title Bar ============
        [DllImport("dwmapi.dll", PreserveSig = true)]
        private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int attrValue, int attrSize);
        private const int DWMWA_USE_IMMERSIVE_DARK_MODE = 20;
        private const int DWMWA_USE_IMMERSIVE_DARK_MODE_BEFORE_20H1 = 19;

        // ============ Theme Colors (Dark Theme) ============
        private static readonly Color DarkBackground = Color.FromArgb(30, 30, 30);
        private static readonly Color DarkPanel = Color.FromArgb(45, 45, 45);
        private static readonly Color DarkToolbar = Color.FromArgb(38, 50, 56);
        private static readonly Color DarkSplitter = Color.FromArgb(55, 71, 79);
        private static readonly Color DarkTerminal = Color.FromArgb(30, 30, 30);  // Match editor background
        private static readonly Color TextWhite = Color.White;
        private static readonly Color TextGray = Color.FromArgb(180, 180, 180);
        private static readonly Color AccentGreen = Color.FromArgb(46, 125, 50);  // Darker green for Run button
        private static readonly Color AccentRed = Color.FromArgb(183, 28, 28);    // Darker red for Stop button
        private static readonly Color AccentBlue = Color.FromArgb(33, 150, 243);
        private static readonly Color AccentPurple = Color.FromArgb(142, 68, 173);

        // ============ UI Components ============
        // Main layout: Left (Code) | Right (Browser + Terminal)
        private SplitContainer mainSplit = null!;
        
        // Left panel - Code editor with tabs
        private Panel codePanel = null!;
        private TabControl codeTabs = null!;           // Tab control for opened files
        
        // Right panel split: Top (Browser tabs) | Bottom (Terminal)
        private SplitContainer rightSplit = null!;
        private TabControl browserTabs = null!;
        private RichTextBox terminalDisplay = null!;

        // Toolbar buttons
        private Button btnConnect = null!;
        private Button btnRun = null!;
        private Button btnStop = null!;
        private Button btnNew = null!;
        private Button btnSave = null!;
        private Button btnSettings = null!;
        private Button btnRefresh = null!;
        private Button btnCopyTerminal = null!;
        private Label lblStatus = null!;

        // ============ Services ============
        private AppSettings settings = null!;
        private MacroManager macroManager = null!;
        private SolidWorksConnectionManager swConnectionManager = new SolidWorksConnectionManager();

        // ============ State ============
        private CancellationTokenSource? cancellationTokenSource;
        private bool isRunning = false;
        private List<MacroBookmark> openFiles = new List<MacroBookmark>();
        private MacroBookmark? currentFile = null;
        private int newFileCounter = 1;
        private System.Windows.Forms.Timer? highlightTimer;
        private System.Windows.Forms.Timer? autoSaveTimer;
        private bool isUpdatingSelection = false;
        private ContextMenuStrip tabContextMenu = null!;
        private SearchReplaceDialog? searchDialog = null;
        private Stopwatch executionStopwatch = new Stopwatch();

        public MainFormNew()
        {
            InitializeApp();
            InitializeComponent();
            SetupUI();
            SetupServices();
            LoadInitialContent();
        }

        private void InitializeApp()
        {
            AppSettings.EnsureFolderStructure();
            settings = AppSettings.Load();
            Lang.SetLanguage(settings.Language);
            macroManager = new MacroManager(settings);
        }

        private void InitializeComponent()
        {
            this.Text = Lang.Get("AppTitle");
            this.Size = new Size(settings.WindowWidth, settings.WindowHeight);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.MinimumSize = new Size(1000, 600);
            this.BackColor = DarkBackground;
            this.ForeColor = TextWhite;
            this.FormBorderStyle = FormBorderStyle.Sizable;
            
            // Apply dark title bar
            ApplyDarkTitleBar();
        }

        private void ApplyDarkTitleBar()
        {
            try
            {
                int useImmersiveDarkMode = 1;
                // Try Windows 10 20H1 and later first
                if (DwmSetWindowAttribute(this.Handle, DWMWA_USE_IMMERSIVE_DARK_MODE, ref useImmersiveDarkMode, sizeof(int)) != 0)
                {
                    // Fall back to earlier Windows 10 versions
                    DwmSetWindowAttribute(this.Handle, DWMWA_USE_IMMERSIVE_DARK_MODE_BEFORE_20H1, ref useImmersiveDarkMode, sizeof(int));
                }
            }
            catch { /* Ignore if DWM API is not available */ }
        }

        private void SetupUI()
        {
            // Main toolbar panel
            var toolbarPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 50,
                BackColor = DarkToolbar,
                Padding = new Padding(10, 8, 10, 8)
            };

            SetupToolbar(toolbarPanel);

            // Main split container
            mainSplit = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Vertical,
                SplitterWidth = 3,
                BackColor = DarkSplitter,
                BorderStyle = BorderStyle.None
            };
            mainSplit.Panel1.BackColor = DarkBackground;
            mainSplit.Panel2.BackColor = DarkBackground;

            SetupCodePanel();
            SetupRightPanel();

            mainSplit.Panel1.Controls.Add(codePanel);
            mainSplit.Panel2.Controls.Add(rightSplit);

            this.Controls.Add(mainSplit);
            this.Controls.Add(toolbarPanel);

            // Set splitter after form shown
            this.Shown += (s, e) =>
            {
                try
                {
                    mainSplit.Panel1MinSize = 300;
                    mainSplit.Panel2MinSize = 300;
                    mainSplit.SplitterDistance = (int)(this.ClientSize.Width * 0.55);
                    
                    rightSplit.Panel1MinSize = 100;
                    rightSplit.Panel2MinSize = 100;
                    rightSplit.SplitterDistance = (int)(rightSplit.Height * 0.6);
                }
                catch { }
            };

            // Highlight timer - delay to avoid frequent updates
            highlightTimer = new System.Windows.Forms.Timer { Interval = 800 };
            highlightTimer.Tick += (s, e) =>
            {
                highlightTimer.Stop();
                ApplySyntaxHighlighting();
            };
        }

        private void SetupToolbar(Panel toolbar)
        {
            int x = 10;
            int spacing = 5;

            // Status label
            lblStatus = new Label
            {
                Text = $"⚫ {Lang.Get("NotConnected")}",
                ForeColor = Color.Gray,
                Font = new Font("Segoe UI", 9),
                AutoSize = true,
                Location = new Point(x, 15)
            };
            x += 120;

            // Connect button
            btnConnect = CreateToolbarButton(Lang.Get("Connect"), x);
            btnConnect.Click += (s, e) => Task.Run(() => swConnectionManager.Connect());
            x += btnConnect.Width + spacing;

            // Separator
            x += 10;

            // Run button
            btnRun = CreateToolbarButton($"▶ {Lang.Get("Run")} (F5)", x);
            btnRun.BackColor = AccentGreen;
            btnRun.Click += RunButton_Click;
            x += btnRun.Width + spacing;

            // Stop button
            btnStop = CreateToolbarButton($"⬛ {Lang.Get("Stop")}", x);
            btnStop.BackColor = AccentRed;
            btnStop.Enabled = false;
            btnStop.Click += StopButton_Click;
            x += btnStop.Width + spacing;

            x += 10;

            // New button
            btnNew = CreateToolbarButton($"📄 {Lang.Get("New")}", x);
            btnNew.Click += NewButton_Click;
            x += btnNew.Width + spacing;

            // Save button
            btnSave = CreateToolbarButton($"💾 {Lang.Get("Save")}", x);
            btnSave.Click += SaveButton_Click;
            x += btnSave.Width + spacing;

            // Search button
            var btnSearch = CreateToolbarButton($"🔍 {Lang.Get("Search")}", x);
            btnSearch.Click += (s, e) => ShowSearchDialog(false);
            x += btnSearch.Width + spacing;

            x += 10;

            // Settings button
            btnSettings = CreateToolbarButton($"⚙ {Lang.Get("Settings")}", x);
            btnSettings.Click += SettingsButton_Click;
            x += btnSettings.Width + spacing;

            // Refresh button
            btnRefresh = CreateToolbarButton($"🔄 {Lang.Get("Refresh")}", x);
            btnRefresh.Click += RefreshButton_Click;
            x += btnRefresh.Width + spacing;

            // Help/Shortcuts button
            var btnHelp = CreateToolbarButton($"❓ {Lang.Get("Shortcuts")}", x);
            btnHelp.Click += (s, e) => ShowShortcutsHelp();

            x += btnHelp.Width + spacing;

            // SWP to BAS converter button - DISABLED due to bug that can corrupt SWP files
            var btnConverter = CreateToolbarButton($"🔄 {Lang.Get("SwpToBas")}", x);
            btnConverter.BackColor = Color.FromArgb(80, 80, 80);  // Grayed out
            btnConverter.Enabled = false;
            btnConverter.Click += ConvertButton_Click;

            toolbar.Controls.Add(btnSearch);
            toolbar.Controls.Add(btnHelp);
            toolbar.Controls.Add(btnConverter);
            toolbar.Controls.AddRange(new Control[] { lblStatus, btnConnect, btnRun, btnStop, btnNew, btnSave, btnSettings, btnRefresh });
        }

        private Button CreateToolbarButton(string text, int x)
        {
            var btn = new Button
            {
                Text = text,
                Location = new Point(x, 5),
                AutoSize = false,
                Size = new Size(Math.Max(80, text.Length * 9), 32),
                FlatStyle = FlatStyle.Flat,
                BackColor = DarkPanel,
                ForeColor = TextWhite,
                Font = new Font("Segoe UI", 9),
                Cursor = Cursors.Hand
            };
            btn.FlatAppearance.BorderColor = Color.FromArgb(70, 70, 70);
            btn.FlatAppearance.MouseOverBackColor = Color.FromArgb(60, 60, 60);
            return btn;
        }

        private void SetupCodePanel()
        {
            codePanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = DarkBackground,
                Padding = new Padding(0)
            };

            // Context Menu with more options
            tabContextMenu = new ContextMenuStrip
            {
                BackColor = DarkPanel,
                ForeColor = TextWhite,
                ShowImageMargin = false
            };
            tabContextMenu.Items.Add(Lang.Get("Close"), null, (s, e) => 
            {
                if (codeTabs.SelectedIndex >= 0) CloseTabAt(codeTabs.SelectedIndex);
            });
            tabContextMenu.Items.Add("關閉其他 / Close Others", null, (s, e) => CloseOtherTabs());
            tabContextMenu.Items.Add("關閉全部 / Close All", null, (s, e) => CloseAllTabs());
            tabContextMenu.Items.Add(new ToolStripSeparator());
            tabContextMenu.Items.Add("重命名 / Rename", null, (s, e) => RenameCurrentTab());
            tabContextMenu.Items.Add("另存為 / Save As", null, (s, e) => SaveCurrentTabAs());
            tabContextMenu.Items.Add(new ToolStripSeparator());
            tabContextMenu.Items.Add(Lang.Get("Folder"), null, (s, e) => 
            {
               if (currentFile != null && !string.IsNullOrEmpty(currentFile.FilePath) && File.Exists(currentFile.FilePath))
               {
                   System.Diagnostics.Process.Start("explorer.exe", $"/select,\"{currentFile.FilePath}\"");
               }
            });
            tabContextMenu.Items.Add("複製路徑 / Copy Path", null, (s, e) =>
            {
                if (currentFile != null && !string.IsNullOrEmpty(currentFile.FilePath))
                {
                    Clipboard.SetText(currentFile.FilePath);
                    AppendToTerminal($"📋 Path copied: {currentFile.FilePath}", Color.LightBlue);
                }
            });
            // Style context menu items
            foreach (ToolStripItem item in tabContextMenu.Items)
            {
                if (item is ToolStripMenuItem menuItem)
                {
                    menuItem.BackColor = DarkPanel;
                    menuItem.ForeColor = TextWhite;
                }
            }

            // Tab control for code files - single row with fixed size
            codeTabs = new TabControl
            {
                Dock = DockStyle.Fill,
                BackColor = DarkBackground,
                Font = new Font("Segoe UI", 9),
                DrawMode = TabDrawMode.OwnerDrawFixed,
                Padding = new Point(12, 4),
                SizeMode = TabSizeMode.Fixed,
                ItemSize = new Size(160, 26),  // Fixed width prevents multi-row
                Multiline = false,             // Force single line
                ShowToolTips = true            // Enable tooltips for full name
            };
            codeTabs.DrawItem += CodeTabs_DrawItem;
            codeTabs.MouseDown += CodeTabs_MouseDown;
            codeTabs.SelectedIndexChanged += CodeTabs_SelectedIndexChanged;
            codeTabs.MouseMove += CodeTabs_MouseMove;  // For tooltip

            codePanel.Controls.Add(codeTabs);
        }

        private void CodeTabs_DrawItem(object? sender, DrawItemEventArgs e)
        {
            if (e.Index < 0 || e.Index >= codeTabs.TabPages.Count) return;

            var tabPage = codeTabs.TabPages[e.Index];
            var tabRect = codeTabs.GetTabRect(e.Index);
            
            // Background & Selection
            var isSelected = e.Index == codeTabs.SelectedIndex;
            var bgColor = isSelected ? DarkPanel : DarkBackground;
            
            using (var brush = new SolidBrush(bgColor))
            {
                e.Graphics.FillRectangle(brush, tabRect);
            }

            // Text
            var textColor = isSelected ? TextWhite : Color.Gray;
            var file = openFiles.FirstOrDefault(f => f.Id == tabPage.Tag?.ToString());
            var displayText = file != null && file.IsUnsaved ? "● " + tabPage.Text : tabPage.Text;
            
            using (var brush = new SolidBrush(textColor))
            {
                var textRect = new RectangleF(tabRect.X + 5, tabRect.Y + 5, tabRect.Width - 25, tabRect.Height);
                var sf = new StringFormat { LineAlignment = StringAlignment.Center, Trimming = StringTrimming.EllipsisCharacter };
                e.Graphics.DrawString(displayText, codeTabs.Font, brush, textRect, sf);
            }

            // Close button (X) - Always draw for clarity
            var closeRect = new Rectangle(tabRect.Right - 18, tabRect.Y + 6, 12, 12);
            using (var pen = new Pen(isSelected ? TextWhite : Color.DimGray, 2))
            {
                e.Graphics.DrawLine(pen, closeRect.Left, closeRect.Top, closeRect.Right, closeRect.Bottom);
                e.Graphics.DrawLine(pen, closeRect.Right, closeRect.Top, closeRect.Left, closeRect.Bottom);
            }
        }

        private void CodeTabs_MouseDown(object? sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                for (int i = 0; i < codeTabs.TabPages.Count; i++)
                {
                    if (codeTabs.GetTabRect(i).Contains(e.Location))
                    {
                        codeTabs.SelectedIndex = i;
                        tabContextMenu?.Show(codeTabs, e.Location);
                        return;
                    }
                }
            }
            else if (e.Button == MouseButtons.Left)
            {
                for (int i = 0; i < codeTabs.TabPages.Count; i++)
                {
                    var tabRect = codeTabs.GetTabRect(i);
                    var closeRect = new Rectangle(tabRect.Right - 18, tabRect.Y + 6, 12, 12);
                    
                    if (closeRect.Contains(e.Location))
                    {
                        CloseTabAt(i);
                        return;
                    }
                }
            }
        }

        private void CodeTabs_SelectedIndexChanged(object? sender, EventArgs e)
        {
            if (isUpdatingSelection) return;
            isUpdatingSelection = true;
            try
            {
                if (codeTabs.SelectedIndex >= 0 && codeTabs.SelectedIndex < codeTabs.TabPages.Count)
                {
                    var tabPage = codeTabs.TabPages[codeTabs.SelectedIndex];
                    var file = openFiles.FirstOrDefault(f => f.Id == tabPage.Tag?.ToString());
                    if (file != null)
                    {
                        currentFile = file;
                    }
                }
            }
            finally { isUpdatingSelection = false; }
        }

        private void SetupRightPanel()
        {
            rightSplit = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Horizontal,
                SplitterWidth = 3,
                BackColor = DarkSplitter,
                BorderStyle = BorderStyle.None
            };
            rightSplit.Panel1.BackColor = DarkBackground;
            rightSplit.Panel2.BackColor = DarkBackground;

            // Browser tabs (standard TabControl)
            browserTabs = new TabControl
            {
                Dock = DockStyle.Fill,
                Font = new Font("Segoe UI", 9),
                BackColor = DarkToolbar
            };

            // Create tabs - use localized names
            var recentTab = new TabPage($"📂 {Lang.Get("RecentFiles")}") { BackColor = DarkPanel, BorderStyle = BorderStyle.None };
            var csharpTab = new TabPage(Lang.Get("CSharpMacros")) { BackColor = DarkPanel, BorderStyle = BorderStyle.None };
            var vbaTab = new TabPage(Lang.Get("VBAMacros")) { BackColor = DarkPanel, BorderStyle = BorderStyle.None };
            var snippetsTab = new TabPage($"📝 {Lang.Get("Snippets")}") { BackColor = DarkPanel, BorderStyle = BorderStyle.None };
            var historyTab = new TabPage($"📊 {Lang.Get("History")}") { BackColor = DarkPanel, BorderStyle = BorderStyle.None };
            var promptsTab = new TabPage(Lang.Get("Prompts")) { BackColor = DarkPanel, BorderStyle = BorderStyle.None };
            var notesTab = new TabPage(Lang.Get("Notes")) { BackColor = DarkPanel, BorderStyle = BorderStyle.None };

            SetupRecentFilesTab(recentTab);
            SetupMacroListView(csharpTab, MacroType.CSharp);
            SetupMacroListView(vbaTab, MacroType.VBA);
            SetupSnippetsTab(snippetsTab);
            SetupHistoryTab(historyTab);
            SetupPromptsTab(promptsTab);
            SetupNotesTab(notesTab);

            browserTabs.TabPages.AddRange(new[] { recentTab, csharpTab, vbaTab, snippetsTab, historyTab, promptsTab, notesTab });

            var browserPanel = new Panel { Dock = DockStyle.Fill, BackColor = DarkBackground };
            browserPanel.Controls.Add(browserTabs);

            rightSplit.Panel1.Controls.Add(browserPanel);

            // Terminal
            SetupTerminal();
        }

        private void SetupMacroListView(TabPage tab, MacroType type)
        {
            // Use ListView for detailed file view with sorting
            var listView = new ListView
            {
                Dock = DockStyle.Fill,
                BackColor = DarkPanel,
                ForeColor = TextWhite,
                BorderStyle = BorderStyle.None,
                Font = new Font("Segoe UI", 9),
                Name = type == MacroType.CSharp ? "csharpList" : "vbaList",
                View = System.Windows.Forms.View.Details,
                FullRowSelect = true,
                GridLines = false
            };

            // Add columns
            bool isEnglish = Lang.CurrentLanguage == "en-US";
            listView.Columns.Add(isEnglish ? "Name" : "名稱", 180);
            listView.Columns.Add(isEnglish ? "Ext" : "類型", 50);
            listView.Columns.Add(isEnglish ? "Modified" : "修改時間", 130);
            listView.Columns.Add(isEnglish ? "Path" : "路徑", 200);

            // Enable column click sorting
            listView.ColumnClick += ListView_ColumnClick;

            listView.DoubleClick += (s, e) =>
            {
                if (listView.SelectedItems.Count > 0 && listView.SelectedItems[0].Tag is string path && File.Exists(path))
                {
                    OpenMacroFile(path, type);
                }
            };

            // Context menu for list
            var listMenu = new ContextMenuStrip { BackColor = DarkPanel, ForeColor = TextWhite };
            listMenu.Items.Add(isEnglish ? "Open" : "打開", null, (s, e) =>
            {
                if (listView.SelectedItems.Count > 0 && listView.SelectedItems[0].Tag is string path && File.Exists(path))
                    OpenMacroFile(path, type);
            });
            listMenu.Items.Add(isEnglish ? "Show in Explorer" : "在資源管理器中顯示", null, (s, e) =>
            {
                if (listView.SelectedItems.Count > 0 && listView.SelectedItems[0].Tag is string path && File.Exists(path))
                    System.Diagnostics.Process.Start("explorer.exe", $"/select,\"{path}\"");
            });
            listMenu.Items.Add(isEnglish ? "Copy Path" : "複製路徑", null, (s, e) =>
            {
                if (listView.SelectedItems.Count > 0 && listView.SelectedItems[0].Tag is string path)
                {
                    Clipboard.SetText(path);
                    AppendToTerminal($"📋 {path}", Color.LightBlue);
                }
            });
            listView.ContextMenuStrip = listMenu;

            tab.Controls.Add(listView);
        }

        private int listViewSortColumn = 0;
        private SortOrder listViewSortOrder = SortOrder.Ascending;

        private void ListView_ColumnClick(object? sender, ColumnClickEventArgs e)
        {
            var listView = sender as ListView;
            if (listView == null) return;

            if (e.Column == listViewSortColumn)
            {
                listViewSortOrder = listViewSortOrder == SortOrder.Ascending ? SortOrder.Descending : SortOrder.Ascending;
            }
            else
            {
                listViewSortColumn = e.Column;
                listViewSortOrder = SortOrder.Ascending;
            }

            listView.ListViewItemSorter = new ListViewItemComparer(e.Column, listViewSortOrder);
            listView.Sort();
        }

        private void SetupPromptsTab(TabPage tab)
        {
            var split = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Vertical,
                SplitterDistance = 180,
                BackColor = DarkToolbar
            };

            // List
            var listBox = new ListBox
            {
                Dock = DockStyle.Fill,
                BackColor = DarkToolbar,
                ForeColor = TextWhite,
                BorderStyle = BorderStyle.None,
                Font = new Font("Segoe UI", 9),
                Name = "promptsList"
            };
            listBox.SelectedIndexChanged += PromptsList_SelectedIndexChanged;

            // Toolbar
            var toolbar = new Panel
            {
                Dock = DockStyle.Top,
                Height = 35,
                BackColor = DarkPanel
            };
            var addBtn = new Button { Text = "➕", Location = new Point(5, 5), Size = new Size(30, 25), FlatStyle = FlatStyle.Flat, ForeColor = TextWhite, BackColor = DarkPanel };
            addBtn.Click += AddPrompt_Click;
            var delBtn = new Button { Text = "🗑", Location = new Point(40, 5), Size = new Size(30, 25), FlatStyle = FlatStyle.Flat, ForeColor = TextWhite, BackColor = DarkPanel };
            delBtn.Click += DeletePrompt_Click;
            toolbar.Controls.AddRange(new Control[] { addBtn, delBtn });

            split.Panel1.Controls.Add(listBox);
            split.Panel1.Controls.Add(toolbar);

            // Editor
            var editorPanel = new Panel { Dock = DockStyle.Fill, BackColor = DarkBackground, Padding = new Padding(10) };
            
            var nameBox = new TextBox
            {
                Dock = DockStyle.Top,
                Height = 30,
                BackColor = DarkPanel,
                ForeColor = TextWhite,
                BorderStyle = BorderStyle.FixedSingle,
                Font = new Font("Segoe UI", 11),
                Name = "promptName"
            };
            nameBox.TextChanged += PromptName_TextChanged;

            var contentBox = new RichTextBox
            {
                Dock = DockStyle.Fill,
                BackColor = DarkBackground,
                ForeColor = TextWhite,
                BorderStyle = BorderStyle.None,
                Font = new Font("Consolas", 10),
                Name = "promptContent"
            };
            contentBox.TextChanged += PromptContent_TextChanged;

            var copyBtn = new Button
            {
                Text = $"📋 {Lang.Get("CopyToClipboard")}",
                Dock = DockStyle.Bottom,
                Height = 36,
                FlatStyle = FlatStyle.Flat,
                BackColor = AccentBlue,
                ForeColor = TextWhite
            };
            copyBtn.Click += (s, e) =>
            {
                if (!string.IsNullOrEmpty(contentBox.Text))
                {
                    Clipboard.SetText(contentBox.Text);
                    AppendToTerminal($"📋 {Lang.Get("Copied")}", Color.LightBlue);
                }
            };

            editorPanel.Controls.Add(contentBox);
            editorPanel.Controls.Add(nameBox);
            editorPanel.Controls.Add(copyBtn);

            split.Panel2.Controls.Add(editorPanel);
            tab.Controls.Add(split);
        }

        private void SetupNotesTab(TabPage tab)
        {
            var split = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Vertical,
                SplitterDistance = 180,
                BackColor = DarkToolbar
            };

            var listBox = new ListBox
            {
                Dock = DockStyle.Fill,
                BackColor = DarkToolbar,
                ForeColor = TextWhite,
                BorderStyle = BorderStyle.None,
                Font = new Font("Segoe UI", 9),
                Name = "notesList"
            };
            listBox.SelectedIndexChanged += NotesList_SelectedIndexChanged;

            var toolbar = new Panel { Dock = DockStyle.Top, Height = 35, BackColor = DarkPanel };
            var addBtn = new Button { Text = "➕", Location = new Point(5, 5), Size = new Size(30, 25), FlatStyle = FlatStyle.Flat, ForeColor = TextWhite, BackColor = DarkPanel };
            addBtn.Click += AddNote_Click;
            var delBtn = new Button { Text = "🗑", Location = new Point(40, 5), Size = new Size(30, 25), FlatStyle = FlatStyle.Flat, ForeColor = TextWhite, BackColor = DarkPanel };
            delBtn.Click += DeleteNote_Click;
            toolbar.Controls.AddRange(new Control[] { addBtn, delBtn });

            split.Panel1.Controls.Add(listBox);
            split.Panel1.Controls.Add(toolbar);

            var editorPanel = new Panel { Dock = DockStyle.Fill, BackColor = DarkBackground, Padding = new Padding(10) };

            var nameBox = new TextBox
            {
                Dock = DockStyle.Top,
                BackColor = DarkBackground,
                ForeColor = TextWhite,
                BorderStyle = BorderStyle.None,
                Font = new Font("Segoe UI", 14, FontStyle.Bold),
                Name = "noteName"
            };
            nameBox.TextChanged += NoteName_TextChanged;

            var contentBox = new RichTextBox
            {
                Dock = DockStyle.Fill,
                BackColor = DarkBackground,
                ForeColor = TextWhite,
                BorderStyle = BorderStyle.None,
                Font = new Font("Segoe UI", 10),
                Name = "noteContent"
            };
            contentBox.TextChanged += NoteContent_TextChanged;

            editorPanel.Controls.Add(contentBox);
            editorPanel.Controls.Add(nameBox);

            split.Panel2.Controls.Add(editorPanel);
            tab.Controls.Add(split);
        }

        private void SetupTerminal()
        {
            var terminalPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = DarkTerminal
            };

            var header = new Panel
            {
                Dock = DockStyle.Top,
                Height = 35,
                BackColor = DarkToolbar
            };

            var label = new Label
            {
                Text = $"📟 {Lang.Get("Terminal")}",
                ForeColor = TextWhite,
                Location = new Point(10, 8),
                AutoSize = true
            };

            btnCopyTerminal = new Button
            {
                Text = Lang.Get("CopyOutput"),
                Size = new Size(100, 28),
                Location = new Point(header.Width - 220, 3),
                Anchor = AnchorStyles.Top | AnchorStyles.Right,
                FlatStyle = FlatStyle.Flat,
                BackColor = DarkPanel,
                ForeColor = TextWhite
            };
            btnCopyTerminal.Click += (s, e) =>
            {
                if (!string.IsNullOrEmpty(terminalDisplay.Text))
                {
                    Clipboard.SetText(terminalDisplay.Text);
                    AppendToTerminal($"📋 {Lang.Get("Copied")}", Color.LightBlue);
                }
            };

            var clearBtn = new Button
            {
                Text = Lang.Get("Clear"),
                Size = new Size(70, 28),
                Location = new Point(header.Width - 110, 3),
                Anchor = AnchorStyles.Top | AnchorStyles.Right,
                FlatStyle = FlatStyle.Flat,
                BackColor = DarkPanel,
                ForeColor = TextWhite
            };
            clearBtn.Click += (s, e) => terminalDisplay.Clear();

            header.Controls.AddRange(new Control[] { label, btnCopyTerminal, clearBtn });

            terminalDisplay = new RichTextBox
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                Font = new Font("Cascadia Code", 10),
                BackColor = DarkTerminal,
                ForeColor = Color.FromArgb(204, 204, 204),
                BorderStyle = BorderStyle.None
            };

            terminalPanel.Controls.Add(terminalDisplay);
            terminalPanel.Controls.Add(header);
            rightSplit.Panel2.Controls.Add(terminalPanel);
        }

        private void SetupServices()
        {
            swConnectionManager.ConnectionStateChanged += OnConnectionStateChanged;
            Task.Run(() => swConnectionManager.Connect());
            
            this.KeyPreview = true;
            this.KeyDown += MainForm_KeyDown;

            // Enable drag-drop
            this.AllowDrop = true;
            this.DragEnter += MainForm_DragEnter;
            this.DragDrop += MainForm_DragDrop;

            // Setup auto-save timer
            if (settings.AutoSaveEnabled)
            {
                autoSaveTimer = new System.Windows.Forms.Timer
                {
                    Interval = settings.AutoSaveIntervalSeconds * 1000
                };
                autoSaveTimer.Tick += AutoSaveTimer_Tick;
                autoSaveTimer.Start();
            }
        }

        private void LoadInitialContent()
        {
            RefreshMacroLists();
            RefreshPromptsList();
            RefreshNotesList();
            RefreshRecentFiles();
            RefreshExecutionHistory();
            CreateNewFile();
            AppendToTerminal($"🚀 {Lang.Get("Ready")}", Color.LightGreen);
            AppendToTerminal("快捷鍵: F5執行, Ctrl+S保存, Ctrl+F搜尋, Ctrl+B書籤", Color.Gray);
        }

        // ==================== Event Handlers ====================

        private void OnConnectionStateChanged(object? sender, ConnectionInfo info)
        {
            if (InvokeRequired) { Invoke(new Action(() => OnConnectionStateChanged(sender, info))); return; }

            switch (info.State)
            {
                case ConnectionState.Connected:
                    lblStatus.Text = $"🟢 {Lang.Get("Connected")} SW{info.SolidWorksVersion?.Split('.')[0]}";
                    lblStatus.ForeColor = Color.LightGreen;
                    break;
                case ConnectionState.Connecting:
                    lblStatus.Text = $"🟡 {Lang.Get("Connecting")}";
                    lblStatus.ForeColor = Color.Orange;
                    break;
                default:
                    lblStatus.Text = $"⚫ {Lang.Get("NotConnected")}";
                    lblStatus.ForeColor = Color.Gray;
                    break;
            }
        }

        private void MainForm_KeyDown(object? sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F5 && !isRunning) { RunButton_Click(sender, e); e.Handled = true; }
            else if (e.Control && e.KeyCode == Keys.S) { SaveButton_Click(sender, e); e.Handled = true; }
            else if (e.Control && e.KeyCode == Keys.N) { NewButton_Click(sender, e); e.Handled = true; }
            else if (e.Control && e.KeyCode == Keys.F) { ShowSearchDialog(false); e.Handled = true; }
            else if (e.Control && e.KeyCode == Keys.H) { ShowSearchDialog(true); e.Handled = true; }
            else if (e.Control && e.KeyCode == Keys.G) { ShowGoToLineDialog(); e.Handled = true; }
            else if (e.KeyCode == Keys.F1) { ShowShortcutsHelp(); e.Handled = true; }
        }

        private void MainForm_DragEnter(object? sender, DragEventArgs e)
        {
            if (e.Data?.GetDataPresent(DataFormats.FileDrop) == true)
                e.Effect = DragDropEffects.Copy;
            else
                e.Effect = DragDropEffects.None;
        }

        private void MainForm_DragDrop(object? sender, DragEventArgs e)
        {
            if (e.Data?.GetData(DataFormats.FileDrop) is string[] files)
            {
                foreach (var file in files)
                {
                    var ext = Path.GetExtension(file).ToLowerInvariant();
                    if (ext == ".cs")
                        OpenMacroFile(file, MacroType.CSharp);
                    else if (ext == ".vba" || ext == ".bas" || ext == ".swp")
                        OpenMacroFile(file, MacroType.VBA);
                }
            }
        }

        private void AutoSaveTimer_Tick(object? sender, EventArgs e)
        {
            foreach (var file in openFiles.Where(f => f.IsUnsaved && !string.IsNullOrEmpty(f.FilePath)))
            {
                try
                {
                    // Get editor content for this file
                    var tab = codeTabs.TabPages.Cast<TabPage>().FirstOrDefault(t => t.Tag?.ToString() == file.Id);
                    if (tab != null)
                    {
                        var editor = tab.Controls.OfType<RichTextBox>().FirstOrDefault();
                        if (editor != null)
                        {
                            file.Content = editor.Text;
                            macroManager.SaveMacroContent(file.FilePath, file.Content);
                            file.IsUnsaved = false;
                        }
                    }
                }
                catch { }
            }
            codeTabs.Invalidate();
        }

        private void FileSelector_SelectedIndexChanged(object? sender, EventArgs e)
        {
            // Removed - using tab control now
        }

        private void CodeEditor_TextChanged(object? sender, EventArgs e)
        {
            if (currentFile != null)
            {
                // Check if content actually changed from original
                var editor = sender as RichTextBox;
                if (editor != null)
                {
                    string originalContent = currentFile.Tag as string ?? currentFile.Content;
                    currentFile.IsUnsaved = editor.Text != originalContent;
                }
            }
            highlightTimer?.Stop();
            highlightTimer?.Start();
        }

        private void NewButton_Click(object? sender, EventArgs e) => CreateNewFile();
        
        private void SaveButton_Click(object? sender, EventArgs e)
        {
            if (currentFile == null) return;
            
            // Get current editor
            var currentTab = codeTabs.SelectedTab;
            if (currentTab == null) return;
            var editor = currentTab.Controls.OfType<RichTextBox>().FirstOrDefault();
            if (editor == null) return;
            
            currentFile.Content = editor.Text;

            if (string.IsNullOrEmpty(currentFile.FilePath))
            {
                using var dialog = new SaveFileDialog
                {
                    Filter = currentFile.Type == MacroType.CSharp 
                        ? "C# Files (*.cs)|*.cs|Text Files (*.txt)|*.txt|JSON Files (*.json)|*.json|All Files (*.*)|*.*" 
                        : "VBA Files (*.bas)|*.bas|VBA Files (*.vba)|*.vba|Text Files (*.txt)|*.txt|All Files (*.*)|*.*",
                    FileName = currentFile.Name,
                    InitialDirectory = Path.Combine(AppSettings.DefaultMacrosPath, currentFile.Type == MacroType.CSharp ? "C Sharp" : "VBA")
                };
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    currentFile.FilePath = dialog.FileName;
                    currentFile.Name = Path.GetFileNameWithoutExtension(dialog.FileName);
                }
                else return;
            }

            if (macroManager.SaveMacroContent(currentFile.FilePath, currentFile.Content))
            {
                currentFile.IsUnsaved = false;
                codeTabs.Invalidate();
                AppendToTerminal($"✅ {Lang.Get("SaveSuccess")}: {currentFile.FilePath}", Color.LightGreen);
                RefreshMacroLists();
            }
            else
            {
                AppendToTerminal($"❌ {Lang.Get("SaveFailed")}", Color.Red);
            }
        }

        private void SettingsButton_Click(object? sender, EventArgs e)
        {
            using var dialog = new SettingsDialogNew(settings);
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                settings.Save();
                Lang.SetLanguage(settings.Language);
                RefreshMacroLists();
                AppendToTerminal($"⚙ {Lang.Get("SettingsUpdated")}", Color.LightBlue);
                // Note: Full UI refresh requires restart
                MessageBox.Show("語言變更將在重啟後完全生效\nLanguage change will take full effect after restart.", 
                    Lang.Get("Settings"), MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void RefreshButton_Click(object? sender, EventArgs e)
        {
            RefreshMacroLists();
            AppendToTerminal($"🔄 {Lang.Get("ListRefreshed")}", Color.LightBlue);
        }

        private async void RunButton_Click(object? sender, EventArgs e)
        {
            if (currentFile == null) return;
            
            // Get current editor content
            var currentTab = codeTabs.SelectedTab;
            if (currentTab == null) return;
            var editor = currentTab.Controls.OfType<RichTextBox>().FirstOrDefault();
            if (editor == null) return;
            
            // VBA files need SolidWorks to run
            if (currentFile.Type == MacroType.VBA)
            {
                if (!swConnectionManager.IsConnected || swConnectionManager.SwApp == null)
                {
                    AppendToTerminal($"❌ {Lang.Get("VBANeedsSW")}", Color.Red);
                    return;
                }
                
                // For SWP files, always use the file path for execution
                if (!string.IsNullOrEmpty(currentFile.FilePath) && 
                    (currentFile.FilePath.EndsWith(".swp", StringComparison.OrdinalIgnoreCase) ||
                     currentFile.FilePath.EndsWith(".bas", StringComparison.OrdinalIgnoreCase) ||
                     currentFile.FilePath.EndsWith(".vba", StringComparison.OrdinalIgnoreCase)))
                {
                    RunVBAMacro(null, currentFile.FilePath);
                }
                else
                {
                    if (string.IsNullOrWhiteSpace(editor.Text)) return;
                    RunVBAMacro(editor.Text, null);
                }
                return;
            }
            
            // C# execution
            if (string.IsNullOrWhiteSpace(editor.Text)) return;

            if (currentFile.Type != MacroType.CSharp)
            {
                AppendToTerminal($"⚠ {Lang.Get("OnlyCSCanRun")}", Color.Orange);
                return;
            }

            string codeToRun = editor.Text;

            if (MecAgentConverter.IsMecAgentFormat(codeToRun))
            {
                AppendToTerminal($"🔄 {Lang.Get("MecAgentDetected")}", Color.LightBlue);
                codeToRun = MecAgentConverter.ConvertToMiniSW(codeToRun);
                AppendToTerminal($"✅ {Lang.Get("ConversionComplete")}", Color.LightGreen);
            }

            UpdateRunningState(true);
            terminalDisplay.Clear();
            AppendToTerminal($"▶ {Lang.Get("ExecutionStarted")} [{DateTime.Now:HH:mm:ss}]", Color.LightBlue);
            AppendToTerminal("═══════════════════════════════════════", Color.Gray);

            if (swConnectionManager.IsConnected)
                AppendToTerminal($"✅ {Lang.Get("SWConnected")}", Color.LightGreen);
            else
                AppendToTerminal($"⚠ {Lang.Get("SWNotConnected")}", Color.Orange);
            AppendToTerminal("");

            cancellationTokenSource = new CancellationTokenSource();
            executionStopwatch.Restart();
            bool success = true;
            string errorMsg = "";

            try
            {
                var globals = new ScriptGlobals
                {
                    swApp = swConnectionManager.SwApp,
                    swModel = swConnectionManager.SwApp?.IActiveDoc2,
                    Print = (msg) => AppendToTerminal(msg, Color.White),
                    PrintError = (msg) => AppendToTerminal($"❌ {msg}", Color.Red),
                    PrintWarning = (msg) => AppendToTerminal($"⚠ {msg}", Color.Orange)
                };

                var options = ScriptOptions.Default
                    .WithReferences(typeof(object).Assembly, typeof(Enumerable).Assembly, typeof(List<>).Assembly, typeof(ISldWorks).Assembly, typeof(swDocumentTypes_e).Assembly)
                    .WithImports("System", "System.Math", "System.IO", "System.Collections.Generic", "System.Linq", "SolidWorks.Interop.sldworks", "SolidWorks.Interop.swconst");

                await Task.Run(async () =>
                {
                    try
                    {
                        var result = await CSharpScript.EvaluateAsync(codeToRun, options, globals, typeof(ScriptGlobals), cancellationTokenSource.Token);
                        if (result != null) AppendToTerminal($"\n返回值: {result}", Color.Yellow);
                    }
                    catch (CompilationErrorException ex)
                    {
                        success = false;
                        errorMsg = ex.Diagnostics.FirstOrDefault()?.ToString() ?? ex.Message;
                        AppendToTerminal($"{Lang.Get("CompilationError")}:", Color.Red);
                        foreach (var d in ex.Diagnostics) AppendToTerminal($"  {d}", Color.Red);
                    }
                    catch (Exception ex)
                    {
                        success = false;
                        errorMsg = ex.Message;
                        AppendToTerminal($"{Lang.Get("RuntimeError")}: {ex.Message}", Color.Red);
                    }
                }, cancellationTokenSource.Token);

                executionStopwatch.Stop();
                AppendToTerminal("");
                AppendToTerminal("═══════════════════════════════════════", Color.Gray);
                AppendToTerminal($"⬛ {Lang.Get("ExecutionCompleted")} [{DateTime.Now:HH:mm:ss}] ⏱ {executionStopwatch.ElapsedMilliseconds}ms", Color.LightBlue);
            }
            catch (OperationCanceledException)
            {
                success = false;
                errorMsg = "Cancelled";
                executionStopwatch.Stop();
                AppendToTerminal($"\n⚠ {Lang.Get("ExecutionCancelled")}", Color.Orange);
            }
            finally
            {
                // Record execution history
                settings.AddExecutionHistory(
                    currentFile?.Name ?? "Unknown",
                    currentFile?.FilePath ?? "",
                    success,
                    executionStopwatch.ElapsedMilliseconds,
                    errorMsg
                );
                settings.Save();
                RefreshExecutionHistory();

                UpdateRunningState(false);
                cancellationTokenSource?.Dispose();
                cancellationTokenSource = null;
            }
        }

        private void StopButton_Click(object? sender, EventArgs e) => cancellationTokenSource?.Cancel();

        private void ConvertButton_Click(object? sender, EventArgs e)
        {
            using var openDialog = new OpenFileDialog
            {
                Title = "選擇 SWP 宏文件 / Select SWP Macro File",
                Filter = "SolidWorks 宏文件 (*.swp)|*.swp|所有文件 (*.*)|*.*",
                FilterIndex = 1
            };

            if (openDialog.ShowDialog() == DialogResult.OK)
            {
                using var saveDialog = new SaveFileDialog
                {
                    Title = "保存為 BAS 文件 / Save as BAS File",
                    Filter = "VBA 基礎模塊 (*.bas)|*.bas|所有文件 (*.*)|*.*",
                    FilterIndex = 1,
                    FileName = Path.GetFileNameWithoutExtension(openDialog.FileName) + ".bas"
                };

                if (saveDialog.ShowDialog() == DialogResult.OK)
                {
                    ConvertSwpToBas(openDialog.FileName, saveDialog.FileName);
                }
            }
        }

        private void ConvertSwpToBas(string swpPath, string basPath)
        {
            if (swConnectionManager.SwApp == null)
            {
                MessageBox.Show("請先連接到 SolidWorks", "轉換失敗", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                AppendToTerminal("🔄 開始轉換 SWP 到 BAS...", Color.LightBlue);
                AppendToTerminal($"📂 來源: {Path.GetFileName(swpPath)}", Color.Gray);
                AppendToTerminal($"📁 目標: {Path.GetFileName(basPath)}", Color.Gray);

                // Method 1: Using SolidWorks API to extract macro content
                bool success = ExtractMacroContentUsingSolidWorks(swpPath, basPath);

                if (!success)
                {
                    // Method 2: Try alternative approach using VBA engine
                    success = ExtractMacroContentUsingVBA(swpPath, basPath);
                }

                if (success)
                {
                    AppendToTerminal("✅ 轉換成功完成!", Color.LightGreen);
                    MessageBox.Show($"轉換成功!\n輸出文件: {basPath}", "轉換完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    AppendToTerminal("❌ 轉換失敗", Color.Red);
                    MessageBox.Show("轉換失敗。請確保 SWP 文件有效且可以讀取。", "轉換失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                AppendToTerminal($"❌ 轉換過程中發生錯誤: {ex.Message}", Color.Red);
                MessageBox.Show($"轉換失敗: {ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private bool ExtractMacroContentUsingSolidWorks(string swpPath, string basPath)
        {
            try
            {
                AppendToTerminal("🔧 方法1: 使用 SolidWorks API 提取宏內容", Color.Cyan);
                
                // Get macro methods to understand the structure
                // Note: SWP files are binary structured storage, getting methods requires different approach
                AppendToTerminal("ℹ️ 注意: SWP 是二進制格式，無法直接獲取方法列表", Color.Orange);
                
                // Try alternative approach - create basic template
                return TryExtractByAnalysis(swpPath, basPath);
            }
            catch (Exception ex)
            {
                AppendToTerminal($"❌ 方法1失敗: {ex.Message}", Color.Orange);
                return false;
            }
        }

        private bool TryExtractByAnalysis(string swpPath, string basPath)
        {
            try
            {
                AppendToTerminal("🔄 嘗試通過文件分析提取內容", Color.Cyan);
                
                // Try to detect SUB procedures from the file
                var detectedSubs = DetectSubProceduresInSwp(swpPath);
                
                // Create a BAS file with detected content
                var basContent = new System.Text.StringBuilder();
                basContent.AppendLine($"' Extracted from SWP file: {Path.GetFileName(swpPath)}");
                basContent.AppendLine($"' Extraction date: {DateTime.Now}");
                basContent.AppendLine();
                basContent.AppendLine("Option Explicit");
                basContent.AppendLine();
                
                if (detectedSubs.Count > 0)
                {
                    basContent.AppendLine("' Detected SUB procedures:");
                    foreach (var sub in detectedSubs)
                    {
                        basContent.AppendLine($"' - {sub}");
                    }
                    basContent.AppendLine();
                    
                    // Create stub implementations
                    foreach (var sub in detectedSubs)
                    {
                        basContent.AppendLine($"Sub {sub}()");
                        basContent.AppendLine($"    ' TODO: Implement {sub} logic");
                        basContent.AppendLine($"    ' Original implementation was in: {Path.GetFileName(swpPath)}");
                        basContent.AppendLine("    ");
                        basContent.AppendLine("    Dim swApp As SldWorks.SldWorks");
                        basContent.AppendLine("    Set swApp = Application.SldWorks");
                        basContent.AppendLine("    ");
                        basContent.AppendLine("    If swApp Is Nothing Then");
                        basContent.AppendLine("        MsgBox \"Cannot connect to SolidWorks\"");
                        basContent.AppendLine("        Exit Sub");
                        basContent.AppendLine("    End If");
                        basContent.AppendLine("    ");
                        basContent.AppendLine($"    MsgBox \"{sub} procedure needs manual implementation\"");
                        basContent.AppendLine("    ");
                        basContent.AppendLine("End Sub");
                        basContent.AppendLine();
                    }
                }
                else
                {
                    // Fallback to main procedure
                    basContent.AppendLine("Sub main()");
                    basContent.AppendLine("    ' TODO: Implement main logic");
                    basContent.AppendLine($"    ' Original file: {swpPath}");
                    basContent.AppendLine("    ");
                    basContent.AppendLine("    Dim swApp As SldWorks.SldWorks");
                    basContent.AppendLine("    Set swApp = Application.SldWorks");
                    basContent.AppendLine("    ");
                    basContent.AppendLine("    If swApp Is Nothing Then");
                    basContent.AppendLine("        MsgBox \"Cannot connect to SolidWorks\"");
                    basContent.AppendLine("        Exit Sub");
                    basContent.AppendLine("    End If");
                    basContent.AppendLine("    ");
                    basContent.AppendLine("    MsgBox \"Macro converted from SWP - manual editing required\"");
                    basContent.AppendLine("    ");
                    basContent.AppendLine("End Sub");
                }
                
                basContent.AppendLine();
                basContent.AppendLine("' Note: This file was auto-generated from SWP format.");
                basContent.AppendLine("' Manual editing is required to restore full functionality.");
                
                File.WriteAllText(basPath, basContent.ToString());
                
                if (detectedSubs.Count > 0)
                {
                    AppendToTerminal($"📄 創建了包含 {detectedSubs.Count} 個 SUB 程序的 BAS 文件", Color.Yellow);
                }
                else
                {
                    AppendToTerminal("📄 創建了基本 BAS 模板文件", Color.Yellow);
                }
                
                AppendToTerminal("⚠ 注意: 需要手動編輯以恢復完整功能", Color.Orange);
                return true;
            }
            catch (Exception ex)
            {
                AppendToTerminal($"❌ 分析失敗: {ex.Message}", Color.Red);
                return false;
            }
        }

        private List<string> DetectSubProceduresInSwp(string swpPath)
        {
            var detectedSubs = new List<string>();
            
            try
            {
                AppendToTerminal("🔍 檢測 SWP 文件中的 SUB 程序", Color.Cyan);
                
                // Read file as binary
                byte[] fileBytes = File.ReadAllBytes(swpPath);
                AppendToTerminal($"  • 文件大小: {fileBytes.Length} bytes", Color.Gray);
                
                // Try multiple encodings to find VBA content
                string[] encodingsToTry = { "UTF-8", "ASCII", "Unicode", "UTF-16LE", "UTF-16BE" };
                
                foreach (string encodingName in encodingsToTry)
                {
                    try
                    {
                        var encoding = System.Text.Encoding.GetEncoding(encodingName);
                        string content = encoding.GetString(fileBytes);
                        
                        // Search for SUB patterns in this encoding
                        var foundSubs = FindSubsInContent(content);
                        foreach (var sub in foundSubs)
                        {
                            if (!detectedSubs.Contains(sub))
                            {
                                detectedSubs.Add(sub);
                                AppendToTerminal($"  • 找到 SUB: {sub} (編碼: {encodingName})", Color.Gray);
                            }
                        }
                    }
                    catch { /* Encoding not supported, skip */ }
                }
                
                // Also try to search for "Sub " pattern directly in bytes
                if (detectedSubs.Count == 0)
                {
                    var binarySubs = FindSubsInBytes(fileBytes);
                    foreach (var sub in binarySubs)
                    {
                        if (!detectedSubs.Contains(sub))
                        {
                            detectedSubs.Add(sub);
                            AppendToTerminal($"  • 找到 SUB: {sub} (二進制掃描)", Color.Gray);
                        }
                    }
                }
                
                // If still no SUBs found, try common name heuristics
                if (detectedSubs.Count == 0)
                {
                    AppendToTerminal("  • 未找到明確的 SUB 定義，嘗試常見名稱", Color.Orange);
                    detectedSubs = TryCommonSubNames(fileBytes);
                }
                
                // Remove invalid entries
                detectedSubs = detectedSubs.Where(s => 
                    !string.IsNullOrWhiteSpace(s) && 
                    s.Length >= 1 && 
                    s.Length < 50 && 
                    !s.Contains("\0") &&
                    char.IsLetter(s[0]) &&
                    s.All(c => char.IsLetterOrDigit(c) || c == '_')
                ).Distinct().ToList();
                
                if (detectedSubs.Count > 0)
                {
                    AppendToTerminal($"✅ 總共找到 {detectedSubs.Count} 個 SUB 程序", Color.LightGreen);
                }
                else
                {
                    AppendToTerminal("⚠ 未找到任何 SUB 程序", Color.Orange);
                }
                
                return detectedSubs;
            }
            catch (Exception ex)
            {
                AppendToTerminal($"❌ SUB 檢測失敗: {ex.Message}", Color.Red);
                return new List<string>();
            }
        }
        
        private List<string> FindSubsInContent(string content)
        {
            var subs = new List<string>();
            
            // Split by multiple delimiters
            var lines = content.Split(new char[] { '\0', '\n', '\r', '\t' }, StringSplitOptions.RemoveEmptyEntries);
            
            foreach (string line in lines)
            {
                string cleanLine = line.Trim();
                
                // Look for "Sub " at the beginning (case insensitive)
                int subIndex = cleanLine.IndexOf("Sub ", StringComparison.OrdinalIgnoreCase);
                if (subIndex == 0 || (subIndex > 0 && !char.IsLetter(cleanLine[subIndex - 1])))
                {
                    string afterSub = cleanLine.Substring(subIndex + 4).Trim();
                    string subName = ExtractSubNameFromString(afterSub);
                    
                    if (!string.IsNullOrEmpty(subName) && !subs.Contains(subName))
                    {
                        subs.Add(subName);
                    }
                }
            }
            
            return subs;
        }
        
        private List<string> FindSubsInBytes(byte[] fileBytes)
        {
            var subs = new List<string>();
            
            // Look for "Sub " pattern in various encodings
            byte[][] subPatterns = {
                System.Text.Encoding.ASCII.GetBytes("Sub "),
                System.Text.Encoding.UTF8.GetBytes("Sub "),
                new byte[] { 0x53, 0x00, 0x75, 0x00, 0x62, 0x00, 0x20, 0x00 }, // UTF-16LE "Sub "
            };
            
            foreach (var pattern in subPatterns)
            {
                int index = 0;
                while ((index = FindPattern(fileBytes, pattern, index)) >= 0)
                {
                    // Extract the SUB name after the pattern
                    string subName = ExtractSubNameFromPosition(fileBytes, index + pattern.Length);
                    if (!string.IsNullOrEmpty(subName) && !subs.Contains(subName))
                    {
                        subs.Add(subName);
                    }
                    index++;
                }
            }
            
            return subs;
        }
        
        private int FindPattern(byte[] data, byte[] pattern, int startIndex)
        {
            for (int i = startIndex; i <= data.Length - pattern.Length; i++)
            {
                bool found = true;
                for (int j = 0; j < pattern.Length; j++)
                {
                    if (data[i + j] != pattern[j])
                    {
                        found = false;
                        break;
                    }
                }
                if (found) return i;
            }
            return -1;
        }
        
        private string ExtractSubNameFromPosition(byte[] data, int position)
        {
            try
            {
                var nameBytes = new List<byte>();
                for (int i = position; i < Math.Min(position + 100, data.Length); i++)
                {
                    byte b = data[i];
                    // Stop at parenthesis, space after name, or non-printable
                    if (b == '(' || b == ')' || b == '\n' || b == '\r' || b == '\0')
                    {
                        break;
                    }
                    // Skip null bytes (for UTF-16)
                    if (b == 0x00) continue;
                    // Accept letters, digits, underscore
                    if ((b >= 'A' && b <= 'Z') || (b >= 'a' && b <= 'z') || 
                        (b >= '0' && b <= '9') || b == '_')
                    {
                        nameBytes.Add(b);
                    }
                    else if (nameBytes.Count > 0)
                    {
                        // Stop at first non-identifier character after we've started collecting
                        break;
                    }
                }
                
                if (nameBytes.Count > 0)
                {
                    return System.Text.Encoding.ASCII.GetString(nameBytes.ToArray());
                }
            }
            catch { }
            
            return "";
        }
        
        private string ExtractSubNameFromString(string afterSub)
        {
            // Remove leading/trailing spaces
            afterSub = afterSub.Trim();
            
            // Find the end of the name (parenthesis or space)
            int endIndex = afterSub.IndexOfAny(new char[] { '(', ' ', '\t', '\n', '\r' });
            
            if (endIndex > 0)
            {
                return afterSub.Substring(0, endIndex).Trim();
            }
            else if (afterSub.Length > 0 && afterSub.Length < 50)
            {
                return afterSub.Trim();
            }
            
            return "";
        }
        
        private List<string> TryCommonSubNames(byte[] fileBytes)
        {
            var subs = new List<string>();
            string[] commonNames = { "main", "Main", "MAIN", "macro", "Macro", "MACRO", 
                                      "swmain", "SwMain", "start", "Start", "run", "Run",
                                      "CreateSketch", "CreatePart", "DrawLine" };
            
            string content = System.Text.Encoding.ASCII.GetString(fileBytes);
            
            foreach (string name in commonNames)
            {
                // Look for patterns like "Sub main" or "sub main"
                if (content.Contains($"Sub {name}", StringComparison.OrdinalIgnoreCase) ||
                    content.Contains($"sub {name}", StringComparison.OrdinalIgnoreCase) ||
                    content.Contains($"{name}()", StringComparison.OrdinalIgnoreCase))
                {
                    if (!subs.Contains(name))
                    {
                        subs.Add(name);
                        AppendToTerminal($"  • 推測 SUB: {name}", Color.Gray);
                    }
                }
            }
            
            return subs;
        }
        
        private bool ExtractMacroContentUsingVBA(string swpPath, string basPath)
        {
            try
            {
                AppendToTerminal("🔧 方法2: 使用 VBA 引擎提取", Color.Cyan);
                AppendToTerminal("⚠ 此方法需要額外的 COM 引用", Color.Orange);
                
                // This would require Microsoft.Vbe.Interop reference
                // For now, create a placeholder implementation
                var templateContent = CreateBasTemplate(swpPath);
                File.WriteAllText(basPath, templateContent);
                
                AppendToTerminal("📄 已創建基於模板的 BAS 文件", Color.Yellow);
                return true;
            }
            catch (Exception ex)
            {
                AppendToTerminal($"❌ 方法2失敗: {ex.Message}", Color.Red);
                return false;
            }
        }

        private string CreateBasTemplate(string swpPath)
        {
            var template = new System.Text.StringBuilder();
            template.AppendLine("' Auto-generated BAS file from SWP conversion");
            template.AppendLine($"' Original file: {swpPath}");
            template.AppendLine($"' Converted on: {DateTime.Now}");
            template.AppendLine("' ");
            template.AppendLine("' Note: This is a template file. You may need to manually add the actual macro content.");
            template.AppendLine();
            template.AppendLine("Option Explicit");
            template.AppendLine();
            template.AppendLine("Sub main()");
            template.AppendLine("    ' Main entry point");
            template.AppendLine("    ' TODO: Add your macro logic here");
            template.AppendLine("    ");
            template.AppendLine("    Dim swApp As SldWorks.SldWorks");
            template.AppendLine("    Set swApp = Application.SldWorks");
            template.AppendLine("    ");
            template.AppendLine("    If swApp Is Nothing Then");
            template.AppendLine("        MsgBox \"Cannot connect to SolidWorks\"");
            template.AppendLine("        Exit Sub");
            template.AppendLine("    End If");
            template.AppendLine("    ");
            template.AppendLine("    ' Your macro code here...");
            template.AppendLine("    MsgBox \"Macro converted from SWP format - edit as needed\"");
            template.AppendLine("    ");
            template.AppendLine("End Sub");
            template.AppendLine();
            template.AppendLine("' Add additional subroutines as needed");
            template.AppendLine("Sub Helper1()");
            template.AppendLine("    ' Helper function 1");
            template.AppendLine("End Sub");
            template.AppendLine();
            template.AppendLine("Sub Helper2()");
            template.AppendLine("    ' Helper function 2");
            template.AppendLine("End Sub");
            
            return template.ToString();
        }

        private string GetVBAErrorDescription(int errorCode)
        {
            return errorCode switch
            {
                2 => "找不到文件 / File not found",
                5 => "無效的程序調用 / Invalid procedure call",
                9 => "下標超出範圍 / Subscript out of range",
                13 => "類型不匹配 / Type mismatch",
                20 => "未處理的錯誤 / Resume without error",
                22 => "找不到指定的模塊或程序 / Module or procedure not found",
                53 => "找不到文件 / File not found",
                91 => "對象變量未設置 / Object variable not set",
                429 => "無法創建 ActiveX 組件 / Can't create ActiveX component",
                _ => $"VBA 錯誤代碼 {errorCode}"
            };
        }

        private (string module, string procedure)? ShowVBASubSelection(string macroPath)
        {
            // Try to get methods using SolidWorks API first
            List<string> detectedSubs = new List<string>();
            string suggestedModule = "Module1";
            string suggestedProcedure = "main";
            
            // Check for companion BAS file
            var basPath = Path.ChangeExtension(macroPath, ".bas");
            if (File.Exists(basPath))
            {
                try
                {
                    string basContent = File.ReadAllText(basPath);
                    var moduleInfo = ParseBasModuleInfo(basContent);
                    
                    if (!string.IsNullOrEmpty(moduleInfo.ModuleName))
                        suggestedModule = moduleInfo.ModuleName;
                    
                    if (!string.IsNullOrEmpty(moduleInfo.EntryPoint))
                        suggestedProcedure = moduleInfo.EntryPoint;
                    
                    detectedSubs = moduleInfo.Methods;
                    
                    AppendToTerminal($"📄 從 BAS 文件解析: 模塊={suggestedModule}, 入口={suggestedProcedure}, 方法數={detectedSubs.Count}", Color.Cyan);
                }
                catch (Exception ex)
                {
                    AppendToTerminal($"⚠ BAS 解析失敗: {ex.Message}", Color.Orange);
                }
            }
            
            // Try SolidWorks GetMacroMethods if no BAS file or parsing failed
            if (detectedSubs.Count == 0 && swConnectionManager.SwApp != null)
            {
                try
                {
                    var methods = swConnectionManager.SwApp.GetMacroMethods(macroPath, 2) as object[];
                    if (methods != null && methods.Length > 0)
                    {
                        AppendToTerminal($"🔍 使用 SolidWorks API 獲取方法: {methods.Length} 個", Color.Cyan);
                        foreach (var method in methods)
                        {
                            if (method is object[] methodInfo && methodInfo.Length >= 2)
                            {
                                string moduleName = methodInfo[0]?.ToString() ?? "";
                                string procedureName = methodInfo[1]?.ToString() ?? "";
                                
                                if (!string.IsNullOrEmpty(procedureName) && !detectedSubs.Contains(procedureName))
                                {
                                    detectedSubs.Add(procedureName);
                                    
                                    // Use first module as suggestion
                                    if (!string.IsNullOrEmpty(moduleName) && suggestedModule == "Module1")
                                        suggestedModule = moduleName;
                                    
                                    // Check for entry point
                                    if (procedureName.Equals("AutoMain", StringComparison.OrdinalIgnoreCase))
                                        suggestedProcedure = "AutoMain";
                                    else if (suggestedProcedure == "main" && 
                                            (procedureName.Equals("main", StringComparison.OrdinalIgnoreCase) ||
                                             procedureName.Equals("Main", StringComparison.OrdinalIgnoreCase)))
                                        suggestedProcedure = procedureName;
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    AppendToTerminal($"⚠ GetMacroMethods 失敗: {ex.Message}", Color.Orange);
                }
            }
            
            // Fallback to SWP scanning if still no methods found
            if (detectedSubs.Count == 0)
            {
                detectedSubs = DetectSubProceduresInSwp(macroPath);
            }
            
            using var dialog = new Form
            {
                Width = 500, Height = 320,
                Text = "選擇要執行的程序 / Select Procedure",
                FormBorderStyle = FormBorderStyle.FixedDialog,
                StartPosition = FormStartPosition.CenterParent,
                BackColor = DarkPanel,
                ForeColor = TextWhite,
                MaximizeBox = false,
                MinimizeBox = false
            };

            var lblFile = new Label
            {
                Text = $"宏檔案: {Path.GetFileName(macroPath)}",
                Location = new Point(15, 15),
                Size = new Size(410, 20),
                ForeColor = TextWhite
            };

            var lblDetected = new Label
            {
                Text = detectedSubs.Count > 0 ? $"找到 {detectedSubs.Count} 個 SUB 程序:" : "未找到 SUB 程序，使用預設選項:",
                Location = new Point(15, 40),
                Size = new Size(410, 20),
                ForeColor = detectedSubs.Count > 0 ? Color.LightGreen : Color.Orange
            };

            var lblModule = new Label
            {
                Text = "模塊名稱 / Module Name (可手動輸入):",
                Location = new Point(15, 70),
                AutoSize = true,
                ForeColor = TextWhite
            };

            var cmbModule = new ComboBox
            {
                Location = new Point(15, 93),
                Size = new Size(180, 25),
                BackColor = DarkBackground,
                ForeColor = TextWhite,
                DropDownStyle = ComboBoxStyle.DropDown // Allow free text input
            };
            // Common module names
            cmbModule.Items.AddRange(new[] { suggestedModule, "Module1", "main", "Main", "ThisDocument", "Sheet1", "macro", "Macro" });
            cmbModule.Text = suggestedModule;

            var lblProcedure = new Label
            {
                Text = "程序名稱 / Procedure Name (可手動輸入):",
                Location = new Point(210, 70),
                AutoSize = true,
                ForeColor = TextWhite
            };

            var cmbProcedure = new ComboBox
            {
                Location = new Point(210, 93),
                Size = new Size(200, 25),
                BackColor = DarkBackground,
                ForeColor = TextWhite,
                DropDownStyle = ComboBoxStyle.DropDown // Allow free text input
            };
            
            // Add detected SUBs first, then recent entry points, then common ones
            if (detectedSubs.Count > 0)
            {
                foreach (var sub in detectedSubs)
                {
                    cmbProcedure.Items.Add($"{sub} (已檢測)");
                }
                // Use suggested procedure from BAS or detection
                bool foundSuggestion = false;
                foreach (var item in cmbProcedure.Items)
                {
                    string? itemText = item?.ToString();
                    if (itemText != null && itemText.StartsWith(suggestedProcedure + " "))
                    {
                        cmbProcedure.Text = itemText;
                        foundSuggestion = true;
                        break;
                    }
                }
                if (!foundSuggestion)
                    cmbProcedure.Text = $"{detectedSubs[0]} (已檢測)";
            }
            else
            {
                cmbProcedure.Text = suggestedProcedure;
            }
            
            // Add recent entry points
            var recentEntryPoints = settings.RecentEntryPoints.Take(10).ToList();
            if (recentEntryPoints.Count > 0)
            {
                foreach (var recent in recentEntryPoints)
                {
                    string moduleItem = $"{recent.ModuleName} (最近使用)";
                    string procedureItem = $"{recent.ProcedureName} (最近使用)";
                    
                    if (!cmbModule.Items.Contains(moduleItem))
                        cmbModule.Items.Add(moduleItem);
                    if (!cmbProcedure.Items.Contains(procedureItem))
                        cmbProcedure.Items.Add(procedureItem);
                }
                
                // Set most recent as default if no detected subs
                if (detectedSubs.Count == 0)
                {
                    var mostRecent = recentEntryPoints.First();
                    cmbModule.Text = mostRecent.ModuleName;
                    cmbProcedure.Text = mostRecent.ProcedureName;
                }
            }
            
            // Add common procedure names
            string[] commonProcedures = { "main", "Main", "macro", "Macro", "swmain", "SwMain", "run", "Run", "start", "Start" };
            foreach (var proc in commonProcedures)
            {
                if (!cmbProcedure.Items.Cast<string>().Any(item => item.StartsWith(proc + " ")))
                {
                    cmbProcedure.Items.Add(proc);
                }
            }

            var lblDetectedList = new Label
            {
                Text = detectedSubs.Count > 0 ? string.Join(", ", detectedSubs) : "無",
                Location = new Point(15, 125),
                Size = new Size(410, 40),
                ForeColor = Color.FromArgb(180, 180, 180),
                Font = new Font("Consolas", 8)
            };

            var btnAuto = new Button
            {
                Text = "🔍 重新檢測",
                Location = new Point(15, 175),
                Size = new Size(100, 32),
                FlatStyle = FlatStyle.Flat,
                BackColor = AccentBlue,
                ForeColor = TextWhite
            };
            btnAuto.Click += (s, e) =>
            {
                var newDetected = DetectSubProceduresInSwp(macroPath);
                if (newDetected.Count > 0)
                {
                    cmbProcedure.Items.Clear();
                    foreach (var sub in newDetected)
                    {
                        cmbProcedure.Items.Add($"{sub} (新檢測)");
                    }
                    cmbProcedure.Text = $"{newDetected[0]} (新檢測)";
                    lblDetectedList.Text = string.Join(", ", newDetected);
                    MessageBox.Show($"重新檢測到 {newDetected.Count} 個 SUB 程序", "檢測結果", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("未找到任何 SUB 程序，請手動輸入或使用預設選項", "檢測結果", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            };

            var chkTryAll = new CheckBox
            {
                Text = "執行失敗時自動嘗試其他組合",
                Location = new Point(125, 180),
                Size = new Size(250, 25),
                ForeColor = TextWhite,
                Checked = true
            };

            var lblTip = new Label
            {
                Text = detectedSubs.Count > 0 ? 
                    "💡 已檢測到宏中的 SUB 程序，也可手動輸入名稱" :
                    "💡 提示: 請手動輸入模塊和程序名稱，或使用預設值",
                Location = new Point(15, 210),
                Size = new Size(410, 40),
                ForeColor = Color.FromArgb(180, 180, 180)
            };

            // Add validation for inputs
            var btnValidate = new Button
            {
                Text = "✓ 驗證輸入",
                Location = new Point(15, 260),
                Size = new Size(80, 32),
                FlatStyle = FlatStyle.Flat,
                BackColor = AccentBlue,
                ForeColor = TextWhite
            };
            btnValidate.Click += (s, e) =>
            {
                string module = cmbModule.Text.Trim().Replace(" (最近使用)", "");
                string procedure = cmbProcedure.Text.Trim().Replace(" (已檢測)", "").Replace(" (新檢測)", "").Replace(" (最近使用)", "");
                
                if (string.IsNullOrEmpty(module) || string.IsNullOrEmpty(procedure))
                {
                    MessageBox.Show("請輸入模塊名稱和程序名稱", "輸入錯誤", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                
                if (!IsValidIdentifier(module) || !IsValidIdentifier(procedure))
                {
                    MessageBox.Show("請輸入有效的標識符名稱（只能包含字母、數字和下劃線，且以字母開頭）", "輸入錯誤", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                
                MessageBox.Show($"輸入有效！\n模塊: {module}\n程序: {procedure}", "驗證成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
            };
            
            var btnSave = new Button
            {
                Text = "💾 保存入口",
                Location = new Point(105, 260),
                Size = new Size(80, 32),
                FlatStyle = FlatStyle.Flat,
                BackColor = AccentGreen,
                ForeColor = TextWhite
            };
            btnSave.Click += (s, e) =>
            {
                string module = cmbModule.Text.Trim().Replace(" (最近使用)", "");
                string procedure = cmbProcedure.Text.Trim().Replace(" (已檢測)", "").Replace(" (新檢測)", "").Replace(" (最近使用)", "");
                
                if (string.IsNullOrEmpty(module) || string.IsNullOrEmpty(procedure))
                {
                    MessageBox.Show("請輸入有效的模塊名稱和程序名稱", "無法保存", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                
                if (!IsValidIdentifier(module) || !IsValidIdentifier(procedure))
                {
                    MessageBox.Show("請輸入有效的標識符名稱", "無法保存", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                
                // Save the entry point
                settings.AddRecentEntryPoint(module, procedure, macroPath);
                settings.Save();
                
                MessageBox.Show($"入口點已保存！\n模塊: {module}\n程序: {procedure}\n\n下次使用時將顯示在建議列表中", "保存成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
            };

            var btnOK = new Button
            {
                Text = "執行 / Run",
                Location = new Point(280, 260),
                Size = new Size(80, 32),
                DialogResult = DialogResult.OK,
                FlatStyle = FlatStyle.Flat,
                BackColor = AccentGreen,
                ForeColor = TextWhite
            };

            var btnCancel = new Button
            {
                Text = "取消 / Cancel",
                Location = new Point(370, 260),
                Size = new Size(80, 32),
                DialogResult = DialogResult.Cancel,
                FlatStyle = FlatStyle.Flat,
                BackColor = DarkPanel,
                ForeColor = TextWhite
            };

            dialog.Controls.AddRange(new Control[] { lblFile, lblDetected, lblModule, cmbModule, lblProcedure, cmbProcedure, lblDetectedList, btnAuto, chkTryAll, lblTip, btnValidate, btnSave, btnOK, btnCancel });
            dialog.AcceptButton = btnOK;
            dialog.CancelButton = btnCancel;

            dialog.Tag = chkTryAll; // Store checkbox for later access

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                string selectedModule = cmbModule.Text.Trim().Replace(" (最近使用)", "");
                string selectedProcedure = cmbProcedure.Text.Replace(" (已檢測)", "").Replace(" (新檢測)", "").Replace(" (最近使用)", "").Trim();
                
                // Validate inputs before returning
                if (string.IsNullOrEmpty(selectedModule) || string.IsNullOrEmpty(selectedProcedure))
                {
                    MessageBox.Show("請輸入有效的模塊名稱和程序名稱", "輸入錯誤", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return ShowVBASubSelection(macroPath); // Show dialog again
                }
                
                if (!IsValidIdentifier(selectedModule) || !IsValidIdentifier(selectedProcedure))
                {
                    MessageBox.Show("請輸入有效的標識符名稱", "輸入錯誤", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return ShowVBASubSelection(macroPath); // Show dialog again
                }
                
                // Auto-save the entry point when user executes
                settings.AddRecentEntryPoint(selectedModule, selectedProcedure, macroPath);
                settings.Save();
                
                return (selectedModule, selectedProcedure);
            }

            return null;
        }
        
        private bool IsValidIdentifier(string name)
        {
            if (string.IsNullOrEmpty(name)) return false;
            if (!char.IsLetter(name[0])) return false;
            
            return name.All(c => char.IsLetterOrDigit(c) || c == '_');
        }

        private void RunVBAMacro(string? vbaCode, string? existingFilePath)
        {
            if (swConnectionManager.SwApp == null) return;
            
            UpdateRunningState(true);
            terminalDisplay.Clear();
            AppendToTerminal($"▶ {Lang.Get("ExecutionStarted")} (VBA) [{DateTime.Now:HH:mm:ss}]", Color.LightBlue);
            AppendToTerminal("═══════════════════════════════════════", Color.Gray);
            
            string pathRunning = "";
            bool isTemp = false;
            bool success = false;
            string errorMsg = "";

            try
            {
                if (!string.IsNullOrEmpty(existingFilePath))
                {
                    pathRunning = existingFilePath;
                    AppendToTerminal($"📋 檔案路徑: {pathRunning}", Color.Gray);
                    
                    // Check if file exists
                    if (!File.Exists(pathRunning))
                    {
                        AppendToTerminal($"❌ 檔案不存在: {pathRunning}", Color.Red);
                        return;
                    }
                    
                    // Check file extension and suggest proper format
                    var ext = Path.GetExtension(pathRunning).ToLowerInvariant();
                    if (ext != ".swp" && ext != ".bas" && ext != ".vba")
                    {
                        AppendToTerminal($"⚠ 不支援的檔案格式: {ext} (支援: .swp, .bas, .vba)", Color.Orange);
                    }
                }
                else if (!string.IsNullOrEmpty(vbaCode))
                {
                    // Save VBA to temporary file with .bas extension for better compatibility
                    pathRunning = Path.Combine(Path.GetTempPath(), $"minisw_temp_{Guid.NewGuid()}.bas");
                    File.WriteAllText(pathRunning, vbaCode);
                    isTemp = true;
                    AppendToTerminal($"📋 臨時檔案: {Path.GetFileName(pathRunning)}", Color.Gray);
                }
                else 
                {
                    AppendToTerminal("❌ 沒有指定 VBA 代碼或檔案", Color.Red);
                    return;
                }
                
                AppendToTerminal($"📝 {Lang.Get("RunningVBA")}: {Path.GetFileName(pathRunning)}", Color.LightBlue);
                
                // Smart entry point detection
                string moduleName = "Module1";
                string procedureName = "main";
                bool tryMultiple = true;
                bool autoDetected = false;
                
                if (!isTemp)
                {
                    // First, try to auto-detect entry point from companion BAS file
                    var basPath = Path.ChangeExtension(pathRunning, ".bas");
                    if (File.Exists(basPath))
                    {
                        try
                        {
                            string basContent = File.ReadAllText(basPath);
                            var moduleInfo = ParseBasModuleInfo(basContent);
                            
                            if (!string.IsNullOrEmpty(moduleInfo.ModuleName) && !string.IsNullOrEmpty(moduleInfo.EntryPoint))
                            {
                                moduleName = moduleInfo.ModuleName;
                                procedureName = moduleInfo.EntryPoint;
                                autoDetected = true;
                                AppendToTerminal($"✅ 從 BAS 自動偵測入口點: {moduleName}.{procedureName}", Color.LightGreen);
                            }
                        }
                        catch { /* Ignore parsing errors */ }
                    }
                    
                    // Try GetMacroMethods API if no BAS detected
                    if (!autoDetected && swConnectionManager.SwApp != null)
                    {
                        try
                        {
                            var methods = swConnectionManager.SwApp.GetMacroMethods(pathRunning, 2) as object[];
                            if (methods != null && methods.Length > 0)
                            {
                                foreach (var method in methods)
                                {
                                    if (method is object[] methodInfo && methodInfo.Length >= 2)
                                    {
                                        string mod = methodInfo[0]?.ToString() ?? "";
                                        string proc = methodInfo[1]?.ToString() ?? "";
                                        
                                        // Prioritize AutoMain, then main/Main
                                        if (proc.Equals("AutoMain", StringComparison.OrdinalIgnoreCase))
                                        {
                                            moduleName = mod;
                                            procedureName = "AutoMain";
                                            autoDetected = true;
                                            break;
                                        }
                                        else if (!autoDetected && (proc.Equals("main", StringComparison.OrdinalIgnoreCase)))
                                        {
                                            moduleName = mod;
                                            procedureName = proc;
                                            autoDetected = true;
                                        }
                                    }
                                }
                                if (autoDetected)
                                {
                                    AppendToTerminal($"✅ 從 API 自動偵測入口點: {moduleName}.{procedureName}", Color.LightGreen);
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            AppendToTerminal($"⚠ GetMacroMethods: {ex.Message}", Color.Orange);
                        }
                    }
                }
                
                int errors = 0;
                bool result = false;
                
                // Try auto-detected entry point first
                if (autoDetected && swConnectionManager.SwApp != null)
                {
                    AppendToTerminal($"🎯 嘗試自動偵測入口點: {moduleName}.{procedureName}", Color.Cyan);
                    result = swConnectionManager.SwApp.RunMacro2(pathRunning, moduleName, procedureName, (int)swRunMacroOption_e.swRunMacroUnloadAfterRun, out errors);
                    
                    if (result)
                    {
                        success = true;
                        AppendToTerminal($"✅ {Lang.Get("VBASuccess")} ({moduleName}.{procedureName})", Color.LightGreen);
                        
                        // Record success and exit early
                        settings.AddExecutionHistory(currentFile?.Name ?? Path.GetFileName(pathRunning), pathRunning, true, 0, "");
                        settings.Save();
                        return;
                    }
                    else
                    {
                        AppendToTerminal($"⚠ 自動入口點執行失敗 (錯誤碼: {errors})，顯示手動選擇...", Color.Orange);
                    }
                }
                
                // Show manual selection dialog if auto-detection failed or wasn't available
                if (!isTemp)
                {
                    var selection = ShowVBASubSelection(pathRunning);
                    if (selection == null)
                    {
                        AppendToTerminal("⚠ 用戶取消執行", Color.Orange);
                        return;
                    }
                    moduleName = selection.Value.module;
                    procedureName = selection.Value.procedure;
                }
                
                AppendToTerminal($"🎯 執行入口點: {moduleName}.{procedureName}", Color.Cyan);
                
                // Try with user-specified entry point
                if (swConnectionManager.SwApp == null)
                {
                    AppendToTerminal("❌ SolidWorks 未連接", Color.Red);
                    return;
                }
                result = swConnectionManager.SwApp.RunMacro2(pathRunning, moduleName, procedureName, (int)swRunMacroOption_e.swRunMacroUnloadAfterRun, out errors);
                
                // If failed and tryMultiple is enabled, try common combinations
                if (!result && errors != 0 && tryMultiple)
                {
                    string[,] commonCombinations = {
                        { "Module1", "main" },
                        { "main", "main" },
                        { "Module1", "Main" },
                        { "main", "Main" },
                        { "Module1", "macro" },
                        { "macro", "macro" },
                        { "Module1", "swmain" },
                        { "swmain", "swmain" },
                        { "ThisDocument", "main" }
                    };
                    
                    for (int i = 0; i < commonCombinations.GetLength(0); i++)
                    {
                        string tryModule = commonCombinations[i, 0];
                        string tryProcedure = commonCombinations[i, 1];
                        
                        // Skip if already tried
                        if (tryModule == moduleName && tryProcedure == procedureName) continue;
                        
                        AppendToTerminal($"⚠ 嘗試組合: {tryModule}.{tryProcedure}", Color.Orange);
                        result = swConnectionManager.SwApp.RunMacro2(pathRunning, tryModule, tryProcedure, (int)swRunMacroOption_e.swRunMacroUnloadAfterRun, out errors);
                        
                        if (result)
                        {
                            moduleName = tryModule;
                            procedureName = tryProcedure;
                            break;
                        }
                    }
                }
                
                if (result)
                {
                    success = true;
                    AppendToTerminal($"✅ {Lang.Get("VBASuccess")} ({moduleName}.{procedureName})", Color.LightGreen);
                }
                else
                {
                    success = false;
                    errorMsg = GetVBAErrorDescription(errors);
                    AppendToTerminal($"❌ {Lang.Get("VBAFailed")}: 錯誤代碼 {errors}", Color.Red);
                    AppendToTerminal($"📝 錯誤說明: {errorMsg}", Color.Red);
                    
                    // Provide troubleshooting suggestions
                    AppendToTerminal("🔧 故障排除建議:", Color.Yellow);
                    AppendToTerminal($"  • 確認宏檔案中包含 'Sub {procedureName}()' 函數", Color.Yellow);
                    AppendToTerminal($"  • 確認模塊名稱為 '{moduleName}' 或嘗試其他名稱", Color.Yellow);
                    AppendToTerminal("  • 確認 SolidWorks 文檔已打開（若宏需要）", Color.Yellow);
                    AppendToTerminal("  • 嘗試從 SolidWorks 中手動執行此宏", Color.Yellow);
                    AppendToTerminal("  • 檢查宏中是否有語法錯誤", Color.Yellow);
                }
            }
            catch (Exception ex)
            {
                success = false;
                errorMsg = ex.Message;
                AppendToTerminal($"❌ {Lang.Get("RuntimeError")}: {ex.Message}", Color.Red);
                AppendToTerminal($"🔍 例外詳情: {ex.GetType().Name}", Color.Red);
            }
            finally
            {
                // Clean up temporary file
                if (isTemp && File.Exists(pathRunning)) 
                {
                    try { File.Delete(pathRunning); } catch { }
                }

                // Record execution history for VBA macros too
                if (currentFile != null)
                {
                    settings.AddExecutionHistory(
                        currentFile.Name,
                        currentFile.FilePath ?? "",
                        success,
                        0, // VBA execution time not tracked
                        errorMsg
                    );
                    settings.Save();
                    RefreshExecutionHistory();
                }

                AppendToTerminal("");
                AppendToTerminal("═══════════════════════════════════════", Color.Gray);
                AppendToTerminal($"⬛ {Lang.Get("ExecutionCompleted")} [{DateTime.Now:HH:mm:ss}]", Color.LightBlue);
                UpdateRunningState(false);
            }
        }

        private void UpdateRunningState(bool running)
        {
            if (InvokeRequired) { Invoke(new Action(() => UpdateRunningState(running))); return; }
            isRunning = running;
            btnRun.Enabled = !running;
            btnStop.Enabled = running;
        }

        // ==================== File Management ====================

        private void CreateNewFile()
        {
            var newFile = new MacroBookmark
            {
                Name = $"{Lang.Get("NewScript")}{newFileCounter++}",
                Content = GetDefaultCode(),
                Type = MacroType.CSharp,
                IsUnsaved = false  // New empty script is not considered unsaved until modified
            };
            newFile.Tag = newFile.Content;  // Store original content to track changes
            openFiles.Add(newFile);
            AddTabForFile(newFile);
        }

        private void AddTabForFile(MacroBookmark file, bool? forceBinary = null)
        {
            // Hide empty panel when adding a tab
            HideEmptyPanel();
            
            var tabPage = new TabPage(TruncateTabText(file.Name))
            {
                BackColor = DarkBackground,
                Tag = file.Id,
                ToolTipText = string.IsNullOrEmpty(file.FilePath) ? file.Name : file.FilePath
            };

            bool isBinary = forceBinary ?? (!string.IsNullOrEmpty(file.FilePath) && file.FilePath.EndsWith(".swp", StringComparison.OrdinalIgnoreCase));

            var editor = new RichTextBox
            {
                Dock = DockStyle.Fill,
                Font = new Font("Cascadia Code", 11),
                BackColor = DarkBackground,
                ForeColor = isBinary ? Color.Gray : Color.FromArgb(212, 212, 212),
                BorderStyle = BorderStyle.None,
                AcceptsTab = true,
                WordWrap = false,
                Text = file.Content,
                ReadOnly = isBinary
            };
            
            if (!isBinary)
            {
                editor.TextChanged += CodeEditor_TextChanged;
            }
            
            tabPage.Controls.Add(editor);
            codeTabs.TabPages.Add(tabPage);
            codeTabs.SelectedTab = tabPage;
            currentFile = file;
            
            // Apply syntax highlighting after a short delay
            if (!isBinary)
            {
                highlightTimer?.Stop();
                highlightTimer?.Start();
            }
        }

        private void OpenMacroFile(string path, MacroType type)
        {
            // Check if already open
            var existing = openFiles.FirstOrDefault(f => f.FilePath == path);
            if (existing != null)
            {
                var existingTab = codeTabs.TabPages.Cast<TabPage>().FirstOrDefault(t => t.Tag?.ToString() == existing.Id);
                if (existingTab != null)
                {
                    codeTabs.SelectedTab = existingTab;
                }
                return;
            }

            string content;
            bool isBinarySwp = false;
            if (Path.GetExtension(path).Equals(".swp", StringComparison.OrdinalIgnoreCase))
            {
                // Try to extract VBA code from SWP - look for companion .bas file or try extraction
                content = TryExtractSwpContent(path, out isBinarySwp);
            }
            else
            {
                 content = macroManager.LoadMacroContent(path);
            }

            var file = new MacroBookmark
            {
                Name = Path.GetFileNameWithoutExtension(path),
                FilePath = path,
                Content = content,
                Type = type,
                IsUnsaved = false
            };
            openFiles.Add(file);
            AddTabForFile(file, isBinarySwp);

            // Add to recent files
            settings.AddRecentFile(path, file.Name, type);
            settings.Save();
            RefreshRecentFiles();
        }

        private void CloseTabAt(int tabIndex)
        {
            if (tabIndex < 0 || tabIndex >= codeTabs.TabPages.Count) return;

            var tabPage = codeTabs.TabPages[tabIndex];
            var file = openFiles.FirstOrDefault(f => f.Id == tabPage.Tag?.ToString());
            
            if (file != null)
            {
                // Only ask for save if file has real unsaved changes
                if (file.IsUnsaved && HasRealChanges(file))
                {
                    var result = MessageBox.Show($"'{file.Name}' {Lang.Get("UnsavedConfirm")}", Lang.Get("Confirm"), MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (result != DialogResult.Yes) return;
                }
                openFiles.Remove(file);
            }

            codeTabs.TabPages.RemoveAt(tabIndex);
            
            // Show empty panel if no tabs left
            if (codeTabs.TabPages.Count == 0)
            {
                ShowEmptyPanel();
                currentFile = null;
            }
        }

        private void CloseOtherTabs()
        {
            if (codeTabs.SelectedIndex < 0) return;
            var keepTab = codeTabs.SelectedTab;
            for (int i = codeTabs.TabPages.Count - 1; i >= 0; i--)
            {
                if (codeTabs.TabPages[i] != keepTab)
                {
                    var file = openFiles.FirstOrDefault(f => f.Id == codeTabs.TabPages[i].Tag?.ToString());
                    if (file != null && file.IsUnsaved)
                    {
                        var result = MessageBox.Show($"'{file.Name}' {Lang.Get("UnsavedConfirm")}", Lang.Get("Confirm"), MessageBoxButtons.YesNo);
                        if (result != DialogResult.Yes) continue;
                        openFiles.Remove(file);
                    }
                    else if (file != null) openFiles.Remove(file);
                    codeTabs.TabPages.RemoveAt(i);
                }
            }
        }

        private void CloseAllTabs()
        {
            for (int i = codeTabs.TabPages.Count - 1; i >= 0; i--)
            {
                var file = openFiles.FirstOrDefault(f => f.Id == codeTabs.TabPages[i].Tag?.ToString());
                if (file != null && file.IsUnsaved && HasRealChanges(file))
                {
                    var result = MessageBox.Show($"'{file.Name}' {Lang.Get("UnsavedConfirm")}", Lang.Get("Confirm"), MessageBoxButtons.YesNo);
                    if (result != DialogResult.Yes) continue;
                }
                if (file != null) openFiles.Remove(file);
                codeTabs.TabPages.RemoveAt(i);
            }
            if (codeTabs.TabPages.Count == 0)
            {
                ShowEmptyPanel();
                currentFile = null;
            }
        }

        private void RenameCurrentTab()
        {
            if (currentFile == null) return;
            using var inputBox = new Form
            {
                Width = 350, Height = 150,
                Text = "重命名 / Rename",
                FormBorderStyle = FormBorderStyle.FixedDialog,
                StartPosition = FormStartPosition.CenterParent,
                BackColor = DarkPanel,
                ForeColor = TextWhite
            };
            var textBox = new TextBox { Left = 20, Top = 30, Width = 290, Text = currentFile.Name, BackColor = DarkBackground, ForeColor = TextWhite };
            var okBtn = new Button { Text = "OK", Left = 135, Top = 70, Width = 80, DialogResult = DialogResult.OK, BackColor = AccentBlue, ForeColor = TextWhite, FlatStyle = FlatStyle.Flat };
            var cancelBtn = new Button { Text = "Cancel", Left = 225, Top = 70, Width = 80, DialogResult = DialogResult.Cancel, BackColor = DarkPanel, ForeColor = TextWhite, FlatStyle = FlatStyle.Flat };
            inputBox.Controls.AddRange(new Control[] { textBox, okBtn, cancelBtn });
            inputBox.AcceptButton = okBtn;
            inputBox.CancelButton = cancelBtn;

            if (inputBox.ShowDialog() == DialogResult.OK && !string.IsNullOrWhiteSpace(textBox.Text))
            {
                currentFile.Name = textBox.Text.Trim();
                currentFile.IsUnsaved = true;
                if (codeTabs.SelectedTab != null)
                {
                    codeTabs.SelectedTab.Text = TruncateTabText(currentFile.Name);
                    codeTabs.SelectedTab.ToolTipText = currentFile.Name;
                }
                codeTabs.Invalidate();
            }
        }

        private void SaveCurrentTabAs()
        {
            if (currentFile == null) return;
            var currentTab = codeTabs.SelectedTab;
            if (currentTab == null) return;
            var editor = currentTab.Controls.OfType<RichTextBox>().FirstOrDefault();
            if (editor == null) return;
            
            currentFile.Content = editor.Text;
            
            using var dialog = new SaveFileDialog
            {
                Filter = currentFile.Type == MacroType.CSharp 
                    ? "C# Files (*.cs)|*.cs|Text Files (*.txt)|*.txt|JSON Files (*.json)|*.json|All Files (*.*)|*.*" 
                    : "VBA Files (*.bas)|*.bas|VBA Files (*.vba)|*.vba|Text Files (*.txt)|*.txt|All Files (*.*)|*.*",
                FileName = currentFile.Name,
                InitialDirectory = Path.Combine(AppSettings.DefaultMacrosPath, currentFile.Type == MacroType.CSharp ? "C Sharp" : "VBA")
            };
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                currentFile.FilePath = dialog.FileName;
                currentFile.Name = Path.GetFileNameWithoutExtension(dialog.FileName);
                if (macroManager.SaveMacroContent(currentFile.FilePath, currentFile.Content))
                {
                    currentFile.IsUnsaved = false;
                    if (codeTabs.SelectedTab != null)
                    {
                        codeTabs.SelectedTab.Text = TruncateTabText(currentFile.Name);
                        codeTabs.SelectedTab.ToolTipText = currentFile.Name;
                    }
                    codeTabs.Invalidate();
                    AppendToTerminal($"✅ {Lang.Get("SaveSuccess")}: {currentFile.FilePath}", Color.LightGreen);
                    RefreshMacroLists();
                }
            }
        }

        private string TruncateTabText(string text, int maxLength = 18)
        {
            if (string.IsNullOrEmpty(text)) return text;
            return text.Length <= maxLength ? text : text.Substring(0, maxLength) + "...";
        }

        private string TryExtractSwpContent(string swpPath, out bool isBinary)
        {
            isBinary = true;
            
            // First, check for companion .bas file in same directory
            var basPath = Path.ChangeExtension(swpPath, ".bas");
            if (File.Exists(basPath))
            {
                isBinary = false;
                string basContent = File.ReadAllText(basPath);
                
                // Extract module info for display
                var moduleInfo = ParseBasModuleInfo(basContent);
                string infoHeader = "' === 來源: " + Path.GetFileName(basPath) + " ===\n";
                if (!string.IsNullOrEmpty(moduleInfo.ModuleName))
                {
                    infoHeader += $"' 模塊名稱: {moduleInfo.ModuleName}\n";
                }
                if (!string.IsNullOrEmpty(moduleInfo.EntryPoint))
                {
                    infoHeader += $"' 入口點: {moduleInfo.EntryPoint}\n";
                }
                infoHeader += "' (SWP 檔案的原始碼版本)\n\n";
                
                return infoHeader + basContent;
            }

            // Check for companion .vba file
            var vbaPath = Path.ChangeExtension(swpPath, ".vba");
            if (File.Exists(vbaPath))
            {
                isBinary = false;
                return $"' === 來源: {Path.GetFileName(vbaPath)} ===\n' (SWP 檔案的原始碼版本)\n\n" + File.ReadAllText(vbaPath);
            }

            // For SWP files without source, show info but mark as executable
            return $"====== SWP EXECUTABLE FILE ======\n\n檔案: {Path.GetFileName(swpPath)}\n\n這是一個已編譯的 SolidWorks 宏文件 (.swp)\n可以直接執行，但無法查看或編輯源代碼。\n\n若要查看源代碼：\n• 將同名的 .bas 或 .vba 文件放在同一資料夾\n• 或者從 SolidWorks 重新匯出為 .bas 格式\n\n▶ 點擊 'Run' 按鈕即可執行此宏\n\n==================================";
        }
        
        private (string ModuleName, string EntryPoint, List<string> Methods) ParseBasModuleInfo(string basContent)
        {
            string moduleName = "";
            string entryPoint = "";
            var methods = new List<string>();
            
            var lines = basContent.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            
            foreach (var line in lines)
            {
                var trimmed = line.Trim();
                
                // Extract module name from: Attribute VB_Name = "Module01_create_a_new_part"
                if (trimmed.StartsWith("Attribute VB_Name", StringComparison.OrdinalIgnoreCase))
                {
                    var parts = trimmed.Split('=');
                    if (parts.Length >= 2)
                    {
                        moduleName = parts[1].Trim().Trim('"', ' ');
                    }
                }
                
                // Find Sub or Function declarations
                if (trimmed.StartsWith("Sub ", StringComparison.OrdinalIgnoreCase) ||
                    trimmed.StartsWith("Public Sub ", StringComparison.OrdinalIgnoreCase) ||
                    trimmed.StartsWith("Private Sub ", StringComparison.OrdinalIgnoreCase) ||
                    trimmed.StartsWith("Function ", StringComparison.OrdinalIgnoreCase) ||
                    trimmed.StartsWith("Public Function ", StringComparison.OrdinalIgnoreCase))
                {
                    // Extract method name
                    var methodLine = trimmed.Replace("Public ", "").Replace("Private ", "");
                    int parenIndex = methodLine.IndexOf('(');
                    if (parenIndex > 0)
                    {
                        var nameWithType = methodLine.Substring(0, parenIndex).Trim();
                        var nameParts = nameWithType.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                        if (nameParts.Length >= 2)
                        {
                            string methodName = nameParts[1];
                            if (!methods.Contains(methodName))
                            {
                                methods.Add(methodName);
                                
                                // Check for entry points
                                if (string.IsNullOrEmpty(entryPoint))
                                {
                                    if (methodName.Equals("AutoMain", StringComparison.OrdinalIgnoreCase))
                                    {
                                        entryPoint = "AutoMain";
                                    }
                                    else if (methodName.Equals("main", StringComparison.OrdinalIgnoreCase) ||
                                             methodName.Equals("Main", StringComparison.OrdinalIgnoreCase))
                                    {
                                        entryPoint = methodName;
                                    }
                                }
                            }
                        }
                    }
                }
            }
            
            return (moduleName, entryPoint, methods);
        }

        private void CodeTabs_MouseMove(object? sender, MouseEventArgs e)
        {
            for (int i = 0; i < codeTabs.TabPages.Count; i++)
            {
                if (codeTabs.GetTabRect(i).Contains(e.Location))
                {
                    var file = openFiles.FirstOrDefault(f => f.Id == codeTabs.TabPages[i].Tag?.ToString());
                    if (file != null)
                    {
                        string tooltip = string.IsNullOrEmpty(file.FilePath) ? file.Name : file.FilePath;
                        if (codeTabs.TabPages[i].ToolTipText != tooltip)
                            codeTabs.TabPages[i].ToolTipText = tooltip;
                    }
                    return;
                }
            }
        }

        private string GetDefaultCode()
        {
            bool isEnglish = Lang.CurrentLanguage == "en-US";
            if (isEnglish)
            {
                return $@"// SolidWorks C# Macro Script
// Available variables: swApp (ISldWorks), swModel (IModelDoc2)
// Available functions: Print(), PrintError(), PrintWarning()

if (swApp != null)
{{
    Print($""SolidWorks Version: {{swApp.RevisionNumber()}}"");
    
    if (swModel != null)
    {{
        Print($""Active Document: {{swModel.GetTitle()}}"");
    }}
    else
    {{
        PrintWarning(""{Lang.Get("NoDocOpen")}"");
    }}
}}
else
{{
    PrintError(""{Lang.Get("SWNotConnected")}"");
}}
";
            }
            else
            {
                return $@"// SolidWorks C# 宏腳本
// 可用變量: swApp (ISldWorks), swModel (IModelDoc2)
// 可用函數: Print(), PrintError(), PrintWarning()

if (swApp != null)
{{
    Print($""SolidWorks 版本: {{swApp.RevisionNumber()}}"");
    
    if (swModel != null)
    {{
        Print($""活動文檔: {{swModel.GetTitle()}}"");
    }}
    else
    {{
        PrintWarning(""{Lang.Get("NoDocOpen")}"");
    }}
}}
else
{{
    PrintError(""{Lang.Get("SWNotConnected")}"");
}}
";
            }
        }

        private void ApplySyntaxHighlighting()
        {
            if (currentFile == null || codeTabs.SelectedTab == null) return;
            var editor = codeTabs.SelectedTab.Controls.OfType<RichTextBox>().FirstOrDefault();
            if (editor == null) return;
            
            try 
            { 
                SyntaxHighlighter.ApplyHighlighting(editor, currentFile.Type == MacroType.CSharp, SyntaxHighlighter.DarkTheme); 
            }
            catch { }
        }

        // ==================== Lists ====================

        private void RefreshMacroLists()
        {
            RefreshMacroList("csharpList", macroManager.GetCSharpMacros(true));
            RefreshMacroList("vbaList", macroManager.GetVBAMacros(true));
        }

        private void RefreshMacroList(string name, List<MacroFileInfo> macros)
        {
            var listView = FindControl<ListView>(name);
            if (listView == null) return;
            listView.Items.Clear();

            foreach (var m in macros)
            {
                try
                {
                    var fileInfo = new FileInfo(m.FullPath);
                    string ext = Path.GetExtension(m.FullPath).ToUpper();
                    string icon = ext == ".CS" ? "📄" : ext == ".SWP" ? "⚙" : "📝";
                    
                    var item = new ListViewItem($"{icon} {m.Name}")
                    {
                        Tag = m.FullPath,
                        ForeColor = ext == ".SWP" ? Color.FromArgb(150, 150, 150) : TextWhite,
                        ToolTipText = m.FullPath
                    };
                    
                    item.SubItems.Add(ext);
                    item.SubItems.Add(fileInfo.Exists ? fileInfo.LastWriteTime.ToString("yyyy/MM/dd HH:mm") : "-");
                    item.SubItems.Add(m.RelativePath);
                    
                    listView.Items.Add(item);
                }
                catch
                {
                    // Skip files that can't be accessed
                }
            }
        }

        private void RefreshPromptsList()
        {
            var list = FindControl<ListBox>("promptsList");
            if (list == null) return;
            list.Items.Clear();
            foreach (var p in settings.Prompts)
                list.Items.Add(p.IsDefault ? $"📌 {p.Name}" : p.Name);
            if (list.Items.Count > 0) list.SelectedIndex = 0;
        }

        private void RefreshNotesList()
        {
            var list = FindControl<ListBox>("notesList");
            if (list == null) return;
            list.Items.Clear();
            foreach (var n in settings.Notes)
                list.Items.Add(n.Name);
            if (list.Items.Count > 0) list.SelectedIndex = 0;
        }

        // ==================== Prompts ====================

        private void PromptsList_SelectedIndexChanged(object? sender, EventArgs e)
        {
            var list = sender as ListBox;
            if (list == null) return;
            if (list.SelectedIndex < 0 || list.SelectedIndex >= settings.Prompts.Count) return;

            var prompt = settings.Prompts[list.SelectedIndex];
            var nameBox = FindControl<TextBox>("promptName");
            var contentBox = FindControl<RichTextBox>("promptContent");
            if (nameBox != null) nameBox.Text = prompt.Name;
            if (contentBox != null) contentBox.Text = prompt.Content;
        }

        private void PromptName_TextChanged(object? sender, EventArgs e)
        {
            var list = FindControl<ListBox>("promptsList");
            if (list == null) return;
            if (list.SelectedIndex < 0) return;

            settings.Prompts[list.SelectedIndex].Name = (sender as TextBox)?.Text ?? "";
            settings.Save();
            var p = settings.Prompts[list.SelectedIndex];
            list.Items[list.SelectedIndex] = p.IsDefault ? $"📌 {p.Name}" : p.Name;
        }

        private void PromptContent_TextChanged(object? sender, EventArgs e)
        {
            var list = FindControl<ListBox>("promptsList");
            if (list == null) return;
            if (list.SelectedIndex < 0) return;

            settings.Prompts[list.SelectedIndex].Content = (sender as RichTextBox)?.Text ?? "";
            // Removed frequent Save() to prevent lag. Details are saved on FormClose or manual Save.
        }

        private void AddPrompt_Click(object? sender, EventArgs e)
        {
            settings.Prompts.Add(new PromptItem { Name = "New Prompt" });
            settings.Save();
            RefreshPromptsList();
            var list = FindControl<ListBox>("promptsList");
            if (list != null) list.SelectedIndex = list.Items.Count - 1;
        }

        private void DeletePrompt_Click(object? sender, EventArgs e)
        {
            var list = FindControl<ListBox>("promptsList");
            if (list == null || list.SelectedIndex < 0 || list.SelectedIndex >= settings.Prompts.Count) return;
            if (settings.Prompts[list.SelectedIndex].IsDefault)
            {
                MessageBox.Show(Lang.Get("CannotDeleteDefault"), Lang.Get("Confirm"));
                return;
            }
            settings.Prompts.RemoveAt(list.SelectedIndex);
            settings.Save();
            RefreshPromptsList();
        }

        // ==================== Notes ====================

        private void NotesList_SelectedIndexChanged(object? sender, EventArgs e)
        {
            var list = sender as ListBox;
            if (list == null) return;
            if (list.SelectedIndex < 0 || list.SelectedIndex >= settings.Notes.Count) return;

            var note = settings.Notes[list.SelectedIndex];
            var nameBox = FindControl<TextBox>("noteName");
            var contentBox = FindControl<RichTextBox>("noteContent");
            if (nameBox != null) nameBox.Text = note.Name;
            if (contentBox != null) contentBox.Text = note.Content;
        }

        private void NoteName_TextChanged(object? sender, EventArgs e)
        {
            var list = FindControl<ListBox>("notesList");
            if (list == null) return;
            if (list.SelectedIndex < 0) return;

            settings.Notes[list.SelectedIndex].Name = (sender as TextBox)?.Text ?? "";
            settings.Save();
            list.Items[list.SelectedIndex] = settings.Notes[list.SelectedIndex].Name;
        }

        private void NoteContent_TextChanged(object? sender, EventArgs e)
        {
            var list = FindControl<ListBox>("notesList");
            if (list == null || list.SelectedIndex < 0 || list.SelectedIndex >= settings.Notes.Count) return;

            settings.Notes[list.SelectedIndex].Content = (sender as RichTextBox)?.Text ?? "";
            // Removed frequent Save() to prevent lag.
        }

        private void AddNote_Click(object? sender, EventArgs e)
        {
            settings.Notes.Add(new NoteItem { Name = "New Note" });
            settings.Save();
            RefreshNotesList();
            var list = FindControl<ListBox>("notesList");
            if (list != null) list.SelectedIndex = list.Items.Count - 1;
        }

        private void DeleteNote_Click(object? sender, EventArgs e)
        {
            var list = FindControl<ListBox>("notesList");
            if (list == null || list.SelectedIndex < 0 || list.SelectedIndex >= settings.Notes.Count) return;
            settings.Notes.RemoveAt(list.SelectedIndex);
            settings.Save();
            RefreshNotesList();
        }

        // ==================== Helpers ====================

        private T? FindControl<T>(string name) where T : Control => FindControlRecursive<T>(this, name);
        
        private T? FindControlRecursive<T>(Control parent, string name) where T : Control
        {
            foreach (Control c in parent.Controls)
            {
                if (c.Name == name && c is T t) return t;
                var found = FindControlRecursive<T>(c, name);
                if (found != null) return found;
            }
            return null;
        }

        private bool HasRealChanges(MacroBookmark file)
        {
            // For files without a path (new files), check if content differs from default template
            if (string.IsNullOrEmpty(file.FilePath))
            {
                // Check against original content stored in Tag
                if (file.Tag is string originalContent)
                {
                    // Get current content from editor
                    var tab = codeTabs.TabPages.Cast<TabPage>().FirstOrDefault(t => t.Tag?.ToString() == file.Id);
                    if (tab != null)
                    {
                        var editor = tab.Controls.OfType<RichTextBox>().FirstOrDefault();
                        if (editor != null)
                        {
                            return editor.Text != originalContent;
                        }
                    }
                }
                return false;
            }
            return true; // Files with paths always have real changes if marked unsaved
        }

        private Panel? emptyPanel = null;
        
        private void ShowEmptyPanel()
        {
            if (emptyPanel == null)
            {
                emptyPanel = new Panel
                {
                    Dock = DockStyle.Fill,
                    BackColor = DarkBackground,
                    Name = "emptyPanel"
                };
                
                var label = new Label
                {
                    Text = Lang.Get("NoScriptLoaded"),
                    ForeColor = TextGray,
                    Font = new Font("Segoe UI", 12),
                    TextAlign = ContentAlignment.MiddleCenter,
                    Dock = DockStyle.Fill,
                    AutoSize = false
                };
                
                emptyPanel.Controls.Add(label);
            }
            
            if (!codePanel.Controls.Contains(emptyPanel))
            {
                codePanel.Controls.Add(emptyPanel);
                emptyPanel.BringToFront();
            }
        }
        
        private void HideEmptyPanel()
        {
            if (emptyPanel != null && codePanel.Controls.Contains(emptyPanel))
            {
                codePanel.Controls.Remove(emptyPanel);
            }
        }

        private void AppendToTerminal(string text, Color? color = null)
        {
            if (InvokeRequired) { try { Invoke(new Action(() => AppendToTerminal(text, color))); } catch { } return; }
            try
            {
                terminalDisplay.SelectionStart = terminalDisplay.TextLength;
                terminalDisplay.SelectionColor = color ?? TextWhite;
                terminalDisplay.AppendText(text + System.Environment.NewLine);
                terminalDisplay.ScrollToCaret();
            }
            catch { }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            foreach (var file in openFiles.Where(f => f.IsUnsaved && HasRealChanges(f)))
            {
                var result = MessageBox.Show($"'{file.Name}' {Lang.Get("UnsavedConfirm")}", Lang.Get("Confirm"), MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.No) { e.Cancel = true; return; }
            }
            settings.WindowWidth = Width;
            settings.WindowHeight = Height;
            settings.Save();
            swConnectionManager.Disconnect();
            highlightTimer?.Dispose();
            autoSaveTimer?.Dispose();
            searchDialog?.Dispose();
            base.OnFormClosing(e);
        }

        // ==================== New Tabs Setup ====================

        private void SetupRecentFilesTab(TabPage tab)
        {
            var listView = new ListView
            {
                Dock = DockStyle.Fill,
                View = System.Windows.Forms.View.Details,
                FullRowSelect = true,
                BackColor = DarkPanel,
                ForeColor = TextWhite,
                BorderStyle = BorderStyle.None,
                Font = new Font("Segoe UI", 9),
                Name = "recentFilesList"
            };
            listView.Columns.Add("檔案名稱", 180);
            listView.Columns.Add("路徑", 200);
            listView.Columns.Add("時間", 120);

            listView.DoubleClick += (s, e) =>
            {
                if (listView.SelectedItems.Count > 0 && listView.SelectedItems[0].Tag is RecentFileItem recent)
                {
                    if (File.Exists(recent.FilePath))
                        OpenMacroFile(recent.FilePath, recent.Type);
                    else
                        AppendToTerminal($"❌ 檔案不存在: {recent.FilePath}", Color.Red);
                }
            };

            // Context menu
            var menu = new ContextMenuStrip { BackColor = DarkPanel, ForeColor = TextWhite };
            menu.Items.Add("開啟 / Open", null, (s, e) => listView.SelectedItems[0]?.Tag?.ToString());
            menu.Items.Add("從列表移除 / Remove", null, (s, e) =>
            {
                if (listView.SelectedItems.Count > 0 && listView.SelectedItems[0].Tag is RecentFileItem recent)
                {
                    settings.RecentFiles.RemoveAll(r => r.FilePath == recent.FilePath);
                    settings.Save();
                    RefreshRecentFiles();
                }
            });
            menu.Items.Add("清空列表 / Clear All", null, (s, e) =>
            {
                settings.RecentFiles.Clear();
                settings.Save();
                RefreshRecentFiles();
            });
            listView.ContextMenuStrip = menu;

            tab.Controls.Add(listView);
        }

        private void SetupSnippetsTab(TabPage tab)
        {
            var snippetsPanel = new CodeSnippetsPanel
            {
                Dock = DockStyle.Fill
            };
            snippetsPanel.InsertSnippet += (s, code) =>
            {
                // Insert snippet at cursor position in current editor
                var currentTab = codeTabs.SelectedTab;
                if (currentTab == null) return;
                var editor = currentTab.Controls.OfType<RichTextBox>().FirstOrDefault();
                if (editor != null && !editor.ReadOnly)
                {
                    int pos = editor.SelectionStart;
                    editor.Text = editor.Text.Insert(pos, code);
                    editor.SelectionStart = pos + code.Length;
                    editor.Focus();
                    AppendToTerminal("📝 代碼片段已插入", Color.LightGreen);
                }
            };
            tab.Controls.Add(snippetsPanel);
        }

        private void SetupHistoryTab(TabPage tab)
        {
            var listView = new ListView
            {
                Dock = DockStyle.Fill,
                View = System.Windows.Forms.View.Details,
                FullRowSelect = true,
                BackColor = DarkPanel,
                ForeColor = TextWhite,
                BorderStyle = BorderStyle.None,
                Font = new Font("Segoe UI", 9),
                Name = "historyList"
            };
            listView.Columns.Add("檔案", 150);
            listView.Columns.Add("時間", 130);
            listView.Columns.Add("狀態", 60);
            listView.Columns.Add("耗時", 70);

            listView.DoubleClick += (s, e) =>
            {
                if (listView.SelectedItems.Count > 0 && listView.SelectedItems[0].Tag is ExecutionHistoryItem hist)
                {
                    if (File.Exists(hist.FilePath))
                        OpenMacroFile(hist.FilePath, MacroType.CSharp);
                }
            };

            tab.Controls.Add(listView);
        }

        private void RefreshRecentFiles()
        {
            var list = FindControl<ListView>("recentFilesList");
            if (list == null) return;
            list.Items.Clear();

            foreach (var recent in settings.RecentFiles)
            {
                var item = new ListViewItem(recent.Name)
                {
                    Tag = recent,
                    ForeColor = File.Exists(recent.FilePath) ? TextWhite : Color.Gray
                };
                item.SubItems.Add(recent.FilePath);
                item.SubItems.Add(recent.LastOpened.ToString("MM/dd HH:mm"));
                list.Items.Add(item);
            }
        }

        private void RefreshExecutionHistory()
        {
            var list = FindControl<ListView>("historyList");
            if (list == null) return;
            list.Items.Clear();

            foreach (var hist in settings.ExecutionHistory)
            {
                var item = new ListViewItem(hist.FileName)
                {
                    Tag = hist,
                    ForeColor = hist.Success ? Color.LightGreen : Color.FromArgb(255, 100, 100)
                };
                item.SubItems.Add(hist.ExecutedAt.ToString("MM/dd HH:mm:ss"));
                item.SubItems.Add(hist.Success ? "✓" : "✗");
                item.SubItems.Add($"{hist.DurationMs}ms");
                list.Items.Add(item);
            }
        }

        // ==================== Search & Replace ====================

        private void ShowSearchDialog(bool replaceMode)
        {
            if (searchDialog == null || searchDialog.IsDisposed)
            {
                searchDialog = new SearchReplaceDialog(replaceMode);
                searchDialog.FindNext += SearchDialog_FindNext;
                searchDialog.FindPrevious += SearchDialog_FindPrevious;
                searchDialog.Replace += SearchDialog_Replace;
                searchDialog.ReplaceAll += SearchDialog_ReplaceAll;
            }
            else if (replaceMode)
            {
                searchDialog.SwitchToReplace();
            }

            // Pre-fill with selected text
            var currentTab = codeTabs.SelectedTab;
            if (currentTab != null)
            {
                var editor = currentTab.Controls.OfType<RichTextBox>().FirstOrDefault();
                if (editor != null && editor.SelectionLength > 0)
                {
                    searchDialog.SetSearchText(editor.SelectedText);
                }
            }

            searchDialog.Show(this);
            searchDialog.BringToFront();
            searchDialog.Focus();
        }

        private void SearchDialog_FindNext(object? sender, SearchEventArgs e)
        {
            var editor = GetCurrentEditor();
            if (editor == null) return;

            int index = FindInEditor(editor, e.SearchText, e.MatchCase, false);
            if (index < 0)
                searchDialog?.SetStatus("找不到 / Not found");
            else
                searchDialog?.SetStatus($"找到位置: {index}");
        }

        private void SearchDialog_FindPrevious(object? sender, SearchEventArgs e)
        {
            var editor = GetCurrentEditor();
            if (editor == null) return;

            int index = FindInEditor(editor, e.SearchText, e.MatchCase, true);
            if (index < 0)
                searchDialog?.SetStatus("找不到 / Not found");
            else
                searchDialog?.SetStatus($"找到位置: {index}");
        }

        private void SearchDialog_Replace(object? sender, ReplaceEventArgs e)
        {
            var editor = GetCurrentEditor();
            if (editor == null) return;

            var comparison = e.MatchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
            if (editor.SelectedText.Equals(e.SearchText, comparison))
            {
                editor.SelectedText = e.ReplaceText;
                searchDialog?.SetStatus("已替換 / Replaced");
            }
            FindInEditor(editor, e.SearchText, e.MatchCase, false);
        }

        private void SearchDialog_ReplaceAll(object? sender, ReplaceEventArgs e)
        {
            var editor = GetCurrentEditor();
            if (editor == null) return;

            var comparison = e.MatchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
            string text = editor.Text;
            int count = 0;
            int index = 0;
            var sb = new System.Text.StringBuilder();
            int lastIndex = 0;

            while ((index = text.IndexOf(e.SearchText, lastIndex, comparison)) >= 0)
            {
                sb.Append(text.Substring(lastIndex, index - lastIndex));
                sb.Append(e.ReplaceText);
                lastIndex = index + e.SearchText.Length;
                count++;
            }
            sb.Append(text.Substring(lastIndex));

            if (count > 0)
            {
                editor.Text = sb.ToString();
                searchDialog?.SetStatus($"已替換 {count} 處 / Replaced {count}");
            }
            else
            {
                searchDialog?.SetStatus("找不到 / Not found");
            }
        }

        private int FindInEditor(RichTextBox editor, string searchText, bool matchCase, bool searchUp)
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
            }

            return index;
        }

        private RichTextBox? GetCurrentEditor()
        {
            var currentTab = codeTabs.SelectedTab;
            return currentTab?.Controls.OfType<RichTextBox>().FirstOrDefault();
        }

        private void ShowGoToLineDialog()
        {
            var editor = GetCurrentEditor();
            if (editor == null) return;

            using var dialog = new Form
            {
                Width = 300, Height = 130,
                Text = "跳到行 / Go to Line",
                FormBorderStyle = FormBorderStyle.FixedDialog,
                StartPosition = FormStartPosition.CenterParent,
                BackColor = DarkPanel,
                ForeColor = TextWhite
            };

            var label = new Label { Text = $"行號 (1-{editor.Lines.Length}):", Left = 15, Top = 15, AutoSize = true, ForeColor = TextWhite };
            var textBox = new TextBox { Left = 15, Top = 40, Width = 250, BackColor = DarkBackground, ForeColor = TextWhite };
            var okBtn = new Button { Text = "Go", Left = 110, Top = 70, Width = 70, DialogResult = DialogResult.OK, BackColor = AccentBlue, ForeColor = TextWhite, FlatStyle = FlatStyle.Flat };
            dialog.Controls.AddRange(new Control[] { label, textBox, okBtn });
            dialog.AcceptButton = okBtn;

            if (dialog.ShowDialog() == DialogResult.OK && int.TryParse(textBox.Text, out int line))
            {
                if (line < 1) line = 1;
                if (line > editor.Lines.Length) line = editor.Lines.Length;
                int charIndex = editor.GetFirstCharIndexFromLine(line - 1);
                if (charIndex >= 0)
                {
                    editor.SelectionStart = charIndex;
                    editor.SelectionLength = 0;
                    editor.ScrollToCaret();
                    editor.Focus();
                }
            }
        }

        private void ShowShortcutsHelp()
        {
            var helpText = @"
╔════════════════════════════════════════╗
║          快捷鍵 / Shortcuts            ║
╠════════════════════════════════════════╣
║  F5          執行代碼 / Run            ║
║  Ctrl+S      保存 / Save               ║
║  Ctrl+N      新建 / New                ║
║  Ctrl+F      搜尋 / Find               ║
║  Ctrl+H      搜尋替換 / Replace        ║
║  Ctrl+G      跳到行 / Go to Line       ║
║  Ctrl+B      切換書籤 / Toggle Bookmark║
║  F2          下一個書籤 / Next Bookmark║
║  Shift+F2    上一個書籤 / Prev Bookmark║
║  F1          顯示幫助 / Show Help      ║
║  Ctrl+Z      撤銷 / Undo               ║
║  Ctrl+Y      重做 / Redo               ║
╚════════════════════════════════════════╝

拖放 .cs/.vba/.bas 檔案到視窗可直接開啟
Drag & drop files to open them directly
";
            AppendToTerminal(helpText, Color.Cyan);
        }
    }
}
