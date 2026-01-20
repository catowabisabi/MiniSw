using System;
using System.Drawing;
using System.Windows.Forms;

namespace MiniSolidworkAutomator.Controls
{
    /// <summary>
    /// Search and Replace dialog
    /// </summary>
    public class SearchReplaceDialog : Form
    {
        private TextBox txtSearch = null!;
        private TextBox txtReplace = null!;
        private CheckBox chkMatchCase = null!;
        private Button btnFindNext = null!;
        private Button btnFindPrev = null!;
        private Button btnReplace = null!;
        private Button btnReplaceAll = null!;
        private Label lblStatus = null!;
        private bool isReplaceMode;

        // Theme colors
        private static readonly Color DarkBackground = Color.FromArgb(45, 45, 45);
        private static readonly Color DarkPanel = Color.FromArgb(60, 60, 60);
        private static readonly Color TextWhite = Color.White;
        private static readonly Color AccentBlue = Color.FromArgb(33, 150, 243);

        public string SearchText => txtSearch.Text;
        public string ReplaceText => txtReplace.Text;
        public bool MatchCase => chkMatchCase.Checked;

        public event EventHandler<SearchEventArgs>? FindNext;
        public event EventHandler<SearchEventArgs>? FindPrevious;
        public event EventHandler<ReplaceEventArgs>? Replace;
        public event EventHandler<ReplaceEventArgs>? ReplaceAll;

        public SearchReplaceDialog(bool replaceMode = false)
        {
            isReplaceMode = replaceMode;
            InitializeComponents();
        }

        public void SetSearchText(string text)
        {
            txtSearch.Text = text;
            txtSearch.SelectAll();
        }

        public void SetStatus(string message)
        {
            lblStatus.Text = message;
        }

        public void SwitchToReplace()
        {
            if (!isReplaceMode)
            {
                isReplaceMode = true;
                this.Height = 200;
                txtReplace.Visible = true;
                btnReplace.Visible = true;
                btnReplaceAll.Visible = true;
                this.Text = "搜尋和替換 / Search & Replace";
            }
        }

        private void InitializeComponents()
        {
            this.Text = isReplaceMode ? "搜尋和替換 / Search & Replace" : "搜尋 / Search";
            this.Size = new Size(420, isReplaceMode ? 200 : 150);
            this.FormBorderStyle = FormBorderStyle.FixedToolWindow;
            this.StartPosition = FormStartPosition.CenterParent;
            this.BackColor = DarkBackground;
            this.ForeColor = TextWhite;
            this.TopMost = true;
            this.ShowInTaskbar = false;
            this.KeyPreview = true;

            // Search label and textbox
            var lblSearch = new Label
            {
                Text = "搜尋 / Find:",
                Location = new Point(10, 15),
                AutoSize = true,
                ForeColor = TextWhite
            };

            txtSearch = new TextBox
            {
                Location = new Point(100, 12),
                Size = new Size(200, 25),
                BackColor = DarkPanel,
                ForeColor = TextWhite,
                BorderStyle = BorderStyle.FixedSingle
            };
            txtSearch.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Enter)
                {
                    if (e.Shift)
                        btnFindPrev.PerformClick();
                    else
                        btnFindNext.PerformClick();
                    e.Handled = true;
                    e.SuppressKeyPress = true;
                }
            };

            // Replace label and textbox
            var lblReplace = new Label
            {
                Text = "替換 / Replace:",
                Location = new Point(10, 45),
                AutoSize = true,
                ForeColor = TextWhite,
                Visible = isReplaceMode
            };

            txtReplace = new TextBox
            {
                Location = new Point(100, 42),
                Size = new Size(200, 25),
                BackColor = DarkPanel,
                ForeColor = TextWhite,
                BorderStyle = BorderStyle.FixedSingle,
                Visible = isReplaceMode
            };

            // Options
            chkMatchCase = new CheckBox
            {
                Text = "區分大小寫 / Match Case",
                Location = new Point(10, isReplaceMode ? 75 : 45),
                AutoSize = true,
                ForeColor = TextWhite
            };

            // Buttons
            int buttonY = isReplaceMode ? 105 : 75;

            btnFindNext = CreateButton("下一個 ▼", new Point(10, buttonY));
            btnFindNext.Click += (s, e) => FindNext?.Invoke(this, new SearchEventArgs(SearchText, MatchCase));

            btnFindPrev = CreateButton("上一個 ▲", new Point(95, buttonY));
            btnFindPrev.Click += (s, e) => FindPrevious?.Invoke(this, new SearchEventArgs(SearchText, MatchCase));

            btnReplace = CreateButton("替換", new Point(180, buttonY));
            btnReplace.Visible = isReplaceMode;
            btnReplace.Click += (s, e) => Replace?.Invoke(this, new ReplaceEventArgs(SearchText, ReplaceText, MatchCase));

            btnReplaceAll = CreateButton("全部替換", new Point(260, buttonY));
            btnReplaceAll.Visible = isReplaceMode;
            btnReplaceAll.Click += (s, e) => ReplaceAll?.Invoke(this, new ReplaceEventArgs(SearchText, ReplaceText, MatchCase));

            // Status label
            lblStatus = new Label
            {
                Location = new Point(10, buttonY + 35),
                Size = new Size(380, 20),
                ForeColor = Color.LightGray
            };

            // Add controls
            this.Controls.AddRange(new Control[] { lblSearch, txtSearch, lblReplace, txtReplace, chkMatchCase, btnFindNext, btnFindPrev, btnReplace, btnReplaceAll, lblStatus });

            // Handle Escape key
            this.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Escape)
                {
                    this.Hide();
                    e.Handled = true;
                }
            };
        }

        private Button CreateButton(string text, Point location)
        {
            return new Button
            {
                Text = text,
                Location = location,
                Size = new Size(80, 28),
                FlatStyle = FlatStyle.Flat,
                BackColor = DarkPanel,
                ForeColor = TextWhite,
                Cursor = Cursors.Hand
            };
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
                this.Hide();
            }
            base.OnFormClosing(e);
        }
    }

    public class SearchEventArgs : EventArgs
    {
        public string SearchText { get; }
        public bool MatchCase { get; }

        public SearchEventArgs(string searchText, bool matchCase)
        {
            SearchText = searchText;
            MatchCase = matchCase;
        }
    }

    public class ReplaceEventArgs : EventArgs
    {
        public string SearchText { get; }
        public string ReplaceText { get; }
        public bool MatchCase { get; }

        public ReplaceEventArgs(string searchText, string replaceText, bool matchCase)
        {
            SearchText = searchText;
            ReplaceText = replaceText;
            MatchCase = matchCase;
        }
    }
}
