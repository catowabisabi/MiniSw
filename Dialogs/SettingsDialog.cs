using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using MiniSolidworkAutomator.Models;

namespace MiniSolidworkAutomator.Dialogs
{
    public class SettingsDialog : Form
    {
        private ListBox pathListBox = null!;
        private Button addButton = null!;
        private Button removeButton = null!;
        private Button moveUpButton = null!;
        private Button moveDownButton = null!;
        private Button okButton = null!;
        private Button cancelButton = null!;

        public List<string> MacroPaths { get; private set; } = new List<string>();

        public SettingsDialog(List<string> currentPaths)
        {
            MacroPaths = new List<string>(currentPaths);
            InitializeUI();
            LoadPaths();
        }

        private void InitializeUI()
        {
            this.Text = "設定 - 宏文件路徑";
            this.Size = new Size(600, 450);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.BackColor = Color.FromArgb(245, 245, 245);

            // Title label
            var titleLabel = new Label
            {
                Text = "宏文件搜索路徑",
                Location = new Point(20, 15),
                Size = new Size(200, 25),
                Font = new Font("Microsoft YaHei UI", 11, FontStyle.Bold),
                ForeColor = Color.FromArgb(50, 50, 50)
            };

            var descLabel = new Label
            {
                Text = "程序會自動掃描以下路徑及其子文件夾中的所有 C# (.cs) 和 VBA (.vba, .bas, .swp) 文件",
                Location = new Point(20, 45),
                Size = new Size(540, 35),
                Font = new Font("Microsoft YaHei UI", 9),
                ForeColor = Color.FromArgb(100, 100, 100)
            };

            // Path list
            pathListBox = new ListBox
            {
                Location = new Point(20, 85),
                Size = new Size(440, 250),
                Font = new Font("Consolas", 10),
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = Color.White
            };
            pathListBox.SelectedIndexChanged += PathListBox_SelectedIndexChanged;

            // Buttons panel
            addButton = CreateButton("添加路徑", new Point(470, 85), Color.FromArgb(0, 122, 204));
            addButton.Click += AddButton_Click;

            removeButton = CreateButton("移除", new Point(470, 125), Color.FromArgb(200, 80, 80));
            removeButton.Click += RemoveButton_Click;
            removeButton.Enabled = false;

            moveUpButton = CreateButton("上移 ↑", new Point(470, 175), Color.FromArgb(100, 100, 100));
            moveUpButton.Click += MoveUpButton_Click;
            moveUpButton.Enabled = false;

            moveDownButton = CreateButton("下移 ↓", new Point(470, 215), Color.FromArgb(100, 100, 100));
            moveDownButton.Click += MoveDownButton_Click;
            moveDownButton.Enabled = false;

            // Bottom buttons
            okButton = new Button
            {
                Text = "確定",
                Location = new Point(380, 360),
                Size = new Size(90, 35),
                BackColor = Color.FromArgb(0, 122, 204),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Microsoft YaHei UI", 9, FontStyle.Bold),
                DialogResult = DialogResult.OK
            };
            okButton.FlatAppearance.BorderSize = 0;

            cancelButton = new Button
            {
                Text = "取消",
                Location = new Point(480, 360),
                Size = new Size(90, 35),
                BackColor = Color.FromArgb(180, 180, 180),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Microsoft YaHei UI", 9),
                DialogResult = DialogResult.Cancel
            };
            cancelButton.FlatAppearance.BorderSize = 0;

            // Add controls
            this.Controls.AddRange(new Control[] {
                titleLabel, descLabel, pathListBox,
                addButton, removeButton, moveUpButton, moveDownButton,
                okButton, cancelButton
            });
        }

        private Button CreateButton(string text, Point location, Color backColor)
        {
            var btn = new Button
            {
                Text = text,
                Location = location,
                Size = new Size(95, 30),
                BackColor = backColor,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Microsoft YaHei UI", 9)
            };
            btn.FlatAppearance.BorderSize = 0;
            return btn;
        }

        private void LoadPaths()
        {
            pathListBox.Items.Clear();
            foreach (var path in MacroPaths)
            {
                pathListBox.Items.Add(path);
            }
        }

        private void PathListBox_SelectedIndexChanged(object? sender, EventArgs e)
        {
            bool hasSelection = pathListBox.SelectedIndex >= 0;
            bool isDefault = hasSelection && 
                pathListBox.SelectedItem?.ToString()?.Equals(AppSettings.DefaultMacrosPath, StringComparison.OrdinalIgnoreCase) == true;

            removeButton.Enabled = hasSelection && !isDefault;
            moveUpButton.Enabled = hasSelection && pathListBox.SelectedIndex > 0;
            moveDownButton.Enabled = hasSelection && pathListBox.SelectedIndex < pathListBox.Items.Count - 1;
        }

        private void AddButton_Click(object? sender, EventArgs e)
        {
            using var dialog = new FolderBrowserDialog
            {
                Description = "選擇包含宏文件的文件夾",
                ShowNewFolderButton = true
            };

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                string path = dialog.SelectedPath;
                if (!MacroPaths.Contains(path))
                {
                    MacroPaths.Add(path);
                    pathListBox.Items.Add(path);
                }
                else
                {
                    MessageBox.Show("此路徑已存在於列表中。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        private void RemoveButton_Click(object? sender, EventArgs e)
        {
            if (pathListBox.SelectedIndex >= 0)
            {
                int index = pathListBox.SelectedIndex;
                MacroPaths.RemoveAt(index);
                pathListBox.Items.RemoveAt(index);
            }
        }

        private void MoveUpButton_Click(object? sender, EventArgs e)
        {
            int index = pathListBox.SelectedIndex;
            if (index > 0)
            {
                string path = MacroPaths[index];
                MacroPaths.RemoveAt(index);
                MacroPaths.Insert(index - 1, path);
                LoadPaths();
                pathListBox.SelectedIndex = index - 1;
            }
        }

        private void MoveDownButton_Click(object? sender, EventArgs e)
        {
            int index = pathListBox.SelectedIndex;
            if (index < MacroPaths.Count - 1)
            {
                string path = MacroPaths[index];
                MacroPaths.RemoveAt(index);
                MacroPaths.Insert(index + 1, path);
                LoadPaths();
                pathListBox.SelectedIndex = index + 1;
            }
        }
    }
}
