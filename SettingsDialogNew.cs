using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using MiniSolidworkAutomator.Models;
using MiniSolidworkAutomator.Localization;

namespace MiniSolidworkAutomator
{
    public class SettingsDialogNew : Form
    {
        // ============ Theme Colors (Dark Theme) ============
        private static readonly Color DarkBackground = Color.FromArgb(30, 30, 30);
        private static readonly Color DarkPanel = Color.FromArgb(45, 45, 45);
        private static readonly Color TextWhite = Color.White;
        private static readonly Color AccentBlue = Color.FromArgb(33, 150, 243);

        private readonly AppSettings settings;

        private ListBox lstPaths = null!;
        private ComboBox cmbLanguage = null!;
        private Button btnSave = null!;
        private Button btnCancel = null!;
        private Button btnAddPath = null!;
        private Button btnRemovePath = null!;

        public SettingsDialogNew(AppSettings settings)
        {
            this.settings = settings;

            InitializeComponent();
            LoadSettings();
        }

        private void InitializeComponent()
        {
            this.Text = Lang.Get("Settings");
            this.Size = new Size(550, 400);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.BackColor = DarkBackground;
            this.ForeColor = TextWhite;

            int y = 30;
            int padding = 20;

            // Macro Paths section
            var lblPaths = new Label
            {
                Text = Lang.Get("MacroPaths") + ":",
                Location = new Point(padding, y),
                Size = new Size(200, 25),
                ForeColor = TextWhite
            };
            y += 30;

            lstPaths = new ListBox
            {
                Location = new Point(padding, y),
                Size = new Size(380, 100),
                BackColor = Color.FromArgb(50, 50, 50),
                ForeColor = TextWhite,
                BorderStyle = BorderStyle.FixedSingle
            };

            btnAddPath = new Button
            {
                Text = "+",
                Location = new Point(padding + 390, y),
                Size = new Size(45, 36),
                FlatStyle = FlatStyle.Flat,
                BackColor = DarkPanel,
                ForeColor = TextWhite
            };
            btnAddPath.Click += BtnAddPath_Click;

            btnRemovePath = new Button
            {
                Text = "-",
                Location = new Point(padding + 390, y + 45),
                Size = new Size(45, 36),
                FlatStyle = FlatStyle.Flat,
                BackColor = DarkPanel,
                ForeColor = TextWhite
            };
            btnRemovePath.Click += BtnRemovePath_Click;

            y += 120;

            // Language Selection
            var lblLanguage = new Label
            {
                Text = Lang.Get("Language") + ":",
                Location = new Point(padding, y),
                Size = new Size(100, 25),
                ForeColor = TextWhite
            };

            cmbLanguage = new ComboBox
            {
                Location = new Point(padding + 110, y),
                Size = new Size(200, 30),
                DropDownStyle = ComboBoxStyle.DropDownList,
                BackColor = Color.FromArgb(50, 50, 50),
                ForeColor = TextWhite,
                FlatStyle = FlatStyle.Flat
            };

            // Add language options
            var languages = Lang.GetAvailableLanguages();
            foreach (var lang in languages)
            {
                cmbLanguage.Items.Add(new LanguageItem(lang.Key, lang.Value));
            }

            y += 70;

            // Buttons
            btnSave = new Button
            {
                Text = Lang.Get("Save"),
                Location = new Point(this.ClientSize.Width - 200, y),
                Size = new Size(80, 36),
                FlatStyle = FlatStyle.Flat,
                BackColor = AccentBlue,
                ForeColor = TextWhite
            };
            btnSave.Click += BtnSave_Click;

            btnCancel = new Button
            {
                Text = Lang.Get("Cancel"),
                Location = new Point(this.ClientSize.Width - 100, y),
                Size = new Size(80, 36),
                FlatStyle = FlatStyle.Flat,
                BackColor = DarkPanel,
                ForeColor = TextWhite
            };
            btnCancel.Click += (s, e) => { this.DialogResult = DialogResult.Cancel; this.Close(); };

            this.Controls.AddRange(new Control[]
            {
                lblPaths, lstPaths, btnAddPath, btnRemovePath,
                lblLanguage, cmbLanguage,
                btnSave, btnCancel
            });
        }

        private void LoadSettings()
        {
            lstPaths.Items.Clear();
            foreach (var path in settings.MacroPaths)
            {
                lstPaths.Items.Add(path);
            }

            // Select current language
            for (int i = 0; i < cmbLanguage.Items.Count; i++)
            {
                if (cmbLanguage.Items[i] is LanguageItem item && item.Code == settings.Language)
                {
                    cmbLanguage.SelectedIndex = i;
                    break;
                }
            }

            if (cmbLanguage.SelectedIndex < 0 && cmbLanguage.Items.Count > 0)
                cmbLanguage.SelectedIndex = 0;
        }

        private void BtnAddPath_Click(object? sender, System.EventArgs e)
        {
            using var dialog = new FolderBrowserDialog
            {
                Description = Lang.Get("SelectFolder")
            };

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                if (!lstPaths.Items.Contains(dialog.SelectedPath))
                {
                    lstPaths.Items.Add(dialog.SelectedPath);
                }
            }
        }

        private void BtnRemovePath_Click(object? sender, System.EventArgs e)
        {
            if (lstPaths.SelectedIndex >= 0)
            {
                // Don't allow removing the default path (first one)
                if (lstPaths.SelectedIndex == 0 && lstPaths.Items.Count > 0)
                {
                    var defaultPath = AppSettings.DefaultMacrosPath;
                    if (lstPaths.Items[0]?.ToString() == defaultPath)
                    {
                        MessageBox.Show(Lang.Get("CannotRemoveDefault"), Lang.Get("Confirm"));
                        return;
                    }
                }
                lstPaths.Items.RemoveAt(lstPaths.SelectedIndex);
            }
        }

        private void BtnSave_Click(object? sender, System.EventArgs e)
        {
            settings.MacroPaths.Clear();
            foreach (var item in lstPaths.Items)
            {
                if (item != null)
                    settings.MacroPaths.Add(item.ToString()!);
            }

            if (cmbLanguage.SelectedItem is LanguageItem langItem)
            {
                settings.Language = langItem.Code;
            }

            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        // Helper class for language dropdown
        private class LanguageItem
        {
            public string Code { get; }
            public string DisplayName { get; }

            public LanguageItem(string code, string displayName)
            {
                Code = code;
                DisplayName = displayName;
            }

            public override string ToString() => DisplayName;
        }
    }
}
