using System;
using System.Drawing;
using System.Windows.Forms;
using System.Threading;
using System.Threading.Tasks;
using System.Reflection;
using Microsoft.CodeAnalysis.CSharp.Scripting;
using Microsoft.CodeAnalysis.Scripting;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace MiniSolidworkAutomator
{
    /// <summary>
    /// Globals class that will be available to scripts.
    /// Scripts can access swApp directly.
    /// </summary>
    public class ScriptGlobals
    {
        public ISldWorks? swApp { get; set; }
        public IModelDoc2? swModel { get; set; }
        public Action<string>? Print { get; set; }
        public Action<string>? PrintError { get; set; }
        public Action<string>? PrintWarning { get; set; }
    }

    public partial class MainForm : Form
    {
        private SplitContainer mainSplitContainer = null!;
        private TextBox codeEditor = null!;
        private RichTextBox terminalDisplay = null!;
        private Button runButton = null!;
        private Button stopButton = null!;
        private Button connectButton = null!;
        private Button refreshButton = null!;
        private Panel statusPanel = null!;
        private Label statusLabel = null!;
        private Panel statusIndicator = null!;
        private Label connectionLabel = null!;
        private Panel connectionIndicator = null!;
        
        private CancellationTokenSource? cancellationTokenSource;
        private bool isRunning = false;
        
        // SolidWorks connection manager
        private SolidWorksConnectionManager swConnectionManager = new SolidWorksConnectionManager();
        private System.Windows.Forms.Timer? connectionCheckTimer;

        public MainForm()
        {
            InitializeComponent();
            SetupUI();
            SetupConnectionManager();
        }

        private void InitializeComponent()
        {
            this.Text = "Mini Solidwork C# Marco Automator";
            this.Size = new Size(1200, 700);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.MinimumSize = new Size(800, 500);
        }

        private void SetupConnectionManager()
        {
            swConnectionManager.ConnectionStateChanged += OnConnectionStateChanged;
            
            // Auto-connect on startup
            Task.Run(() =>
            {
                Thread.Sleep(500); // Small delay for UI to initialize
                this.Invoke(new Action(() =>
                {
                    AppendToTerminal("正在嘗試連接 SolidWorks...", Color.Cyan);
                }));
                swConnectionManager.Connect();
            });

            // Setup periodic connection check (every 5 seconds)
            connectionCheckTimer = new System.Windows.Forms.Timer();
            connectionCheckTimer.Interval = 5000;
            connectionCheckTimer.Tick += (s, e) =>
            {
                if (swConnectionManager.IsConnected)
                {
                    swConnectionManager.RefreshStatus();
                }
            };
            connectionCheckTimer.Start();
        }

        private void OnConnectionStateChanged(object? sender, ConnectionInfo info)
        {
            if (InvokeRequired)
            {
                Invoke(new Action(() => OnConnectionStateChanged(sender, info)));
                return;
            }

            UpdateConnectionUI(info);
        }

        private void UpdateConnectionUI(ConnectionInfo info)
        {
            switch (info.State)
            {
                case ConnectionState.Connected:
                    connectionIndicator.BackColor = Color.LimeGreen;
                    string docInfo = info.ActiveDocumentName != null 
                        ? $" | 文檔: {info.ActiveDocumentName}" 
                        : " | 無活動文檔";
                    connectionLabel.Text = $"已連接 SW{info.SolidWorksVersion?.Split('.')[0]}{docInfo}";
                    connectButton.Text = "重連";
                    AppendToTerminal($"✅ {info.Message}", Color.LimeGreen);
                    if (info.ActiveDocumentName != null)
                    {
                        AppendToTerminal($"   活動文檔: {info.ActiveDocumentName}", Color.White);
                    }
                    AppendToTerminal($"   打開文檔數: {info.OpenDocumentCount}", Color.White);
                    break;
                    
                case ConnectionState.Connecting:
                    connectionIndicator.BackColor = Color.Yellow;
                    connectionLabel.Text = "連接中...";
                    break;
                    
                case ConnectionState.Failed:
                    connectionIndicator.BackColor = Color.Red;
                    connectionLabel.Text = "連接失敗";
                    connectButton.Text = "連接";
                    AppendToTerminal($"❌ {info.Message}", Color.Red);
                    break;
                    
                case ConnectionState.Disconnected:
                    connectionIndicator.BackColor = Color.Gray;
                    connectionLabel.Text = "未連接";
                    connectButton.Text = "連接";
                    if (!string.IsNullOrEmpty(info.Message))
                    {
                        AppendToTerminal($"⚠️ {info.Message}", Color.Orange);
                    }
                    break;
            }
            connectionIndicator.Invalidate();
        }

        private void SetupUI()
        {
            // Top connection panel
            Panel topPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 50,
                BackColor = Color.FromArgb(30, 30, 30),
                BorderStyle = BorderStyle.FixedSingle
            };

            // Connection indicator
            connectionIndicator = new Panel
            {
                Location = new Point(15, 15),
                Size = new Size(20, 20),
                BackColor = Color.Gray
            };
            connectionIndicator.Paint += (s, e) =>
            {
                e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                e.Graphics.FillEllipse(new SolidBrush(connectionIndicator.BackColor), 0, 0, 20, 20);
            };

            // Connection label
            connectionLabel = new Label
            {
                Text = "未連接 SolidWorks",
                Location = new Point(45, 15),
                Size = new Size(350, 20),
                ForeColor = Color.White,
                Font = new Font("Microsoft YaHei UI", 9)
            };

            // Connect button
            connectButton = new Button
            {
                Text = "連接",
                Location = new Point(420, 10),
                Size = new Size(80, 30),
                BackColor = Color.FromArgb(0, 122, 204),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Microsoft YaHei UI", 9)
            };
            connectButton.FlatAppearance.BorderSize = 0;
            connectButton.Click += ConnectButton_Click;

            // Refresh button
            refreshButton = new Button
            {
                Text = "刷新",
                Location = new Point(510, 10),
                Size = new Size(80, 30),
                BackColor = Color.FromArgb(80, 80, 80),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Microsoft YaHei UI", 9)
            };
            refreshButton.FlatAppearance.BorderSize = 0;
            refreshButton.Click += RefreshButton_Click;

            topPanel.Controls.Add(connectionIndicator);
            topPanel.Controls.Add(connectionLabel);
            topPanel.Controls.Add(connectButton);
            topPanel.Controls.Add(refreshButton);

            // Main split container (left/right)
            mainSplitContainer = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Vertical,
                SplitterDistance = 580,
                BorderStyle = BorderStyle.FixedSingle
            };

            // Left panel - Code Editor
            codeEditor = new TextBox
            {
                Dock = DockStyle.Fill,
                Multiline = true,
                Font = new Font("Consolas", 10),
                ScrollBars = ScrollBars.Both,
                WordWrap = false,
                BackColor = Color.FromArgb(30, 30, 30),
                ForeColor = Color.White,
                Text = GetDefaultCode()
            };

            Label leftLabel = new Label
            {
                Text = "C# 代碼編輯器 (可使用 swApp, swModel, Print())",
                Dock = DockStyle.Top,
                Height = 30,
                TextAlign = ContentAlignment.MiddleLeft,
                BackColor = Color.FromArgb(45, 45, 48),
                ForeColor = Color.White,
                Padding = new Padding(10, 0, 0, 0)
            };

            mainSplitContainer.Panel1.Controls.Add(codeEditor);
            mainSplitContainer.Panel1.Controls.Add(leftLabel);

            // Right panel - Terminal Display
            terminalDisplay = new RichTextBox
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                Font = new Font("Consolas", 9),
                BackColor = Color.Black,
                ForeColor = Color.LimeGreen,
                Text = "終端輸出顯示區域\n準備就緒...\n"
            };

            Label rightLabel = new Label
            {
                Text = "終端輸出",
                Dock = DockStyle.Top,
                Height = 30,
                TextAlign = ContentAlignment.MiddleLeft,
                BackColor = Color.FromArgb(45, 45, 48),
                ForeColor = Color.White,
                Padding = new Padding(10, 0, 0, 0)
            };

            mainSplitContainer.Panel2.Controls.Add(terminalDisplay);
            mainSplitContainer.Panel2.Controls.Add(rightLabel);

            // Bottom control panel
            Panel bottomPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 60,
                BackColor = Color.FromArgb(240, 240, 240),
                BorderStyle = BorderStyle.FixedSingle
            };

            // Run button
            runButton = new Button
            {
                Text = "運行 (F5)",
                Location = new Point(20, 15),
                Size = new Size(120, 35),
                BackColor = Color.FromArgb(0, 122, 204),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Microsoft YaHei UI", 9, FontStyle.Bold)
            };
            runButton.FlatAppearance.BorderSize = 0;
            runButton.Click += RunButton_Click;

            // Stop button
            stopButton = new Button
            {
                Text = "停止 (Stop)",
                Location = new Point(150, 15),
                Size = new Size(120, 35),
                BackColor = Color.FromArgb(200, 50, 50),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Microsoft YaHei UI", 9, FontStyle.Bold),
                Enabled = false
            };
            stopButton.FlatAppearance.BorderSize = 0;
            stopButton.Click += StopButton_Click;

            // Status panel
            statusPanel = new Panel
            {
                Location = new Point(300, 15),
                Size = new Size(250, 35),
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = Color.White
            };

            // Status indicator (circle)
            statusIndicator = new Panel
            {
                Location = new Point(10, 7),
                Size = new Size(20, 20),
                BackColor = Color.Gray
            };
            statusIndicator.Paint += (s, e) =>
            {
                e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                e.Graphics.FillEllipse(new SolidBrush(statusIndicator.BackColor), 0, 0, 20, 20);
            };

            // Status label
            statusLabel = new Label
            {
                Text = "已停止 (Stopped)",
                Location = new Point(40, 7),
                Size = new Size(200, 20),
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Microsoft YaHei UI", 9)
            };

            statusPanel.Controls.Add(statusIndicator);
            statusPanel.Controls.Add(statusLabel);

            bottomPanel.Controls.Add(runButton);
            bottomPanel.Controls.Add(stopButton);
            bottomPanel.Controls.Add(statusPanel);

            // Add controls to form (order matters for docking)
            this.Controls.Add(mainSplitContainer);
            this.Controls.Add(bottomPanel);
            this.Controls.Add(topPanel);

            // Add keyboard shortcut
            this.KeyPreview = true;
            this.KeyDown += MainForm_KeyDown;
        }

        private string GetDefaultCode()
        {
            return @"// ========================================
// SolidWorks 自動化腳本
// 可用變量：swApp (ISldWorks), swModel (IModelDoc2)
// 可用函數：Print(), PrintError(), PrintWarning()
// ========================================

// 示例 1: 獲取 SolidWorks 版本
if (swApp != null)
{
    Print($""SolidWorks 版本: {swApp.RevisionNumber()}"");
}
else
{
    PrintError(""swApp 為 null - 請先連接 SolidWorks"");
}

// 示例 2: 獲取活動文檔信息
if (swModel != null)
{
    Print($""活動文檔: {swModel.GetTitle()}"");
    Print($""文檔路徑: {swModel.GetPathName()}"");
}
else
{
    PrintWarning(""沒有活動文檔"");
}

// 示例 3: 遍歷所有打開的文檔
if (swApp != null)
{
    var docs = swApp.GetDocuments() as object[];
    if (docs != null && docs.Length > 0)
    {
        Print($""\n打開的文檔 ({docs.Length} 個):"");
        foreach (IModelDoc2 doc in docs)
        {
            Print($""  - {doc.GetTitle()}"");
        }
    }
}
";
        }

        private void MainForm_KeyDown(object? sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F5 && !isRunning)
            {
                RunButton_Click(sender, e);
                e.Handled = true;
            }
        }

        private void ConnectButton_Click(object? sender, EventArgs e)
        {
            AppendToTerminal("\n正在連接 SolidWorks...", Color.Cyan);
            Task.Run(() => swConnectionManager.Connect());
        }

        private void RefreshButton_Click(object? sender, EventArgs e)
        {
            swConnectionManager.RefreshStatus();
            var info = swConnectionManager.ConnectionInfo;
            if (info.State == ConnectionState.Connected)
            {
                AppendToTerminal($"🔄 連接狀態已刷新", Color.Cyan);
                if (info.ActiveDocumentName != null)
                {
                    AppendToTerminal($"   活動文檔: {info.ActiveDocumentName}", Color.White);
                }
            }
        }

        private void UpdateStatus(bool running)
        {
            if (InvokeRequired)
            {
                Invoke(new Action(() => UpdateStatus(running)));
                return;
            }

            isRunning = running;
            if (running)
            {
                statusIndicator.BackColor = Color.LimeGreen;
                statusLabel.Text = "運行中 (Running)";
                runButton.Enabled = false;
                stopButton.Enabled = true;
            }
            else
            {
                statusIndicator.BackColor = Color.Gray;
                statusLabel.Text = "已停止 (Stopped)";
                runButton.Enabled = true;
                stopButton.Enabled = false;
            }
            statusIndicator.Invalidate();
        }

        private void AppendToTerminal(string text, Color? color = null)
        {
            if (InvokeRequired)
            {
                Invoke(new Action(() => AppendToTerminal(text, color)));
                return;
            }

            terminalDisplay.SelectionStart = terminalDisplay.TextLength;
            terminalDisplay.SelectionLength = 0;
            terminalDisplay.SelectionColor = color ?? Color.LimeGreen;
            terminalDisplay.AppendText(text + System.Environment.NewLine);
            terminalDisplay.SelectionColor = terminalDisplay.ForeColor;
            terminalDisplay.ScrollToCaret();
        }

        private async void RunButton_Click(object? sender, EventArgs e)
        {
            string code = codeEditor.Text.Trim();
            
            if (string.IsNullOrWhiteSpace(code))
            {
                MessageBox.Show("請輸入要執行的 C# 代碼！", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            UpdateStatus(true);
            terminalDisplay.Clear();
            AppendToTerminal("========================================");
            AppendToTerminal($"開始執行 [{DateTime.Now:HH:mm:ss}]", Color.Cyan);
            AppendToTerminal("========================================");
            
            // Show connection status
            if (swConnectionManager.IsConnected)
            {
                AppendToTerminal($"✅ SolidWorks 已連接", Color.LimeGreen);
            }
            else
            {
                AppendToTerminal($"⚠️ SolidWorks 未連接 - swApp 將為 null", Color.Orange);
            }
            AppendToTerminal("");

            cancellationTokenSource = new CancellationTokenSource();

            try
            {
                // Prepare globals for script
                var globals = new ScriptGlobals
                {
                    swApp = swConnectionManager.SwApp,
                    swModel = swConnectionManager.SwApp?.IActiveDoc2,
                    Print = (msg) => AppendToTerminal(msg, Color.LimeGreen),
                    PrintError = (msg) => AppendToTerminal($"❌ {msg}", Color.Red),
                    PrintWarning = (msg) => AppendToTerminal($"⚠️ {msg}", Color.Orange)
                };

                // Get assembly references for SolidWorks interop
                var swInteropAssembly = typeof(ISldWorks).Assembly;
                var swConstAssembly = typeof(swDocumentTypes_e).Assembly;

                // Setup script options with SolidWorks references
                var options = ScriptOptions.Default
                    .WithReferences(
                        typeof(object).Assembly,                    // System
                        typeof(System.Linq.Enumerable).Assembly,    // System.Linq
                        typeof(System.Collections.Generic.List<>).Assembly,  // Collections
                        swInteropAssembly,                          // SolidWorks.Interop.sldworks
                        swConstAssembly                             // SolidWorks.Interop.swconst
                    )
                    .WithImports(
                        "System",
                        "System.Math",
                        "System.Collections.Generic",
                        "System.Linq",
                        "SolidWorks.Interop.sldworks",
                        "SolidWorks.Interop.swconst"
                    );

                await Task.Run(async () =>
                {
                    try
                    {
                        var result = await CSharpScript.EvaluateAsync(
                            code, 
                            options, 
                            globals: globals,
                            globalsType: typeof(ScriptGlobals),
                            cancellationToken: cancellationTokenSource.Token
                        );

                        // Display result if not null
                        if (result != null)
                        {
                            AppendToTerminal($"\n返回值: {result}", Color.Yellow);
                        }
                    }
                    catch (CompilationErrorException ex)
                    {
                        AppendToTerminal($"編譯錯誤:", Color.Red);
                        AppendToTerminal(string.Join(System.Environment.NewLine, ex.Diagnostics), Color.Red);
                    }
                    catch (Exception ex)
                    {
                        AppendToTerminal($"運行時錯誤: {ex.Message}", Color.Red);
                        if (ex.InnerException != null)
                        {
                            AppendToTerminal($"內部錯誤: {ex.InnerException.Message}", Color.Red);
                        }
                    }
                }, cancellationTokenSource.Token);

                AppendToTerminal("");
                AppendToTerminal("========================================");
                AppendToTerminal($"執行完成 [{DateTime.Now:HH:mm:ss}]", Color.Cyan);
                AppendToTerminal("========================================");
            }
            catch (OperationCanceledException)
            {
                AppendToTerminal("\n執行已被用戶取消！", Color.Orange);
            }
            catch (Exception ex)
            {
                AppendToTerminal($"\n發生錯誤: {ex.Message}", Color.Red);
            }
            finally
            {
                UpdateStatus(false);
                cancellationTokenSource?.Dispose();
                cancellationTokenSource = null;
            }
        }

        private void StopButton_Click(object? sender, EventArgs e)
        {
            if (cancellationTokenSource != null && !cancellationTokenSource.IsCancellationRequested)
            {
                AppendToTerminal("\n正在停止執行...", Color.Orange);
                cancellationTokenSource.Cancel();
            }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (isRunning)
            {
                var result = MessageBox.Show(
                    "代碼正在執行中，確定要關閉嗎？",
                    "確認關閉",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question
                );

                if (result == DialogResult.No)
                {
                    e.Cancel = true;
                    return;
                }

                cancellationTokenSource?.Cancel();
            }

            // Cleanup
            connectionCheckTimer?.Stop();
            connectionCheckTimer?.Dispose();
            swConnectionManager.Disconnect();

            base.OnFormClosing(e);
        }
    }
}
