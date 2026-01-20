using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using SolidWorks.Interop.sldworks;

namespace MiniSolidworkAutomator
{
    public enum ConnectionState
    {
        Disconnected,
        Connecting,
        Connected,
        Failed
    }

    public class ConnectionInfo
    {
        public ConnectionState State { get; set; } = ConnectionState.Disconnected;
        public string? Message { get; set; }
        public string? SolidWorksVersion { get; set; }
        public int? ProcessId { get; set; }
        public string? ActiveDocumentName { get; set; }
        public int OpenDocumentCount { get; set; }
    }

    /// <summary>
    /// Manages connection to SolidWorks application.
    /// Uses Running Object Table (ROT) to find and connect to running SolidWorks instances.
    /// Reference: Based on MecAgent Electron app's SolidworksConnexionController.cs
    /// </summary>
    public class SolidWorksConnectionManager
    {
        private ISldWorks? _swApp;
        private ConnectionInfo _connectionInfo = new ConnectionInfo();

        public event EventHandler<ConnectionInfo>? ConnectionStateChanged;

        public ISldWorks? SwApp => _swApp;
        public ConnectionInfo ConnectionInfo => _connectionInfo;
        public bool IsConnected => _connectionInfo.State == ConnectionState.Connected && _swApp != null;

        // DllImport for ROT (Running Object Table) connection
        [DllImport("ole32.dll")]
        private static extern int CreateBindCtx(uint reserved, out IBindCtx ppbc);

        /// <summary>
        /// Attempts to connect to a running SolidWorks instance.
        /// Priority: Visible window with active document > Visible window > Any running instance
        /// </summary>
        public bool Connect()
        {
            _connectionInfo = new ConnectionInfo { State = ConnectionState.Connecting, Message = "正在連接 SolidWorks..." };
            OnConnectionStateChanged();

            try
            {
                // First check if any SolidWorks processes are running
                var swProcesses = Process.GetProcessesByName("SLDWORKS");
                if (swProcesses.Length == 0)
                {
                    _connectionInfo = new ConnectionInfo
                    {
                        State = ConnectionState.Failed,
                        Message = "未找到 SolidWorks 進程。請先啟動 SolidWorks。"
                    };
                    OnConnectionStateChanged();
                    return false;
                }

                Console.WriteLine($"[SW連接] 找到 {swProcesses.Length} 個 SolidWorks 進程");

                // Try to connect via ROT (Running Object Table)
                _swApp = TryConnectViaROT(swProcesses);

                if (_swApp != null)
                {
                    // Verify connection and get info
                    string version = _swApp.RevisionNumber();
                    
                    // Get active document info
                    string? activeDocName = null;
                    int docCount = 0;
                    try
                    {
                        var activeDoc = _swApp.IActiveDoc2;
                        if (activeDoc != null)
                        {
                            string path = activeDoc.GetPathName();
                            activeDocName = string.IsNullOrEmpty(path) 
                                ? "未保存的文檔" 
                                : System.IO.Path.GetFileName(path);
                        }
                        
                        // Count open documents
                        var docs = _swApp.GetDocuments() as object[];
                        docCount = docs?.Length ?? 0;
                    }
                    catch (Exception docEx)
                    {
                        Console.WriteLine($"[SW連接] 無法獲取文檔信息: {docEx.Message}");
                    }

                    _connectionInfo = new ConnectionInfo
                    {
                        State = ConnectionState.Connected,
                        Message = $"已連接到 SolidWorks {version}",
                        SolidWorksVersion = version,
                        ActiveDocumentName = activeDocName,
                        OpenDocumentCount = docCount
                    };
                    OnConnectionStateChanged();
                    
                    Console.WriteLine($"[SW連接] ✅ 成功連接到 SolidWorks {version}");
                    Console.WriteLine($"[SW連接] 活動文檔: {activeDocName ?? "無"}, 打開文檔數: {docCount}");
                    
                    return true;
                }
                else
                {
                    _connectionInfo = new ConnectionInfo
                    {
                        State = ConnectionState.Failed,
                        Message = "無法連接到 SolidWorks。請確保 SolidWorks 已完全啟動。"
                    };
                    OnConnectionStateChanged();
                    return false;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[SW連接] ❌ 連接失敗: {ex.Message}");
                _connectionInfo = new ConnectionInfo
                {
                    State = ConnectionState.Failed,
                    Message = $"連接錯誤: {ex.Message}"
                };
                OnConnectionStateChanged();
                return false;
            }
        }

        /// <summary>
        /// Disconnect from SolidWorks
        /// </summary>
        public void Disconnect()
        {
            if (_swApp != null)
            {
                try
                {
                    Marshal.ReleaseComObject(_swApp);
                }
                catch { }
                _swApp = null;
            }

            _connectionInfo = new ConnectionInfo
            {
                State = ConnectionState.Disconnected,
                Message = "已斷開連接"
            };
            OnConnectionStateChanged();
        }

        /// <summary>
        /// Refresh connection info (check if still connected, update document info)
        /// </summary>
        public void RefreshStatus()
        {
            if (_swApp == null)
            {
                if (_connectionInfo.State == ConnectionState.Connected)
                {
                    _connectionInfo = new ConnectionInfo
                    {
                        State = ConnectionState.Disconnected,
                        Message = "連接已丟失"
                    };
                    OnConnectionStateChanged();
                }
                return;
            }

            try
            {
                // Test if connection is still valid
                string version = _swApp.RevisionNumber();
                
                // Update document info
                string? activeDocName = null;
                int docCount = 0;
                
                var activeDoc = _swApp.IActiveDoc2;
                if (activeDoc != null)
                {
                    string path = activeDoc.GetPathName();
                    activeDocName = string.IsNullOrEmpty(path) 
                        ? "未保存的文檔" 
                        : System.IO.Path.GetFileName(path);
                }
                
                var docs = _swApp.GetDocuments() as object[];
                docCount = docs?.Length ?? 0;

                _connectionInfo = new ConnectionInfo
                {
                    State = ConnectionState.Connected,
                    Message = $"已連接到 SolidWorks {version}",
                    SolidWorksVersion = version,
                    ActiveDocumentName = activeDocName,
                    OpenDocumentCount = docCount
                };
                OnConnectionStateChanged();
            }
            catch (COMException)
            {
                // Connection lost
                HandleConnectionLost("SolidWorks 可能已關閉");
            }
            catch (InvalidCastException)
            {
                // COM object became invalid (SolidWorks closed or connection stale)
                HandleConnectionLost("COM 連接已失效");
            }
            catch (Exception ex)
            {
                // Any other exception - treat as connection lost
                Console.WriteLine($"[RefreshStatus] 未預期的錯誤: {ex.Message}");
                HandleConnectionLost("連接錯誤");
            }
        }

        /// <summary>
        /// Helper method to handle connection lost scenarios
        /// </summary>
        private void HandleConnectionLost(string reason)
        {
            if (_swApp != null)
            {
                try
                {
                    Marshal.ReleaseComObject(_swApp);
                }
                catch { }
                _swApp = null;
            }
            
            _connectionInfo = new ConnectionInfo
            {
                State = ConnectionState.Disconnected,
                Message = $"連接已丟失 ({reason})"
            };
            OnConnectionStateChanged();
        }

        /// <summary>
        /// Gets SolidWorks application from Running Object Table by process ID
        /// </summary>
        private ISldWorks? GetSwAppFromProcess(int processId)
        {
            var monikerName = "SolidWorks_PID_" + processId.ToString();
            IBindCtx? context = null;
            IRunningObjectTable? rot = null;
            IEnumMoniker? monikers = null;

            try
            {
                CreateBindCtx(0, out context);
                context.GetRunningObjectTable(out rot);
                if (rot == null) return null;

                rot.EnumRunning(out monikers);
                if (monikers == null) return null;
                var moniker = new IMoniker[1];

                while (monikers.Next(1, moniker, IntPtr.Zero) == 0)
                {
                    var curMoniker = moniker.FirstOrDefault();
                    string? name = null;

                    if (curMoniker != null)
                    {
                        try
                        {
                            curMoniker.GetDisplayName(context, null, out name);
                        }
                        catch (UnauthorizedAccessException) { }
                    }

                    if (string.Equals(monikerName, name, StringComparison.CurrentCultureIgnoreCase) && curMoniker != null)
                    {
                        rot.GetObject(curMoniker, out object app);
                        return app as ISldWorks;
                    }
                }
            }
            finally
            {
                if (monikers != null) Marshal.ReleaseComObject(monikers);
                if (rot != null) Marshal.ReleaseComObject(rot);
                if (context != null) Marshal.ReleaseComObject(context);
            }

            return null;
        }

        /// <summary>
        /// Tries to connect to SolidWorks via ROT by finding the best instance.
        /// Priority: Visible window with active doc > Visible window > Most recent process
        /// </summary>
        private ISldWorks? TryConnectViaROT(Process[] swProcesses)
        {
            Console.WriteLine($"[ROT] 開始搜索 {swProcesses.Length} 個 SolidWorks 進程...");

            // Collect process info with scoring
            var processInfoList = new List<(Process Process, int Score, string WindowTitle, bool HasWindow)>();

            foreach (var proc in swProcesses)
            {
                try
                {
                    if (!proc.Responding)
                    {
                        Console.WriteLine($"[ROT] PID {proc.Id} - 跳過 (無響應)");
                        continue;
                    }

                    string windowTitle = string.Empty;
                    bool hasWindow = false;
                    try
                    {
                        windowTitle = proc.MainWindowTitle ?? string.Empty;
                        hasWindow = !string.IsNullOrEmpty(windowTitle);
                    }
                    catch { }

                    int score = CalculateProcessScore(proc, hasWindow);
                    processInfoList.Add((proc, score, windowTitle, hasWindow));

                    Console.WriteLine($"[ROT] PID {proc.Id} - 分數: {score}, 窗口: '{windowTitle}'");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"[ROT] PID {proc.Id} - 評估錯誤: {ex.Message}");
                }
            }

            if (processInfoList.Count == 0)
            {
                Console.WriteLine("[ROT] 未找到有效的 SolidWorks 進程");
                return null;
            }

            // Sort by score descending
            var sortedProcesses = processInfoList
                .OrderByDescending(p => p.Score)
                .ThenByDescending(p =>
                {
                    try { return p.Process.StartTime; }
                    catch { return DateTime.MinValue; }
                })
                .ToArray();

            // Try to connect in priority order
            ISldWorks? fallbackApp = null;
            
            foreach (var (proc, score, windowTitle, hasWindow) in sortedProcesses)
            {
                try
                {
                    Console.WriteLine($"[ROT] >>> 嘗試連接 PID {proc.Id} (分數: {score})...");

                    var swApp = GetSwAppFromProcess(proc.Id);

                    if (swApp != null)
                    {
                        string version = swApp.RevisionNumber();
                        Console.WriteLine($"[ROT] 連接成功 PID {proc.Id}, 版本: {version}");

                        // Check for active document
                        bool hasActiveDoc = false;
                        try
                        {
                            var activeDoc = swApp.IActiveDoc2;
                            hasActiveDoc = activeDoc != null;
                        }
                        catch { }

                        // Prefer instances with visible window or active document
                        if (hasWindow || hasActiveDoc)
                        {
                            Console.WriteLine($"[ROT] ✅ 選擇 PID {proc.Id} ({(hasWindow ? "有可見窗口" : "有活動文檔")})");
                            _connectionInfo.ProcessId = proc.Id;
                            return swApp;
                        }
                        else
                        {
                            // Keep as fallback
                            Console.WriteLine($"[ROT] ⚠️ PID {proc.Id} 無窗口/文檔，繼續搜索...");
                            fallbackApp ??= swApp;
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"[ROT] PID {proc.Id} - 連接失敗: {ex.Message}");
                }
            }

            // Return fallback if no better option
            if (fallbackApp != null)
            {
                Console.WriteLine("[ROT] ⚠️ 使用備用連接");
                return fallbackApp;
            }

            Console.WriteLine("[ROT] ❌ 所有連接嘗試失敗");
            return null;
        }

        /// <summary>
        /// Calculate priority score for a process
        /// </summary>
        private int CalculateProcessScore(Process process, bool hasWindow)
        {
            int score = 0;

            // Factor 1: Visible window (+100)
            if (hasWindow)
            {
                score += 100;
            }

            // Factor 2: Recency (0-50 points)
            try
            {
                var allSwProcesses = Process.GetProcessesByName("SLDWORKS");
                if (allSwProcesses.Length > 1)
                {
                    var sortedByTime = allSwProcesses
                        .OrderByDescending(p =>
                        {
                            try { return p.StartTime; }
                            catch { return DateTime.MinValue; }
                        })
                        .ToArray();

                    int index = Array.FindIndex(sortedByTime, p => p.Id == process.Id);
                    if (index >= 0)
                    {
                        score += Math.Max(0, 50 - (index * 10));
                    }
                }
                else
                {
                    score += 50;
                }
            }
            catch { }

            return score;
        }

        protected virtual void OnConnectionStateChanged()
        {
            ConnectionStateChanged?.Invoke(this, _connectionInfo);
        }
    }
}
