# Mini Solidwork C# Marco Automator

一個簡單的 C# 代碼執行器，用於測試和執行 SolidWorks 相關的 C# 腳本。**現支持自動連接到運行中的 SolidWorks！**

## 項目結構

```
MiniSw/
├── solidwork-api/                  # 包含所有必需的 SolidWorks DLL 文件
├── MainForm.cs                     # 主窗口界面實現
├── SolidWorksConnectionManager.cs  # SolidWorks 連接管理器
├── Program.cs                      # 程序入口點
├── MiniSolidworkAutomator.csproj   # 項目文件
├── app.manifest                    # 應用程序清單
└── README.md                       # 本文件
```

## 功能特點

- **SolidWorks 自動連接**：
  - 程序啟動時自動嘗試連接運行中的 SolidWorks
  - 使用 Running Object Table (ROT) 技術，與 Electron 應用相同的連接方式
  - 智能選擇最佳 SolidWorks 實例（優先選擇有可見窗口/活動文檔的）
  - 支持多 SolidWorks 實例環境
  - 每 5 秒自動刷新連接狀態

- **頂部連接狀態欄**：
  - 綠色指示燈：已連接
  - 黃色指示燈：連接中
  - 紅色指示燈：連接失敗
  - 灰色指示燈：未連接
  - 顯示 SolidWorks 版本和活動文檔

- **雙面板界面**：
  - 左側：C# 代碼編輯器（支持 F5 快捷鍵）
  - 右側：終端輸出顯示

- **腳本可用變量和函數**：
  - `swApp` (ISldWorks) - SolidWorks 應用程序對象
  - `swModel` (IModelDoc2) - 當前活動文檔
  - `Print(msg)` - 輸出綠色文字
  - `PrintError(msg)` - 輸出紅色錯誤文字
  - `PrintWarning(msg)` - 輸出橙色警告文字

## 編譯和運行

### 快速啟動（推薦）

**已經編譯好了！** 直接雙擊以下文件即可運行：
- `Start.bat` （Windows 命令提示符）
- `Start.ps1` （PowerShell）

或直接運行：
```
publish\MiniSolidworkAutomator.exe
```

### 方法 1：使用 Visual Studio

1. 在 Visual Studio 中打開 `MiniSolidworkAutomator.csproj`
2. 按 F5 或點擊"開始"按鈕編譯並運行

### 方法 2：使用命令行

```powershell
# 在 MiniSw 文件夾中執行
dotnet build
dotnet run
```

### 方法 3：重新發布為獨立 EXE

```powershell
# 發布為自包含的可執行文件
dotnet publish -c Release -r win-x64 --self-contained true -o ./publish

# 生成的 EXE 文件位於：./publish/MiniSolidworkAutomator.exe
```

## 使用示例

### 基本 SolidWorks 操作

```csharp
// 獲取 SolidWorks 版本
if (swApp != null)
{
    Print($"SolidWorks 版本: {swApp.RevisionNumber()}");
}

// 獲取活動文檔信息
if (swModel != null)
{
    Print($"活動文檔: {swModel.GetTitle()}");
    Print($"文檔路徑: {swModel.GetPathName()}");
}

// 遍歷所有打開的文檔
var docs = swApp.GetDocuments() as object[];
if (docs != null)
{
    Print($"打開的文檔 ({docs.Length} 個):");
    foreach (IModelDoc2 doc in docs)
    {
        Print($"  - {doc.GetTitle()}");
    }
}
```

### 多文件處理示例

```csharp
// 遍歷所有打開的零件文檔
var docs = swApp.GetDocuments() as object[];
foreach (IModelDoc2 doc in docs)
{
    var docType = (swDocumentTypes_e)doc.GetType();
    if (docType == swDocumentTypes_e.swDocPART)
    {
        Print($"處理零件: {doc.GetTitle()}");
        // 在這裡添加處理邏輯
    }
}
```

## 依賴項

項目已包含以下 SolidWorks 相關的 DLL：
- SolidWorks.Interop.sldworks.dll
- SolidWorks.Interop.swconst.dll
- CADBooster.SolidDna.dll

以及其他必需的依賴庫。

## 系統要求

- Windows 10 或更高版本
- SolidWorks 2018 或更高版本（用於自動化功能）
- .NET 8.0 運行時或 SDK

## 注意事項

- 代碼在沙盒環境中執行，某些系統級操作可能受限
- 執行長時間運行的代碼時，可以使用"停止"按鈕取消執行
- 關閉窗口時，如果有代碼正在執行，會提示確認

## 許可證

此項目為示例項目，僅供學習和測試使用。
