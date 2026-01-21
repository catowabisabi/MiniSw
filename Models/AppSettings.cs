using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;

namespace MiniSolidworkAutomator.Models
{
    public class MacroBookmark
    {
        public string Id { get; set; } = Guid.NewGuid().ToString();
        public string Name { get; set; } = "新書籤";
        public string FilePath { get; set; } = "";
        public string Content { get; set; } = "";
        public MacroType Type { get; set; } = MacroType.CSharp;
        public bool IsUnsaved { get; set; } = false;
        [System.Text.Json.Serialization.JsonIgnore]
        public object? Tag { get; set; } = null;  // For storing original content to track changes
    }

    public enum MacroType
    {
        CSharp,
        VBA
    }

    public class PromptItem
    {
        public string Id { get; set; } = Guid.NewGuid().ToString();
        public string Name { get; set; } = "新提示";
        public string Content { get; set; } = "";
        public bool IsDefault { get; set; } = false;
    }

    public class NoteItem
    {
        public string Id { get; set; } = Guid.NewGuid().ToString();
        public string Name { get; set; } = "新筆記";
        public string Content { get; set; } = "";
        public DateTime CreatedAt { get; set; } = DateTime.Now;
        public DateTime ModifiedAt { get; set; } = DateTime.Now;
    }

    public class RecentFileItem
    {
        public string FilePath { get; set; } = "";
        public string Name { get; set; } = "";
        public DateTime LastOpened { get; set; } = DateTime.Now;
        public MacroType Type { get; set; } = MacroType.CSharp;
    }

    public class RecentEntryPoint
    {
        public string ModuleName { get; set; } = "";
        public string ProcedureName { get; set; } = "";
        public DateTime LastUsed { get; set; } = DateTime.Now;
        public string FilePath { get; set; } = ""; // Optional: associate with specific file
    }

    public class ExecutionHistoryItem
    {
        public string FileName { get; set; } = "";
        public string FilePath { get; set; } = "";
        public DateTime ExecutionTime { get; set; } = DateTime.Now;
        public string Language { get; set; } = "C#";
        public bool Success { get; set; } = true;
        public long DurationMs { get; set; } = 0;
        public string ErrorMessage { get; set; } = "";
        public string UsedModuleName { get; set; } = "";
        public string UsedProcedureName { get; set; } = "";
        
        // For backward compatibility
        public DateTime ExecutedAt 
        { 
            get => ExecutionTime; 
            set => ExecutionTime = value; 
        }
    }

    public class AppSettings
    {
        public List<string> MacroPaths { get; set; } = new List<string>();
        public List<PromptItem> Prompts { get; set; } = new List<PromptItem>();
        public List<NoteItem> Notes { get; set; } = new List<NoteItem>();
        public List<RecentFileItem> RecentFiles { get; set; } = new List<RecentFileItem>();
        public List<ExecutionHistoryItem> ExecutionHistory { get; set; } = new List<ExecutionHistoryItem>();
        public List<RecentEntryPoint> RecentEntryPoints { get; set; } = new List<RecentEntryPoint>();
        public string LastOpenedFile { get; set; } = "";
        public int WindowWidth { get; set; } = 1400;
        public int WindowHeight { get; set; } = 850;
        public string Language { get; set; } = "zh-TW";
        public bool AutoSaveEnabled { get; set; } = true;
        public int AutoSaveIntervalSeconds { get; set; } = 60;
        public int MaxRecentFiles { get; set; } = 15;
        public int MaxExecutionHistory { get; set; } = 50;
        public int MaxRecentEntryPoints { get; set; } = 20;

        private static string SettingsPath => Path.Combine(
            AppDomain.CurrentDomain.BaseDirectory, 
            "settings.json"
        );

        public static string DefaultMacrosPath => Path.Combine(
            AppDomain.CurrentDomain.BaseDirectory,
            "My SW Macros"
        );

        public void AddRecentFile(string filePath, string name, MacroType type)
        {
            // Remove if already exists
            RecentFiles.RemoveAll(f => f.FilePath.Equals(filePath, StringComparison.OrdinalIgnoreCase));
            
            // Add to front
            RecentFiles.Insert(0, new RecentFileItem
            {
                FilePath = filePath,
                Name = name,
                Type = type,
                LastOpened = DateTime.Now
            });

            // Trim to max
            if (RecentFiles.Count > MaxRecentFiles)
                RecentFiles = RecentFiles.Take(MaxRecentFiles).ToList();
        }

        public void AddExecutionHistory(string fileName, string filePath, bool success, long durationMs, string errorMessage = "")
        {
            ExecutionHistory.Insert(0, new ExecutionHistoryItem
            {
                FileName = fileName,
                FilePath = filePath,
                ExecutedAt = DateTime.Now,
                Success = success,
                DurationMs = durationMs,
                ErrorMessage = errorMessage
            });

            // Trim to max
            if (ExecutionHistory.Count > MaxExecutionHistory)
                ExecutionHistory = ExecutionHistory.Take(MaxExecutionHistory).ToList();
        }
        
        public void AddRecentEntryPoint(string moduleName, string procedureName, string filePath = "")
        {
            // Remove existing entry if it exists
            var existing = RecentEntryPoints.FirstOrDefault(e => 
                e.ModuleName.Equals(moduleName, StringComparison.OrdinalIgnoreCase) && 
                e.ProcedureName.Equals(procedureName, StringComparison.OrdinalIgnoreCase));
            
            if (existing != null)
            {
                RecentEntryPoints.Remove(existing);
            }
            
            // Add new entry at the beginning
            RecentEntryPoints.Insert(0, new RecentEntryPoint
            {
                ModuleName = moduleName,
                ProcedureName = procedureName,
                FilePath = filePath,
                LastUsed = DateTime.Now
            });
            
            // Trim to max size
            if (RecentEntryPoints.Count > MaxRecentEntryPoints)
            {
                RecentEntryPoints.RemoveRange(MaxRecentEntryPoints, RecentEntryPoints.Count - MaxRecentEntryPoints);
            }
        }

        public static AppSettings Load()
        {
            try
            {
                if (File.Exists(SettingsPath))
                {
                    string json = File.ReadAllText(SettingsPath);
                    var settings = JsonSerializer.Deserialize<AppSettings>(json) ?? new AppSettings();
                    settings.EnsureDefaults();
                    return settings;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"載入設定失敗: {ex.Message}");
            }

            var newSettings = new AppSettings();
            newSettings.EnsureDefaults();
            return newSettings;
        }

        public void Save()
        {
            try
            {
                var options = new JsonSerializerOptions { WriteIndented = true };
                string json = JsonSerializer.Serialize(this, options);
                File.WriteAllText(SettingsPath, json);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"保存設定失敗: {ex.Message}");
            }
        }

        private void EnsureDefaults()
        {
            // Ensure default paths
            if (!MacroPaths.Contains(DefaultMacrosPath))
            {
                MacroPaths.Insert(0, DefaultMacrosPath);
            }

            // Ensure default prompts exist
            if (Prompts.Count == 0 || !Prompts.Exists(p => p.IsDefault))
            {
                AddDefaultPrompts();
            }
        }

        private void AddDefaultPrompts()
        {
            Prompts.Insert(0, new PromptItem
            {
                Name = "C# SolidWorks 宏生成",
                IsDefault = true,
                Content = @"請幫我生成 SolidWorks C# 宏代碼。

【執行環境說明】
- 這是一個 C# 腳本環境，不需要 class 包裝
- 已有變量: swApp (ISldWorks), swModel (IModelDoc2)
- 已導入命名空間: System, System.IO, System.Linq, System.Collections.Generic
- 已導入命名空間: SolidWorks.Interop.sldworks, SolidWorks.Interop.swconst
- 輸出函數: Print(msg), PrintError(msg), PrintWarning(msg)

【代碼要求】
1. 直接寫腳本代碼，不需要 class 或 namespace
2. 使用 Print() 輸出信息，不要用 Console.WriteLine
3. 添加空值檢查和錯誤處理
4. 添加中文註釋說明代碼功能

【常用 API 範例】

// 獲取活動文檔
if (swModel == null) { PrintError(""請先打開文檔""); return; }

// 判斷文檔類型
int docType = swModel.GetType();
if (docType == (int)swDocumentTypes_e.swDocPART) { Print(""這是零件""); }
if (docType == (int)swDocumentTypes_e.swDocASSEMBLY) { Print(""這是裝配體""); }
if (docType == (int)swDocumentTypes_e.swDocDRAWING) { Print(""這是工程圖""); }

// 獲取文檔路徑
string path = swModel.GetPathName();
string dir = Path.GetDirectoryName(path);
string name = Path.GetFileNameWithoutExtension(path);

// 遍歷零件中的所有實體
IPartDoc swPart = swModel as IPartDoc;
object[] bodies = swPart.GetBodies2((int)swBodyType_e.swSolidBody, true) as object[];

// 遍歷裝配體中的組件
IAssemblyDoc swAssy = swModel as IAssemblyDoc;
object[] comps = swAssy.GetComponents(true) as object[];
foreach (IComponent2 comp in comps) { Print(comp.Name2); }

// 選擇管理器
ISelectionMgr selMgr = swModel.ISelectionManager;
int count = selMgr.GetSelectedObjectCount2(-1);

// 自定義屬性
IModelDocExtension ext = swModel.Extension;
ICustomPropertyManager propMgr = ext.get_CustomPropertyManager("""");
propMgr.Add3(""屬性名"", (int)swCustomInfoType_e.swCustomInfoText, ""值"", 0);

// 保存文檔
swModel.Save3((int)swSaveAsOptions_e.swSaveAsOptions_Silent, ref errors, ref warnings);

【功能需求】
[在此描述你需要的功能]"
            });

            Prompts.Insert(1, new PromptItem
            {
                Name = "VBA SolidWorks 宏生成",
                IsDefault = true,
                Content = @"請幫我生成 SolidWorks VBA 宏代碼。

【代碼要求】
1. 使用標準 VBA 語法
2. 包含 Sub main() 作為入口點
3. 添加適當的錯誤處理 (On Error GoTo)
4. 添加中文註釋說明代碼功能

【常用 API 範例】

' 基本結構
Sub main()
    Dim swApp As SldWorks.SldWorks
    Dim swModel As SldWorks.ModelDoc2
    
    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc
    
    If swModel Is Nothing Then
        MsgBox ""請先打開文檔""
        Exit Sub
    End If
    
    ' 你的代碼...
End Sub

' 判斷文檔類型
Select Case swModel.GetType
    Case swDocPART: Debug.Print ""零件""
    Case swDocASSEMBLY: Debug.Print ""裝配體""
    Case swDocDRAWING: Debug.Print ""工程圖""
End Select

' 遍歷特徵
Dim swFeat As SldWorks.Feature
Set swFeat = swModel.FirstFeature
Do While Not swFeat Is Nothing
    Debug.Print swFeat.Name
    Set swFeat = swFeat.GetNextFeature
Loop

' 遍歷裝配體組件
Dim swAssy As SldWorks.AssemblyDoc
Dim vComps As Variant
Set swAssy = swModel
vComps = swAssy.GetComponents(True)
Dim i As Long
For i = 0 To UBound(vComps)
    Dim swComp As SldWorks.Component2
    Set swComp = vComps(i)
    Debug.Print swComp.Name2
Next i

' 自定義屬性
Dim swCustPropMgr As SldWorks.CustomPropertyManager
Set swCustPropMgr = swModel.Extension.CustomPropertyManager("""")
swCustPropMgr.Add3 ""屬性名"", swCustomInfoText, ""值"", swCustomPropertyAddOption_OnlyIfNew

【功能需求】
[在此描述你需要的功能]"
            });

            Prompts.Insert(2, new PromptItem
            {
                Name = "批量處理文件",
                IsDefault = true,
                Content = @"請幫我生成批量處理 SolidWorks 文件的代碼。

【任務描述】
我需要對指定文件夾中的多個 SolidWorks 文件進行批量操作。

【處理要求】
- 文件夾路徑: [輸入路徑，例如: D:\Parts]
- 文件類型: [SLDPRT / SLDASM / SLDDRW / 全部]
- 是否包含子文件夾: [是/否]
- 操作內容: [描述需要對每個文件執行的操作]

【範例操作】
1. 批量導出 PDF/DWG/STEP
2. 批量修改自定義屬性
3. 批量重命名文件
4. 批量更新工程圖
5. 批量統計信息

【輸出要求】
- 在終端顯示處理進度
- 記錄成功/失敗的文件
- 完成後顯示統計摘要"
            });

            Prompts.Insert(3, new PromptItem
            {
                Name = "導出文件格式",
                IsDefault = true,
                Content = @"請幫我生成 SolidWorks 文件格式導出代碼。

【導出設置】
- 源文件類型: [零件 / 裝配體 / 工程圖]
- 目標格式: [選擇格式]
  □ PDF (工程圖)
  □ DWG/DXF (工程圖)
  □ STEP (.step)
  □ IGES (.igs)
  □ Parasolid (.x_t)
  □ STL (3D打印)
  □ 3D PDF
  □ eDrawings
  □ 圖片 (PNG/JPG/BMP)

- 保存位置: [同目錄 / 指定路徑 / 相對路徑]
- 文件命名: [相同名稱 / 添加後綴 / 自定義規則]

【特殊選項】
對於 PDF:
- 紙張大小: [A4 / A3 / 自動]
- 包含所有圖紙: [是/否]

對於 STEP:
- 版本: [AP214 / AP203]

對於 STL:
- 質量: [精細 / 標準 / 粗糙]
- 單位: [mm / inch]

【其他需求】
[補充說明]"
            });

            Prompts.Insert(4, new PromptItem
            {
                Name = "幾何操作",
                IsDefault = true,
                Content = @"請幫我生成 SolidWorks 幾何操作代碼。

【操作類型】
□ 獲取面/邊/點信息
□ 計算體積/表面積/重心
□ 測量距離/角度
□ 獲取包圍盒尺寸
□ 查找最大/最小面
□ 遍歷拓撲結構
□ 干涉檢查
□ 截面分析

【常用幾何 API 範例】

// 獲取所有面
IBody2 swBody = ...;
object[] faces = swBody.GetFaces() as object[];
foreach (IFace2 face in faces)
{
    double area = face.GetArea();
    Print($""面積: {area * 1000000:F2} mm²"");
}

// 獲取包圍盒
double[] box = swBody.GetBodyBox() as double[];
double width = (box[3] - box[0]) * 1000;  // 轉換為 mm
double height = (box[4] - box[1]) * 1000;
double depth = (box[5] - box[2]) * 1000;

// 獲取質量屬性
IModelDocExtension ext = swModel.Extension;
IMassProperty massProp = ext.CreateMassProperty();
massProp.UseSystemUnits = false;
double mass = massProp.Mass;
double[] cog = massProp.CenterOfMass as double[];

【具體需求】
[描述你需要的幾何操作]"
            });
        }

        public static void EnsureFolderStructure()
        {
            string basePath = DefaultMacrosPath;
            
            string[] folders = {
                Path.Combine(basePath, "C Sharp"),
                Path.Combine(basePath, "VBA"),
                Path.Combine(basePath, "Prompts"),
                Path.Combine(basePath, "Notes")
            };

            foreach (var folder in folders)
            {
                if (!Directory.Exists(folder))
                {
                    Directory.CreateDirectory(folder);
                }
            }
        }
    }
}
