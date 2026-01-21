using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace MiniSolidworkAutomator.Controls
{
    /// <summary>
    /// Code snippets/templates panel
    /// </summary>
    public class CodeSnippetsPanel : Panel
    {
        private TreeView snippetTree = null!;
        private RichTextBox previewBox = null!;
        private Button insertButton = null!;
        
        private static readonly Color DarkBackground = Color.FromArgb(30, 30, 30);
        private static readonly Color DarkPanel = Color.FromArgb(45, 45, 45);
        private static readonly Color TextWhite = Color.White;
        private static readonly Color AccentBlue = Color.FromArgb(33, 150, 243);

        public event EventHandler<string>? InsertSnippet;

        public CodeSnippetsPanel()
        {
            InitializeComponents();
            LoadSnippets();
        }

        private void InitializeComponents()
        {
            this.BackColor = DarkPanel;
            this.Padding = new Padding(0);  // Remove padding to eliminate white borders
            this.BorderStyle = BorderStyle.None;

            var split = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Horizontal,
                SplitterDistance = 200,
                BackColor = DarkPanel,
                Panel1MinSize = 100,
                Panel2MinSize = 80,
                BorderStyle = BorderStyle.None,
                SplitterWidth = 3
            };
            split.Panel1.BackColor = DarkPanel;
            split.Panel2.BackColor = DarkBackground;

            // Snippet tree
            snippetTree = new TreeView
            {
                Dock = DockStyle.Fill,
                BackColor = DarkPanel,
                ForeColor = TextWhite,
                BorderStyle = BorderStyle.None,
                Font = new Font("Segoe UI", 9),
                ShowLines = true,
                ShowPlusMinus = true,
                ItemHeight = 22
            };
            snippetTree.AfterSelect += SnippetTree_AfterSelect;
            snippetTree.NodeMouseDoubleClick += (s, e) => InsertCurrentSnippet();

            // Preview area
            var previewPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = DarkBackground,
                Padding = new Padding(5)
            };

            previewBox = new RichTextBox
            {
                Dock = DockStyle.Fill,
                BackColor = DarkBackground,
                ForeColor = Color.FromArgb(180, 180, 180),
                BorderStyle = BorderStyle.None,
                Font = new Font("Cascadia Code", 9),
                ReadOnly = true,
                WordWrap = false
            };

            insertButton = new Button
            {
                Text = "📥 插入代碼 / Insert",
                Dock = DockStyle.Bottom,
                Height = 32,
                FlatStyle = FlatStyle.Flat,
                BackColor = AccentBlue,
                ForeColor = TextWhite
            };
            insertButton.Click += (s, e) => InsertCurrentSnippet();

            previewPanel.Controls.Add(previewBox);
            previewPanel.Controls.Add(insertButton);

            split.Panel1.Controls.Add(snippetTree);
            split.Panel2.Controls.Add(previewPanel);

            this.Controls.Add(split);
        }

        private void SnippetTree_AfterSelect(object? sender, TreeViewEventArgs e)
        {
            if (e.Node?.Tag is CodeSnippet snippet)
            {
                ApplySyntaxHighlighting(snippet.Code);
            }
            else
            {
                previewBox.Text = "";
            }
        }

        private void ApplySyntaxHighlighting(string code)
        {
            previewBox.SuspendLayout();
            previewBox.Text = code;
            
            // Define colors
            var keywordColor = Color.FromArgb(86, 156, 214);   // Blue for keywords
            var typeColor = Color.FromArgb(78, 201, 176);       // Teal for types
            var stringColor = Color.FromArgb(214, 157, 133);    // Orange for strings
            var commentColor = Color.FromArgb(106, 153, 85);    // Green for comments
            var numberColor = Color.FromArgb(181, 206, 168);    // Light green for numbers
            
            // Keywords
            string[] keywords = { "if", "else", "return", "var", "int", "string", "bool", "true", "false", 
                                  "null", "new", "foreach", "for", "while", "in", "as", "is", "out", "ref",
                                  "try", "catch", "finally", "throw", "using", "namespace", "class", "public",
                                  "private", "void", "static", "object", "double", "float", "long" };
            
            // Types
            string[] types = { "IPartDoc", "IAssemblyDoc", "IDrawingDoc", "IModelDoc2", "ISldWorks", 
                              "IFeature", "IBody2", "IFace2", "IComponent2", "IView", "ISheet",
                              "Color", "Print", "PrintError", "PrintWarning", "Path", "File" };
            
            // Highlight comments first (// style)
            int pos = 0;
            while ((pos = code.IndexOf("//", pos)) >= 0)
            {
                int endPos = code.IndexOf('\n', pos);
                if (endPos < 0) endPos = code.Length;
                HighlightRange(pos, endPos - pos, commentColor);
                pos = endPos;
            }
            
            // Highlight strings
            pos = 0;
            while ((pos = code.IndexOf('"', pos)) >= 0)
            {
                int endPos = code.IndexOf('"', pos + 1);
                if (endPos < 0) break;
                HighlightRange(pos, endPos - pos + 1, stringColor);
                pos = endPos + 1;
            }
            
            // Highlight keywords
            foreach (var keyword in keywords)
            {
                HighlightWord(keyword, keywordColor);
            }
            
            // Highlight types
            foreach (var type in types)
            {
                HighlightWord(type, typeColor);
            }
            
            previewBox.ResumeLayout();
        }
        
        private void HighlightWord(string word, Color color)
        {
            int pos = 0;
            string text = previewBox.Text;
            while ((pos = text.IndexOf(word, pos, StringComparison.Ordinal)) >= 0)
            {
                // Check if it's a whole word
                bool isStart = pos == 0 || !char.IsLetterOrDigit(text[pos - 1]);
                bool isEnd = pos + word.Length >= text.Length || !char.IsLetterOrDigit(text[pos + word.Length]);
                
                if (isStart && isEnd)
                {
                    previewBox.Select(pos, word.Length);
                    previewBox.SelectionColor = color;
                }
                pos += word.Length;
            }
            previewBox.Select(0, 0);
        }
        
        private void HighlightRange(int start, int length, Color color)
        {
            if (start >= 0 && length > 0 && start + length <= previewBox.Text.Length)
            {
                previewBox.Select(start, length);
                previewBox.SelectionColor = color;
            }
            previewBox.Select(0, 0);
        }

        private void InsertCurrentSnippet()
        {
            if (snippetTree.SelectedNode?.Tag is CodeSnippet snippet)
            {
                InsertSnippet?.Invoke(this, snippet.Code);
            }
        }

        private void LoadSnippets()
        {
            snippetTree.Nodes.Clear();

            // Basic Operations
            var basicNode = new TreeNode("📁 基本操作 / Basic") { ForeColor = Color.FromArgb(255, 213, 79) };
            AddSnippet(basicNode, "檢查文檔", "Check Document", @"// 檢查活動文檔
if (swModel == null) 
{
    PrintError(""請先打開文檔 / Please open a document first"");
    return;
}
Print($""當前文檔: {swModel.GetTitle()}"");
Print($""文檔類型: {swModel.GetType()}"");
Print($""文檔路徑: {swModel.GetPathName()}"");");

            AddSnippet(basicNode, "判斷文檔類型", "Check Doc Type", @"// 判斷文檔類型
int docType = swModel.GetType();
if (docType == (int)swDocumentTypes_e.swDocPART) 
{
    Print(""這是零件 / This is a Part"");
}
else if (docType == (int)swDocumentTypes_e.swDocASSEMBLY) 
{
    Print(""這是裝配體 / This is an Assembly"");
}
else if (docType == (int)swDocumentTypes_e.swDocDRAWING) 
{
    Print(""這是工程圖 / This is a Drawing"");
}");

            AddSnippet(basicNode, "保存文檔", "Save Document", @"// 保存當前文檔
int errors = 0, warnings = 0;
bool success = swModel.Save3(
    (int)swSaveAsOptions_e.swSaveAsOptions_Silent, 
    ref errors, 
    ref warnings
);
if (success)
    Print(""✅ 保存成功"");
else
    PrintError($""保存失敗, 錯誤碼: {errors}"");");

            snippetTree.Nodes.Add(basicNode);

            // Part Operations
            var partNode = new TreeNode("📁 零件操作 / Part") { ForeColor = Color.FromArgb(255, 213, 79) };
            AddSnippet(partNode, "遍歷實體", "Iterate Bodies", @"// 遍歷零件中的所有實體
var swPart = swModel as IPartDoc;
if (swPart == null) { PrintError(""請打開零件文檔""); return; }

object[] bodies = swPart.GetBodies2((int)swBodyType_e.swSolidBody, true) as object[];
if (bodies == null || bodies.Length == 0)
{
    PrintWarning(""沒有找到實體"");
    return;
}

Print($""找到 {bodies.Length} 個實體:"");
foreach (IBody2 body in bodies)
{
    Print($""  - {body.Name}"");
}");

            AddSnippet(partNode, "遍歷特徵", "Iterate Features", @"// 遍歷所有特徵
IFeature feat = swModel.IFirstFeature();
while (feat != null)
{
    Print($""{feat.GetTypeName2()}: {feat.Name}"");
    feat = feat.IGetNextFeature();
}");

            AddSnippet(partNode, "獲取質量屬性", "Mass Properties", @"// 獲取質量屬性
var ext = swModel.Extension;
var massProp = ext.CreateMassProperty2();
if (massProp != null)
{
    Print($""質量: {massProp.Mass:F4} kg"");
    Print($""體積: {massProp.Volume * 1e9:F2} mm³"");
    Print($""表面積: {massProp.SurfaceArea * 1e6:F2} mm²"");
    var cog = massProp.CenterOfMass as double[];
    if (cog != null)
        Print($""重心: ({cog[0]*1000:F2}, {cog[1]*1000:F2}, {cog[2]*1000:F2}) mm"");
}");

            snippetTree.Nodes.Add(partNode);

            // Assembly Operations
            var assyNode = new TreeNode("📁 裝配體操作 / Assembly") { ForeColor = Color.FromArgb(255, 213, 79) };
            AddSnippet(assyNode, "遍歷組件", "Iterate Components", @"// 遍歷裝配體中的組件
var swAssy = swModel as IAssemblyDoc;
if (swAssy == null) { PrintError(""請打開裝配體文檔""); return; }

object[] comps = swAssy.GetComponents(false) as object[];
if (comps == null) { PrintWarning(""沒有組件""); return; }

Print($""共有 {comps.Length} 個組件:"");
foreach (IComponent2 comp in comps)
{
    string status = comp.IsSuppressed() ? ""[抑制]"" : """";
    Print($""  {comp.Name2} {status}"");
}");

            AddSnippet(assyNode, "選中的組件", "Selected Components", @"// 獲取選中的組件
var selMgr = swModel.ISelectionManager;
int count = selMgr.GetSelectedObjectCount2(-1);
Print($""選中了 {count} 個對象"");

for (int i = 1; i <= count; i++)
{
    var comp = selMgr.GetSelectedObjectsComponent4(i, -1) as IComponent2;
    if (comp != null)
    {
        Print($""  組件: {comp.Name2}"");
        Print($""    路徑: {comp.GetPathName()}"");
    }
}");

            snippetTree.Nodes.Add(assyNode);

            // Drawing Operations
            var drawNode = new TreeNode("📁 工程圖操作 / Drawing") { ForeColor = Color.FromArgb(255, 213, 79) };
            AddSnippet(drawNode, "遍歷圖紙", "Iterate Sheets", @"// 遍歷工程圖中的所有圖紙
var swDraw = swModel as IDrawingDoc;
if (swDraw == null) { PrintError(""請打開工程圖文檔""); return; }

var sheetNames = swDraw.GetSheetNames() as string[];
if (sheetNames == null) return;

Print($""共有 {sheetNames.Length} 張圖紙:"");
foreach (string name in sheetNames)
{
    Print($""  - {name}"");
}");

            AddSnippet(drawNode, "遍歷視圖", "Iterate Views", @"// 遍歷當前圖紙的所有視圖
var swDraw = swModel as IDrawingDoc;
if (swDraw == null) { PrintError(""請打開工程圖文檔""); return; }

var sheet = swDraw.IGetCurrentSheet();
Print($""當前圖紙: {sheet.GetName()}"");

var views = sheet.GetViews() as object[];
if (views != null)
{
    Print($""視圖數量: {views.Length}"");
    foreach (IView view in views)
    {
        Print($""  - {view.Name} ({view.Type})"");
    }
}");

            snippetTree.Nodes.Add(drawNode);

            // Custom Properties
            var propNode = new TreeNode("📁 自定義屬性 / Properties") { ForeColor = Color.FromArgb(255, 213, 79) };
            AddSnippet(propNode, "讀取屬性", "Read Properties", @"// 讀取自定義屬性
var ext = swModel.Extension;
var propMgr = ext.get_CustomPropertyManager("""");

string val = """", resolvedVal = """";
bool wasResolved = false;
propMgr.Get6(""屬性名稱"", false, out val, out resolvedVal, out wasResolved, out _);
Print($""屬性值: {val}"");
Print($""解析值: {resolvedVal}"");");

            AddSnippet(propNode, "設置屬性", "Set Property", @"// 設置自定義屬性
var ext = swModel.Extension;
var propMgr = ext.get_CustomPropertyManager("""");

// 添加或更新屬性
propMgr.Add3(
    ""項目編號"",                              // 屬性名稱
    (int)swCustomInfoType_e.swCustomInfoText, // 類型
    ""PRJ-001"",                               // 值
    (int)swCustomPropertyAddOption_e.swCustomPropertyReplaceValue
);
Print(""✅ 屬性已設置"");");

            AddSnippet(propNode, "列出所有屬性", "List Properties", @"// 列出所有自定義屬性
var ext = swModel.Extension;
var propMgr = ext.get_CustomPropertyManager("""");

var names = propMgr.GetNames() as string[];
if (names == null || names.Length == 0)
{
    Print(""沒有自定義屬性"");
    return;
}

Print($""共有 {names.Length} 個自定義屬性:"");
foreach (string name in names)
{
    string val = """", resolvedVal = """";
    propMgr.Get6(name, false, out val, out resolvedVal, out _, out _);
    Print($""  {name}: {val}"");
}");

            snippetTree.Nodes.Add(propNode);

            // Export Operations
            var exportNode = new TreeNode("📁 導出操作 / Export") { ForeColor = Color.FromArgb(255, 213, 79) };
            AddSnippet(exportNode, "導出為PDF", "Export PDF", @"// 導出工程圖為PDF
var swDraw = swModel as IDrawingDoc;
if (swDraw == null) { PrintError(""請打開工程圖文檔""); return; }

string path = swModel.GetPathName();
string pdfPath = Path.ChangeExtension(path, "".pdf"");

var ext = swModel.Extension;
int errors = 0, warnings = 0;

var exportData = swApp.GetExportFileData((int)swExportDataFileType_e.swExportPdfData) as IExportPdfData;
if (exportData != null)
{
    exportData.ExportAsOne = true;
    exportData.ViewPdfAfterSaving = false;
}

bool success = ext.SaveAs3(pdfPath, 0, 0, exportData, null, ref errors, ref warnings);
if (success)
    Print($""✅ PDF 已導出: {pdfPath}"");
else
    PrintError($""導出失敗, 錯誤碼: {errors}"");");

            AddSnippet(exportNode, "導出為STEP", "Export STEP", @"// 導出為STEP格式
string path = swModel.GetPathName();
string stepPath = Path.ChangeExtension(path, "".step"");

int errors = 0, warnings = 0;
bool success = swModel.Extension.SaveAs3(
    stepPath,
    (int)swSaveAsVersion_e.swSaveAsCurrentVersion,
    (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
    null, null, ref errors, ref warnings
);

if (success)
    Print($""✅ STEP 已導出: {stepPath}"");
else
    PrintError($""導出失敗, 錯誤碼: {errors}"");");

            AddSnippet(exportNode, "批量導出", "Batch Export", @"// 批量導出裝配體中的零件
var swAssy = swModel as IAssemblyDoc;
if (swAssy == null) { PrintError(""請打開裝配體""); return; }

string outputDir = Path.GetDirectoryName(swModel.GetPathName());
var comps = swAssy.GetComponents(true) as object[];

int exported = 0;
foreach (IComponent2 comp in comps)
{
    string compPath = comp.GetPathName();
    if (string.IsNullOrEmpty(compPath)) continue;
    
    string name = Path.GetFileNameWithoutExtension(compPath);
    string stepPath = Path.Combine(outputDir, name + "".step"");
    
    var compModel = comp.GetModelDoc2() as IModelDoc2;
    if (compModel != null)
    {
        int err = 0, warn = 0;
        compModel.Extension.SaveAs3(stepPath, 0, 0, null, null, ref err, ref warn);
        exported++;
        Print($""  導出: {name}.step"");
    }
}
Print($""✅ 共導出 {exported} 個文件"");");

            snippetTree.Nodes.Add(exportNode);

            // Selection Operations
            var selNode = new TreeNode("📁 選擇操作 / Selection") { ForeColor = Color.FromArgb(255, 213, 79) };
            AddSnippet(selNode, "獲取選擇", "Get Selection", @"// 獲取當前選擇的對象
var selMgr = swModel.ISelectionManager;
int count = selMgr.GetSelectedObjectCount2(-1);

if (count == 0)
{
    PrintWarning(""請先選擇對象"");
    return;
}

Print($""選中了 {count} 個對象:"");
for (int i = 1; i <= count; i++)
{
    int type = selMgr.GetSelectedObjectType3(i, -1);
    Print($""  [{i}] 類型: {type}"");
}");

            AddSnippet(selNode, "選擇面", "Select Face", @"// 選擇所有平面
swModel.ClearSelection2(true);

var swPart = swModel as IPartDoc;
if (swPart == null) return;

var bodies = swPart.GetBodies2((int)swBodyType_e.swSolidBody, true) as object[];
int faceCount = 0;

foreach (IBody2 body in bodies)
{
    var faces = body.GetFaces() as object[];
    foreach (IFace2 face in faces)
    {
        var surf = face.IGetSurface();
        if (surf.IsPlane())
        {
            face.Select4(true, null);
            faceCount++;
        }
    }
}
Print($""選中了 {faceCount} 個平面"");");

            snippetTree.Nodes.Add(selNode);

            snippetTree.ExpandAll();
        }

        private void AddSnippet(TreeNode parent, string nameZh, string nameEn, string code)
        {
            var node = new TreeNode($"📄 {nameZh}")
            {
                Tag = new CodeSnippet { Name = $"{nameZh} / {nameEn}", Code = code },
                ForeColor = TextWhite
            };
            parent.Nodes.Add(node);
        }
    }

    public class CodeSnippet
    {
        public string Name { get; set; } = "";
        public string Code { get; set; } = "";
    }
}
