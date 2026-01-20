using System.Collections.Generic;
using System.Globalization;

namespace MiniSolidworkAutomator.Localization
{
    public static class Lang
    {
        public static string CurrentLanguage { get; private set; } = "zh-TW";

        private static readonly Dictionary<string, Dictionary<string, string>> Translations = new()
        {
            ["zh-TW"] = new Dictionary<string, string>
            {
                // Window
                ["AppTitle"] = "MiniSW - SolidWorks 宏自動化工具",
                
                // Toolbar
                ["Connect"] = "連接",
                ["Disconnect"] = "斷開",
                ["Run"] = "運行",
                ["Stop"] = "停止",
                ["New"] = "新建",
                ["Save"] = "保存",
                ["Settings"] = "設定",
                ["Refresh"] = "刷新",
                
                // Connection status
                ["NotConnected"] = "未連接",
                ["Connected"] = "已連接",
                ["Connecting"] = "連接中...",
                
                // Tabs
                ["CSharpMacros"] = "C# 宏",
                ["VBAMacros"] = "VBA 宏",
                ["Prompts"] = "提示模板",
                ["Notes"] = "筆記",
                
                // Editor
                ["NewScript"] = "新腳本",
                ["UnsavedChanges"] = "未保存的更改",
                ["CloseTab"] = "關閉",
                ["CloseOthers"] = "關閉其他",
                ["Rename"] = "重命名",
                ["SelectFile"] = "選擇文件...",
                ["As"] = "為",
                
                // Terminal
                ["Terminal"] = "終端輸出",
                ["Clear"] = "清除",
                ["CopyOutput"] = "複製輸出",
                ["Copied"] = "已複製到剪貼板",
                
                // File browser
                ["FileBrowser"] = "文件瀏覽器",
                ["AddNew"] = "新增",
                ["Delete"] = "刪除",
                ["DoubleClickToOpen"] = "雙擊打開文件",
                
                // Prompts
                ["PromptName"] = "提示名稱",
                ["PromptContent"] = "提示內容",
                ["CopyToClipboard"] = "複製到剪貼板",
                ["DefaultPrompt"] = "默認提示",
                ["CustomPrompt"] = "自定義提示",
                
                // Notes
                ["NoteName"] = "筆記標題",
                ["NoteContent"] = "筆記內容",
                
                // Settings
                ["SettingsTitle"] = "設定",
                ["Language"] = "語言",
                ["MacroPaths"] = "宏文件路徑",
                ["AddPath"] = "添加路徑",
                ["RemovePath"] = "移除路徑",
                ["MoveUp"] = "上移",
                ["MoveDown"] = "下移",
                ["DefaultPath"] = "默認路徑",
                ["CannotRemoveDefault"] = "無法移除默認路徑",
                ["Apply"] = "應用",
                ["Cancel"] = "取消",
                ["OK"] = "確定",
                
                // Messages
                ["Ready"] = "準備就緒",
                ["ExecutionStarted"] = "開始執行",
                ["ExecutionCompleted"] = "執行完成",
                ["ExecutionCancelled"] = "執行已取消",
                ["CompilationError"] = "編譯錯誤",
                ["RuntimeError"] = "運行時錯誤",
                ["SaveSuccess"] = "保存成功",
                ["SaveFailed"] = "保存失敗",
                ["SettingsUpdated"] = "設定已更新",
                ["ListRefreshed"] = "列表已刷新",
                ["SWConnected"] = "SolidWorks 已連接",
                ["SWNotConnected"] = "SolidWorks 未連接",
                ["OnlyCSCanRun"] = "只能運行 C# 代碼",
                ["VBANeedsSW"] = "VBA 宏需要 SolidWorks 連接",
                ["RunningVBA"] = "運行 VBA 宏",
                ["VBASuccess"] = "VBA 宏執行成功",
                ["VBAFailed"] = "VBA 宏執行失敗",
                ["MecAgentDetected"] = "檢測到 MecAgent 格式，正在轉換...",
                ["ConversionComplete"] = "轉換完成",
                ["NoDocOpen"] = "沒有打開的文檔",
                ["PleaseOpenDoc"] = "請先打開文檔",
                
                // Dialogs
                ["Confirm"] = "確認",
                ["UnsavedConfirm"] = "有未保存的更改，確定要關閉嗎？",
                ["EnterNewName"] = "輸入新名稱",
                ["CannotDeleteDefault"] = "無法刪除默認項目",
                ["SelectFolder"] = "選擇文件夾",
            },
            
            ["en-US"] = new Dictionary<string, string>
            {
                // Window
                ["AppTitle"] = "MiniSW - SolidWorks Macro Automator",
                
                // Toolbar
                ["Connect"] = "Connect",
                ["Disconnect"] = "Disconnect",
                ["Run"] = "Run",
                ["Stop"] = "Stop",
                ["New"] = "New",
                ["Save"] = "Save",
                ["Settings"] = "Settings",
                ["Refresh"] = "Refresh",
                
                // Connection status
                ["NotConnected"] = "Not Connected",
                ["Connected"] = "Connected",
                ["Connecting"] = "Connecting...",
                
                // Tabs
                ["CSharpMacros"] = "C# Macros",
                ["VBAMacros"] = "VBA Macros",
                ["Prompts"] = "Prompts",
                ["Notes"] = "Notes",
                
                // Editor
                ["NewScript"] = "New Script",
                ["UnsavedChanges"] = "Unsaved Changes",
                ["CloseTab"] = "Close",
                ["CloseOthers"] = "Close Others",
                ["Rename"] = "Rename",
                ["SelectFile"] = "Select file...",
                ["As"] = "As",
                
                // Terminal
                ["Terminal"] = "Terminal Output",
                ["Clear"] = "Clear",
                ["CopyOutput"] = "Copy Output",
                ["Copied"] = "Copied to clipboard",
                
                // File browser
                ["FileBrowser"] = "File Browser",
                ["AddNew"] = "Add New",
                ["Delete"] = "Delete",
                ["DoubleClickToOpen"] = "Double-click to open",
                
                // Prompts
                ["PromptName"] = "Prompt Name",
                ["PromptContent"] = "Prompt Content",
                ["CopyToClipboard"] = "Copy to Clipboard",
                ["DefaultPrompt"] = "Default Prompts",
                ["CustomPrompt"] = "Custom Prompts",
                
                // Notes
                ["NoteName"] = "Note Title",
                ["NoteContent"] = "Note Content",
                
                // Settings
                ["SettingsTitle"] = "Settings",
                ["Language"] = "Language",
                ["MacroPaths"] = "Macro Paths",
                ["AddPath"] = "Add Path",
                ["RemovePath"] = "Remove Path",
                ["MoveUp"] = "Move Up",
                ["MoveDown"] = "Move Down",
                ["DefaultPath"] = "Default Path",
                ["CannotRemoveDefault"] = "Cannot remove default path",
                ["Apply"] = "Apply",
                ["Cancel"] = "Cancel",
                ["OK"] = "OK",
                
                // Messages
                ["Ready"] = "Ready",
                ["ExecutionStarted"] = "Execution started",
                ["ExecutionCompleted"] = "Execution completed",
                ["ExecutionCancelled"] = "Execution cancelled",
                ["CompilationError"] = "Compilation error",
                ["RuntimeError"] = "Runtime error",
                ["SaveSuccess"] = "Saved successfully",
                ["SaveFailed"] = "Save failed",
                ["SettingsUpdated"] = "Settings updated",
                ["ListRefreshed"] = "List refreshed",
                ["SWConnected"] = "SolidWorks connected",
                ["SWNotConnected"] = "SolidWorks not connected",
                ["OnlyCSCanRun"] = "Only C# code can be executed",
                ["VBANeedsSW"] = "VBA macros require SolidWorks connection",
                ["RunningVBA"] = "Running VBA macro",
                ["VBASuccess"] = "VBA macro executed successfully",
                ["VBAFailed"] = "VBA macro execution failed",
                ["MecAgentDetected"] = "MecAgent format detected, converting...",
                ["ConversionComplete"] = "Conversion complete",
                ["NoDocOpen"] = "No document open",
                ["PleaseOpenDoc"] = "Please open a document first",
                
                // Dialogs
                ["Confirm"] = "Confirm",
                ["UnsavedConfirm"] = "There are unsaved changes. Close anyway?",
                ["EnterNewName"] = "Enter new name",
                ["CannotDeleteDefault"] = "Cannot delete default item",
                ["SelectFolder"] = "Select Folder",
            }
        };

        public static void SetLanguage(string langCode)
        {
            if (Translations.ContainsKey(langCode))
            {
                CurrentLanguage = langCode;
            }
        }

        public static string Get(string key)
        {
            if (Translations.TryGetValue(CurrentLanguage, out var dict))
            {
                if (dict.TryGetValue(key, out var value))
                {
                    return value;
                }
            }
            
            // Fallback to English
            if (Translations.TryGetValue("en-US", out var enDict))
            {
                if (enDict.TryGetValue(key, out var value))
                {
                    return value;
                }
            }
            
            return $"[{key}]";
        }

        public static Dictionary<string, string> GetAvailableLanguages() => new()
        {
            ["zh-TW"] = "繁體中文",
            ["en-US"] = "English"
        };
        
        public static string GetLanguageDisplayName(string code) => code switch
        {
            "zh-TW" => "繁體中文",
            "en-US" => "English",
            _ => code
        };
    }
}
