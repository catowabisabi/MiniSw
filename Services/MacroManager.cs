using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using MiniSolidworkAutomator.Models;

namespace MiniSolidworkAutomator.Services
{
    public class MacroFileInfo
    {
        public string Name { get; set; } = "";
        public string FullPath { get; set; } = "";
        public string RelativePath { get; set; } = "";
        public MacroType Type { get; set; }
        public string Source { get; set; } = ""; // Which path it came from
        public DateTime ModifiedDate { get; set; }
    }

    public class MacroManager
    {
        private readonly AppSettings _settings;
        private List<MacroFileInfo> _cachedMacros = new List<MacroFileInfo>();

        public MacroManager(AppSettings settings)
        {
            _settings = settings;
        }

        public List<MacroFileInfo> GetAllMacros(bool forceRefresh = false)
        {
            if (!forceRefresh && _cachedMacros.Count > 0)
                return _cachedMacros;

            _cachedMacros.Clear();

            foreach (var basePath in _settings.MacroPaths)
            {
                if (!Directory.Exists(basePath))
                    continue;

                ScanDirectory(basePath, basePath);
            }

            return _cachedMacros;
        }

        public List<MacroFileInfo> GetCSharpMacros(bool forceRefresh = false)
        {
            return GetAllMacros(forceRefresh)
                .Where(m => m.Type == MacroType.CSharp)
                .OrderBy(m => m.Name)
                .ToList();
        }

        public List<MacroFileInfo> GetVBAMacros(bool forceRefresh = false)
        {
            return GetAllMacros(forceRefresh)
                .Where(m => m.Type == MacroType.VBA)
                .OrderBy(m => m.Name)
                .ToList();
        }

        private void ScanDirectory(string directory, string basePath)
        {
            try
            {
                // Get files in current directory
                var files = Directory.GetFiles(directory);
                
                foreach (var file in files)
                {
                    var ext = Path.GetExtension(file).ToLowerInvariant();
                    MacroType? type = null;

                    switch (ext)
                    {
                        case ".cs":
                            type = MacroType.CSharp;
                            break;
                        case ".vba":
                        case ".bas":
                        case ".swp":
                            type = MacroType.VBA;
                            break;
                    }

                    if (type.HasValue)
                    {
                        var fileInfo = new FileInfo(file);
                        _cachedMacros.Add(new MacroFileInfo
                        {
                            Name = Path.GetFileNameWithoutExtension(file),
                            FullPath = file,
                            RelativePath = GetRelativePath(file, basePath),
                            Type = type.Value,
                            Source = basePath,
                            ModifiedDate = fileInfo.LastWriteTime
                        });
                    }
                }

                // Recursively scan subdirectories
                foreach (var subDir in Directory.GetDirectories(directory))
                {
                    ScanDirectory(subDir, basePath);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"掃描目錄錯誤 {directory}: {ex.Message}");
            }
        }

        private string GetRelativePath(string fullPath, string basePath)
        {
            if (fullPath.StartsWith(basePath, StringComparison.OrdinalIgnoreCase))
            {
                var relative = fullPath.Substring(basePath.Length);
                return relative.TrimStart(Path.DirectorySeparatorChar);
            }
            return Path.GetFileName(fullPath);
        }

        public string LoadMacroContent(string filePath)
        {
            try
            {
                if (File.Exists(filePath))
                {
                    return File.ReadAllText(filePath);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"載入文件錯誤: {ex.Message}");
            }
            return "";
        }

        public bool SaveMacroContent(string filePath, string content)
        {
            try
            {
                string? directory = Path.GetDirectoryName(filePath);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                File.WriteAllText(filePath, content);
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"保存文件錯誤: {ex.Message}");
                return false;
            }
        }

        public Dictionary<string, List<MacroFileInfo>> GroupBySource(List<MacroFileInfo> macros)
        {
            return macros
                .GroupBy(m => GetSourceDisplayName(m.Source))
                .ToDictionary(g => g.Key, g => g.ToList());
        }

        private string GetSourceDisplayName(string path)
        {
            if (path.Equals(AppSettings.DefaultMacrosPath, StringComparison.OrdinalIgnoreCase))
                return "📁 預設";
            
            return $"📁 {Path.GetFileName(path)}";
        }
    }
}
