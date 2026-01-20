using System;
using System.Text;
using System.Text.RegularExpressions;

namespace MiniSolidworkAutomator.Services
{
    /// <summary>
    /// Detects and converts MecAgent format code to MiniSW executable format
    /// </summary>
    public static class MecAgentConverter
    {
        /// <summary>
        /// Check if code is in MecAgent format (has DynamicClass with Execute method)
        /// </summary>
        public static bool IsMecAgentFormat(string code)
        {
            if (string.IsNullOrWhiteSpace(code)) return false;
            
            // MecAgent format markers
            bool hasHeader = code.Contains("// MecAgent Macro") || 
                             code.Contains("// Framework: solidworks");
            bool hasDynamicClass = Regex.IsMatch(code, @"public\s+class\s+DynamicClass");
            bool hasExecuteMethod = Regex.IsMatch(code, @"public\s+static\s+object\s+Execute\s*\(\s*\)");
            
            return hasDynamicClass && hasExecuteMethod;
        }

        /// <summary>
        /// Check if code is standard MiniSW script format
        /// </summary>
        public static bool IsStandardFormat(string code)
        {
            if (string.IsNullOrWhiteSpace(code)) return false;
            
            // Standard format uses swApp and swModel directly without class wrapper
            bool hasNoClass = !Regex.IsMatch(code, @"public\s+class\s+\w+");
            bool usesGlobals = code.Contains("swApp") || code.Contains("swModel") || 
                               code.Contains("Print(") || code.Contains("PrintError(");
            
            return hasNoClass || usesGlobals;
        }

        /// <summary>
        /// Convert MecAgent format to MiniSW executable script
        /// </summary>
        public static string ConvertToMiniSW(string mecAgentCode)
        {
            if (string.IsNullOrWhiteSpace(mecAgentCode)) return mecAgentCode;

            try
            {
                var sb = new StringBuilder();
                
                // Add header comment
                sb.AppendLine("// Converted from MecAgent format");
                sb.AppendLine("// Variables: swApp (ISldWorks), swModel (IModelDoc2)");
                sb.AppendLine("// Functions: Print(), PrintError(), PrintWarning()");
                sb.AppendLine();

                // Extract the content inside Execute() method
                string executeContent = ExtractExecuteMethodContent(mecAgentCode);
                
                if (!string.IsNullOrEmpty(executeContent))
                {
                    // Transform the code
                    string transformed = TransformMecAgentCode(executeContent);
                    sb.Append(transformed);
                }
                else
                {
                    // If extraction failed, try to run as-is with wrapper
                    sb.AppendLine("// Could not extract Execute method, running original code");
                    sb.Append(mecAgentCode);
                }

                return sb.ToString();
            }
            catch (Exception ex)
            {
                return $"// Conversion error: {ex.Message}\n{mecAgentCode}";
            }
        }

        /// <summary>
        /// Extract the body of the Execute() method
        /// </summary>
        private static string ExtractExecuteMethodContent(string code)
        {
            // Find Execute method
            var executeMatch = Regex.Match(code, 
                @"public\s+static\s+object\s+Execute\s*\(\s*\)\s*\{", 
                RegexOptions.Singleline);
            
            if (!executeMatch.Success) return "";

            int startIndex = executeMatch.Index + executeMatch.Length;
            int braceCount = 1;
            int endIndex = startIndex;

            // Find matching closing brace
            for (int i = startIndex; i < code.Length && braceCount > 0; i++)
            {
                if (code[i] == '{') braceCount++;
                else if (code[i] == '}') braceCount--;
                
                if (braceCount == 0)
                {
                    endIndex = i;
                    break;
                }
            }

            if (endIndex > startIndex)
            {
                return code.Substring(startIndex, endIndex - startIndex).Trim();
            }

            return "";
        }

        /// <summary>
        /// Transform MecAgent code patterns to MiniSW patterns
        /// </summary>
        private static string TransformMecAgentCode(string code)
        {
            var result = code;

            // Remove try-catch wrapper at the top level (we handle errors differently)
            result = RemoveOuterTryCatch(result);

            // Replace Console.WriteLine with Print
            result = Regex.Replace(result, @"Console\.WriteLine\s*\(\s*""?\[START\]", "Print(\"[START]");
            result = Regex.Replace(result, @"Console\.WriteLine\s*\(\s*""?\[INFO\]", "Print(\"[INFO]");
            result = Regex.Replace(result, @"Console\.WriteLine\s*\(\s*""?\[ERROR\]", "PrintError(\"[ERROR]");
            result = Regex.Replace(result, @"Console\.WriteLine\s*\(\s*""?\[WARNING\]", "PrintWarning(\"[WARNING]");
            result = Regex.Replace(result, @"Console\.WriteLine\s*\(", "Print(");

            // Replace SolidWorks connection code (MecAgent creates new instance, we use existing)
            result = Regex.Replace(result, 
                @"SldWorks\s+swApp\s*=\s*Activator\.CreateInstance[^;]+;",
                "// swApp is already connected (provided by MiniSW)");
            
            result = Regex.Replace(result,
                @"if\s*\(\s*swApp\s*==\s*null\s*\)\s*\{[^}]+throw[^}]+\}",
                "// Connection check handled by MiniSW",
                RegexOptions.Singleline);

            // Replace return statements that return anonymous objects
            result = Regex.Replace(result,
                @"return\s+new\s*\{[^}]+success\s*=\s*false[^}]+\};",
                m => ConvertReturnToError(m.Value));
            
            result = Regex.Replace(result,
                @"return\s+new\s*\{[^}]+success\s*=\s*true[^}]+\};",
                m => ConvertReturnToSuccess(m.Value));

            // Clean up multiple consecutive newlines
            result = Regex.Replace(result, @"\n{3,}", "\n\n");

            return result.Trim();
        }

        /// <summary>
        /// Remove outer try-catch if present
        /// </summary>
        private static string RemoveOuterTryCatch(string code)
        {
            // Check if code starts with try
            var tryMatch = Regex.Match(code.TrimStart(), @"^try\s*\{");
            if (!tryMatch.Success) return code;

            // Find matching brace for try block
            int startIndex = code.IndexOf('{') + 1;
            int braceCount = 1;
            int tryEndIndex = startIndex;

            for (int i = startIndex; i < code.Length && braceCount > 0; i++)
            {
                if (code[i] == '{') braceCount++;
                else if (code[i] == '}') braceCount--;
                
                if (braceCount == 0)
                {
                    tryEndIndex = i;
                    break;
                }
            }

            // Extract content inside try block
            string tryContent = code.Substring(startIndex, tryEndIndex - startIndex).Trim();

            // Check for catch block and skip it
            string afterTry = code.Substring(tryEndIndex + 1).TrimStart();
            if (afterTry.StartsWith("catch"))
            {
                return tryContent;
            }

            return code;
        }

        /// <summary>
        /// Convert error return to PrintError
        /// </summary>
        private static string ConvertReturnToError(string returnStatement)
        {
            // Extract error message
            var match = Regex.Match(returnStatement, @"error\s*=\s*""([^""]+)""");
            if (match.Success)
            {
                return $"PrintError(\"{match.Groups[1].Value}\"); return null;";
            }
            return "PrintError(\"Operation failed\"); return null;";
        }

        /// <summary>
        /// Convert success return to success message
        /// </summary>
        private static string ConvertReturnToSuccess(string returnStatement)
        {
            // Extract message if present
            var match = Regex.Match(returnStatement, @"message\s*=\s*""([^""]+)""");
            if (match.Success)
            {
                return $"Print(\"✅ {match.Groups[1].Value}\"); return true;";
            }
            return "Print(\"✅ Operation completed successfully\"); return true;";
        }

        /// <summary>
        /// Get code type description
        /// </summary>
        public static string GetCodeTypeDescription(string code)
        {
            if (IsMecAgentFormat(code))
                return "MecAgent 格式 (將自動轉換)";
            else if (IsStandardFormat(code))
                return "標準 MiniSW 格式";
            else
                return "未知格式";
        }
    }
}
