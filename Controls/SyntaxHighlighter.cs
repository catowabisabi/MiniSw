using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace MiniSolidworkAutomator.Controls
{
    /// <summary>
    /// Optimized syntax highlighting for C# and VBA code
    /// Uses cached regex patterns and minimal text manipulation
    /// </summary>
    public class SyntaxHighlighter
    {
        private static readonly HashSet<string> CSharpKeywords = new HashSet<string>
        {
            "abstract", "as", "base", "bool", "break", "byte", "case", "catch", "char",
            "checked", "class", "const", "continue", "decimal", "default", "delegate",
            "do", "double", "else", "enum", "event", "explicit", "extern", "false",
            "finally", "fixed", "float", "for", "foreach", "goto", "if", "implicit",
            "in", "int", "interface", "internal", "is", "lock", "long", "namespace",
            "new", "null", "object", "operator", "out", "override", "params", "private",
            "protected", "public", "readonly", "ref", "return", "sbyte", "sealed",
            "short", "sizeof", "stackalloc", "static", "string", "struct", "switch",
            "this", "throw", "true", "try", "typeof", "uint", "ulong", "unchecked",
            "unsafe", "ushort", "using", "virtual", "void", "volatile", "while", "var",
            "dynamic", "async", "await", "nameof", "when", "where", "yield"
        };

        private static readonly HashSet<string> VBAKeywords = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "And", "As", "Boolean", "ByRef", "Byte", "ByVal", "Call", "Case", "Const",
            "Currency", "Date", "Declare", "Dim", "Do", "Double", "Each", "Else", "ElseIf",
            "End", "Enum", "Error", "Exit", "False", "For", "Friend", "Function", "Get",
            "GoSub", "GoTo", "If", "Integer", "Is", "Let", "Like", "Long", "Loop", "Me",
            "Mod", "New", "Next", "Not", "Nothing", "Null", "Object", "On", "Option",
            "Optional", "Or", "ParamArray", "Preserve", "Private", "Property", "Public",
            "ReDim", "Resume", "Return", "Select", "Set", "Single", "Static", "Step",
            "Stop", "String", "Sub", "Then", "To", "True", "Type", "TypeOf", "Until",
            "Variant", "Wend", "While", "With", "Xor"
        };

        private static readonly HashSet<string> SolidWorksTypes = new HashSet<string>
        {
            "ISldWorks", "SldWorks", "IModelDoc2", "ModelDoc2", "IPartDoc", "PartDoc",
            "IAssemblyDoc", "AssemblyDoc", "IDrawingDoc", "DrawingDoc", "IBody2", "Body2",
            "IFace2", "Face2", "IEdge", "Edge", "IVertex", "Vertex", "IFeature", "Feature",
            "ISketch", "Sketch", "IComponent2", "Component2", "IMate2", "Mate2",
            "IConfiguration", "Configuration", "IView", "ISheet", "Sheet",
            "ISelectionMgr", "SelectionMgr", "IModelDocExtension", "ModelDocExtension"
        };

        // Pre-compiled regex patterns for better performance
        private static readonly Regex CSharpCommentSingle = new Regex(@"//.*$", RegexOptions.Multiline | RegexOptions.Compiled);
        private static readonly Regex CSharpCommentMulti = new Regex(@"/\*[\s\S]*?\*/", RegexOptions.Compiled);
        private static readonly Regex VBAComment = new Regex(@"'.*$", RegexOptions.Multiline | RegexOptions.Compiled);
        private static readonly Regex StringLiteral = new Regex(@"""[^""\\]*(?:\\.[^""\\]*)*""", RegexOptions.Compiled);
        private static readonly Regex VerbatimString = new Regex(@"@""[^""]*(?:""""[^""]*)*""", RegexOptions.Compiled);
        private static readonly Regex NumberLiteral = new Regex(@"\b\d+\.?\d*[fFdDmMlL]?\b", RegexOptions.Compiled);
        private static readonly Regex WordBoundary = new Regex(@"\b\w+\b", RegexOptions.Compiled);

        public class ColorScheme
        {
            public Color Background { get; set; }
            public Color Foreground { get; set; }
            public Color Keyword { get; set; }
            public Color Type { get; set; }
            public Color String { get; set; }
            public Color Comment { get; set; }
            public Color Number { get; set; }
            public Color SolidWorksType { get; set; }
        }

        public static ColorScheme DarkTheme => new ColorScheme
        {
            Background = Color.FromArgb(30, 30, 30),
            Foreground = Color.FromArgb(212, 212, 212),
            Keyword = Color.FromArgb(86, 156, 214),      // Blue
            Type = Color.FromArgb(78, 201, 176),          // Teal
            String = Color.FromArgb(206, 145, 120),       // Orange
            Comment = Color.FromArgb(106, 153, 85),       // Green
            Number = Color.FromArgb(181, 206, 168),       // Light green
            SolidWorksType = Color.FromArgb(184, 215, 163) // Bright green
        };

        public static ColorScheme LightTheme => new ColorScheme
        {
            Background = Color.FromArgb(255, 255, 255),
            Foreground = Color.FromArgb(30, 30, 30),
            Keyword = Color.FromArgb(0, 0, 255),
            Type = Color.FromArgb(43, 145, 175),
            String = Color.FromArgb(163, 21, 21),
            Comment = Color.FromArgb(0, 128, 0),
            Number = Color.FromArgb(9, 134, 88),
            SolidWorksType = Color.FromArgb(128, 0, 128)
        };

        /// <summary>
        /// Apply syntax highlighting with optimized performance
        /// </summary>
        public static void ApplyHighlighting(RichTextBox rtb, bool isCSharp, ColorScheme scheme)
        {
            if (rtb == null || rtb.IsDisposed) return;
            
            string text = rtb.Text;
            if (string.IsNullOrEmpty(text)) return;
            
            // Skip for large files to prevent lag
            if (text.Length > 15000) return;

            try
            {
                // Store current state
                int selStart = rtb.SelectionStart;
                int selLength = rtb.SelectionLength;

                // Disable updates
                rtb.SuspendLayout();
                SendMessage(rtb.Handle, WM_SETREDRAW, 0, 0);

                // Reset to default colors
                rtb.SelectAll();
                rtb.SelectionColor = scheme.Foreground;
                rtb.SelectionBackColor = scheme.Background;

                // Track regions to skip (comments and strings have priority)
                var skipRegions = new List<(int start, int end)>();

                // Apply comments first (highest priority)
                if (isCSharp)
                {
                    ApplyPattern(rtb, text, CSharpCommentSingle, scheme.Comment, skipRegions, true);
                    ApplyPattern(rtb, text, CSharpCommentMulti, scheme.Comment, skipRegions, true);
                }
                else
                {
                    ApplyPattern(rtb, text, VBAComment, scheme.Comment, skipRegions, true);
                }

                // Apply strings
                ApplyPattern(rtb, text, StringLiteral, scheme.String, skipRegions, true);
                if (isCSharp)
                {
                    ApplyPattern(rtb, text, VerbatimString, scheme.String, skipRegions, true);
                }

                // Apply numbers (skip if in comment/string)
                ApplyPattern(rtb, text, NumberLiteral, scheme.Number, skipRegions, false);

                // Apply keywords and types using single word scan
                var keywords = isCSharp ? CSharpKeywords : VBAKeywords;
                foreach (Match match in WordBoundary.Matches(text))
                {
                    if (IsInSkipRegion(match.Index, skipRegions)) continue;

                    string word = match.Value;
                    Color? color = null;

                    if (SolidWorksTypes.Contains(word))
                    {
                        color = scheme.SolidWorksType;
                    }
                    else if (keywords.Contains(word))
                    {
                        color = scheme.Keyword;
                    }

                    if (color.HasValue)
                    {
                        rtb.Select(match.Index, match.Length);
                        rtb.SelectionColor = color.Value;
                    }
                }

                // Restore selection
                rtb.Select(selStart, selLength);

                // Re-enable updates
                SendMessage(rtb.Handle, WM_SETREDRAW, 1, 0);
                rtb.ResumeLayout();
                rtb.Invalidate();
            }
            catch
            {
                // Silently fail - don't crash the app for syntax highlighting
                try
                {
                    SendMessage(rtb.Handle, WM_SETREDRAW, 1, 0);
                    rtb.ResumeLayout();
                }
                catch { }
            }
        }

        private static void ApplyPattern(RichTextBox rtb, string text, Regex pattern, Color color, 
            List<(int start, int end)> skipRegions, bool addToSkip)
        {
            foreach (Match match in pattern.Matches(text))
            {
                if (!addToSkip && IsInSkipRegion(match.Index, skipRegions)) continue;

                rtb.Select(match.Index, match.Length);
                rtb.SelectionColor = color;

                if (addToSkip)
                {
                    skipRegions.Add((match.Index, match.Index + match.Length));
                }
            }
        }

        private static bool IsInSkipRegion(int index, List<(int start, int end)> regions)
        {
            foreach (var (start, end) in regions)
            {
                if (index >= start && index < end) return true;
            }
            return false;
        }

        // Windows API for smooth updates
        private const int WM_SETREDRAW = 0x000B;

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern int SendMessage(IntPtr hWnd, int wMsg, int wParam, int lParam);
    }
}

