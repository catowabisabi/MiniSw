using System;
using System.Drawing;
using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace MiniSolidworkAutomator.Models
{
    public class ThemeColors
    {
        public string DarkBackground { get; set; } = "#1E1E1E";
        public string DarkPanel { get; set; } = "#2D2D2D";
        public string DarkToolbar { get; set; } = "#263238";
        public string DarkSplitter { get; set; } = "#37474F";
        public string DarkTerminal { get; set; } = "#1E1E1E";
        public string TextWhite { get; set; } = "#FFFFFF";
        public string TextGray { get; set; } = "#B4B4B4";
        public string AccentGreen { get; set; } = "#2E7D32";
        public string AccentRed { get; set; } = "#B71C1C";
        public string AccentBlue { get; set; } = "#2196F3";
        public string AccentPurple { get; set; } = "#8E44AD";
        public string EditorForeground { get; set; } = "#D4D4D4";
        public string TerminalForeground { get; set; } = "#CCCCCC";
    }

    public class ThemeConfig
    {
        public string Name { get; set; } = "Dark Theme";
        public ThemeColors Colors { get; set; } = new ThemeColors();

        private static string ThemePath => Path.Combine(
            AppDomain.CurrentDomain.BaseDirectory,
            "theme.json"
        );

        public static ThemeConfig Load()
        {
            try
            {
                if (File.Exists(ThemePath))
                {
                    var json = File.ReadAllText(ThemePath);
                    return JsonSerializer.Deserialize<ThemeConfig>(json) ?? new ThemeConfig();
                }
            }
            catch { }

            // Create default theme file if it doesn't exist
            var defaultTheme = new ThemeConfig();
            defaultTheme.Save();
            return defaultTheme;
        }

        public void Save()
        {
            try
            {
                var options = new JsonSerializerOptions
                {
                    WriteIndented = true
                };
                var json = JsonSerializer.Serialize(this, options);
                File.WriteAllText(ThemePath, json);
            }
            catch { }
        }

        // Helper methods to convert hex to Color
        public static Color HexToColor(string hex)
        {
            try
            {
                hex = hex.TrimStart('#');
                if (hex.Length == 6)
                {
                    return Color.FromArgb(
                        Convert.ToInt32(hex.Substring(0, 2), 16),
                        Convert.ToInt32(hex.Substring(2, 2), 16),
                        Convert.ToInt32(hex.Substring(4, 2), 16)
                    );
                }
                else if (hex.Length == 8)
                {
                    return Color.FromArgb(
                        Convert.ToInt32(hex.Substring(0, 2), 16),
                        Convert.ToInt32(hex.Substring(2, 2), 16),
                        Convert.ToInt32(hex.Substring(4, 2), 16),
                        Convert.ToInt32(hex.Substring(6, 2), 16)
                    );
                }
            }
            catch { }
            return Color.Black;
        }

        public static string ColorToHex(Color color)
        {
            return $"#{color.R:X2}{color.G:X2}{color.B:X2}";
        }

        // Get colors as Color objects
        [JsonIgnore]
        public Color DarkBackgroundColor => HexToColor(Colors.DarkBackground);
        [JsonIgnore]
        public Color DarkPanelColor => HexToColor(Colors.DarkPanel);
        [JsonIgnore]
        public Color DarkToolbarColor => HexToColor(Colors.DarkToolbar);
        [JsonIgnore]
        public Color DarkSplitterColor => HexToColor(Colors.DarkSplitter);
        [JsonIgnore]
        public Color DarkTerminalColor => HexToColor(Colors.DarkTerminal);
        [JsonIgnore]
        public Color TextWhiteColor => HexToColor(Colors.TextWhite);
        [JsonIgnore]
        public Color TextGrayColor => HexToColor(Colors.TextGray);
        [JsonIgnore]
        public Color AccentGreenColor => HexToColor(Colors.AccentGreen);
        [JsonIgnore]
        public Color AccentRedColor => HexToColor(Colors.AccentRed);
        [JsonIgnore]
        public Color AccentBlueColor => HexToColor(Colors.AccentBlue);
        [JsonIgnore]
        public Color AccentPurpleColor => HexToColor(Colors.AccentPurple);
        [JsonIgnore]
        public Color EditorForegroundColor => HexToColor(Colors.EditorForeground);
        [JsonIgnore]
        public Color TerminalForegroundColor => HexToColor(Colors.TerminalForeground);
    }
}
