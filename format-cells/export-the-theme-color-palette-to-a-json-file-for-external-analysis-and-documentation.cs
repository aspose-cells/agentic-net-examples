using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

namespace AsposeCellsThemeExport
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Retrieve the 56‑entry palette colors
            Color[] paletteColors = workbook.Colors;

            // Convert palette colors to a list of hex strings for readability
            List<string> paletteHex = new List<string>();
            foreach (Color c in paletteColors)
            {
                paletteHex.Add(ColorToHex(c));
            }

            // Retrieve theme colors (Background1, Text1, ..., FollowedHyperlink)
            Dictionary<string, string> themeColors = new Dictionary<string, string>();
            foreach (ThemeColorType type in Enum.GetValues(typeof(ThemeColorType)))
            {
                // GetThemeColor returns a System.Drawing.Color
                Color themeColor = workbook.GetThemeColor(type);
                themeColors.Add(type.ToString(), ColorToHex(themeColor));
            }

            // Prepare an anonymous object that represents the export structure
            var exportObject = new
            {
                Palette = paletteHex,
                Theme = themeColors
            };

            // Serialize the object to JSON with indentation
            JsonSerializerOptions jsonOptions = new JsonSerializerOptions
            {
                WriteIndented = true
            };
            string json = JsonSerializer.Serialize(exportObject, jsonOptions);

            // Define output file path
            string outputPath = "ThemePalette.json";

            // Save JSON to file (lifecycle rule: save)
            File.WriteAllText(outputPath, json);

            Console.WriteLine($"Theme palette exported successfully to '{outputPath}'.");
        }

        // Helper method to convert a System.Drawing.Color to a hex string (e.g., "#FF112233")
        private static string ColorToHex(Color color)
        {
            // Include alpha channel for completeness
            return $"#{color.A:X2}{color.R:X2}{color.G:X2}{color.B:X2}";
        }
    }
}