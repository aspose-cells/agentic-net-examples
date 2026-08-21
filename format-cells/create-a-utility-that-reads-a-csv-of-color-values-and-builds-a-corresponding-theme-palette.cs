// Title: C# Utility: Build a Custom Excel Theme from a CSV Palette with Aspose.Cells
// Description: A C# helper that reads a CSV file containing exactly 12 color values (RGB or hex), converts them to System.Drawing.Color objects, creates a new Workbook, applies the colors as a custom theme using Aspose.Cells.CustomTheme, demonstrates the theme on cell A1, and saves the file as an XLSX workbook. If the CSV is missing, a default 12‑color set is generated automatically.
// Keywords: Aspose.Cells custom theme C# | CSV color palette Excel | generate Excel theme from CSV | C# read RGB hex colors | apply custom theme Aspose.Cells | Excel workbook theme palette | SaveFormat.Xlsx example | dotnet Excel theme utility | GitHub Aspose.Cells sample | theme colors array
// Common Searches: how to create an Excel theme from a CSV using Aspose.Cells | C# read hex and RGB colors from a file for Excel theme | apply custom color palette to workbook with Aspose.Cells .NET | Aspose.Cells example for custom theme palette | generate corporate Excel theme programmatically
// Developer Intent: Generate a custom Excel theme from a CSV list of 12 colors and apply it to a workbook using Aspose.Cells in C#.
// Use Cases: Automate branding by storing corporate colors in a CSV and applying them as a reusable Excel theme. | Create a template workbook that showcases the new theme on sample cells for quick visual verification. | Ensure robustness by providing a fallback CSV with default colors when the user‑specified file is absent.
// AI Prompts: Write a method that validates hex and RGB strings from a CSV and returns a System.Drawing.Color[] for Aspose.Cells.CustomTheme. | Show how to style multiple cells with different ThemeColor types (Accent2, Text1, etc.) after setting the custom theme. | Explain how to load an existing workbook and replace its current theme with the CSV‑derived palette using Aspose.Cells.

using System;
using System.Drawing;
using System.IO;
using System.Linq;
using Aspose.Cells;

namespace ThemePaletteUtility
{
    // A C# helper that reads a CSV file containing exactly 12 color values (RGB or hex), converts them to System.Drawing.Color objects, creates a new Workbook, applies the colors as a custom theme using Aspose.Cells.CustomTheme, demonstrates the theme on cell A1, and saves the file as an XLSX workbook. If the CSV is missing, a default 12‑color set is generated automatically.
    public static class ThemeBuilder
    {
        /// <param name="csvPath">Path to the CSV file. Each line should contain a color in one of the following formats:
        ///   - R,G,B (e.g., 255,0,0)
        ///   - Hex string with or without leading '#', e.g., #FF0000 or FF0000</param>
        /// <param name="outputPath">Path where the resulting workbook will be saved.</param>
        /// <param name="themeName">Name of the custom theme to apply.</param>
        public static void ApplyCustomThemeFromCsv(string csvPath, string outputPath, string themeName)
        {
            // Ensure the CSV file exists; if not, create a default one with 12 colors.
            if (!File.Exists(csvPath))
            {
                string[] defaultColors =
                {
                    "255,0,0",   // Red
                    "0,255,0",   // Green
                    "0,0,255",   // Blue
                    "255,255,0", // Yellow
                    "255,0,255", // Magenta
                    "0,255,255", // Cyan
                    "#808080",   // Gray
                    "#800000",   // Maroon
                    "#008000",   // DarkGreen
                    "#000080",   // Navy
                    "#808000",   // Olive
                    "#800080"    // Purple
                };
                File.WriteAllLines(csvPath, defaultColors);
            }

            // Read all non‑empty lines
            string[] lines = File.ReadAllLines(csvPath)
                                 .Select(l => l.Trim())
                                 .Where(l => !string.IsNullOrEmpty(l))
                                 .ToArray();

            if (lines.Length != 12)
                throw new InvalidOperationException($"A theme requires exactly 12 colors, but {lines.Length} were provided.");

            // Parse each line into a System.Drawing.Color
            Color[] themeColors = new Color[12];
            for (int i = 0; i < 12; i++)
            {
                string line = lines[i];

                // Try CSV format: R,G,B
                if (line.Contains(","))
                {
                    string[] parts = line.Split(',');
                    if (parts.Length != 3)
                        throw new FormatException($"Invalid RGB format on line {i + 1}: {line}");

                    int r = int.Parse(parts[0].Trim());
                    int g = int.Parse(parts[1].Trim());
                    int b = int.Parse(parts[2].Trim());

                    themeColors[i] = Color.FromArgb(255, r, g, b);
                }
                else
                {
                    // Assume hex format
                    string hex = line.StartsWith("#") ? line : "#" + line;
                    themeColors[i] = ColorTranslator.FromHtml(hex);
                }
            }

            // Create a new workbook
            Workbook workbook = new Workbook();

            // Apply the custom theme
            workbook.CustomTheme(themeName, themeColors);

            // Demonstrate the theme by writing a sample cell
            Worksheet ws = workbook.Worksheets[0];
            Cell demoCell = ws.Cells["A1"];
            demoCell.PutValue($"Theme: {themeName}");
            Style demoStyle = workbook.CreateStyle();
            demoStyle.Font.ThemeColor = new ThemeColor(ThemeColorType.Accent1, 0.0);
            demoStyle.Font.Size = 14;
            demoCell.SetStyle(demoStyle);

            // Save the workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
        }
    }

    // Example usage
    class Program
    {
        static void Main()
        {
            try
            {
                string csvFile = "themeColors.csv";          // CSV with 12 color definitions
                string outputFile = "CustomThemeWorkbook.xlsx";
                string themeName = "MyCsvTheme";

                ThemeBuilder.ApplyCustomThemeFromCsv(csvFile, outputFile, themeName);

                Console.WriteLine($"Workbook saved to {outputFile} with theme '{themeName}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
