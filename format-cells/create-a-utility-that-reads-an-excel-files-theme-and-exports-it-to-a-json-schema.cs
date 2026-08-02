// Title: Export Excel Workbook Theme Colors to JSON using Aspose.Cells for .NET (C#)
// Description: This C# example shows how to load an .xlsx file with Aspose.Cells, read the workbook's Theme name and all ThemeColorType values, convert each System.Drawing.Color to a hexadecimal string, and serialize the data into a nicely formatted JSON schema. The utility also creates the output folder when needed and handles missing files gracefully.
// Keywords: Aspose.Cells | C# | .NET | Excel theme extraction | ThemeColorType | JSON serialization | color palette export | Workbook.Theme | GetThemeColor | style guide generation | branding consistency | Excel API
// Common Searches: Aspose.Cells export theme to JSON | C# read Excel theme colors | How to get workbook theme name with Aspose.Cells | Convert Excel theme colors to hex | Save Excel theme as JSON file | Extract color palette from .xlsx using .NET
// Developer Intent: Retrieve the theme name and its color palette from an Excel workbook and write the data to a JSON file.
// Use Cases: Create a reusable JSON color palette for web or mobile UI that matches an Excel design template. | Audit multiple workbooks to verify that corporate brand colors are applied consistently. | Feed the exported JSON into a theming engine for automated report generation. | Document Excel style guides by exporting theme details for design teams.
// AI Prompts: Generate a method that returns a Dictionary<string,string> where keys are ThemeColorType names and values are hex color codes from a given Excel file using Aspose.Cells. | Extend the exporter to include the workbook's major and minor font names in the JSON output. | Add logging that records any ThemeColorType that returns an empty or default color and continues processing. | Create a PowerShell wrapper that calls the C# utility and uploads the resulting JSON to a SharePoint library.

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

namespace AsposeCellsThemeExport
{
    // Simple DTO to hold theme information for JSON serialization
    // This C# example shows how to load an .xlsx file with Aspose.Cells, read the workbook's Theme name and all ThemeColorType values, convert each System.Drawing.Color to a hexadecimal string, and serialize the data into a nicely formatted JSON schema. The utility also creates the output folder when needed and handles missing files gracefully.
    public class ThemeInfo
    {
        public string Name { get; set; }
        public Dictionary<string, string> Colors { get; set; }
    }

    public class ThemeExporter
    {
        // Exports the theme of the given Excel file to a JSON file
        public static void ExportThemeToJson(string excelPath, string jsonOutputPath)
        {
            try
            {
                // Verify that the source Excel file exists
                if (!File.Exists(excelPath))
                {
                    Console.Error.WriteLine($"Error: Excel file not found at '{excelPath}'.");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(excelPath);

                // Retrieve the theme name
                string themeName = workbook.Theme;

                // Collect theme colors
                var colors = new Dictionary<string, string>();
                ThemeColorType[] themeColorTypes = new ThemeColorType[]
                {
                    ThemeColorType.Accent1,
                    ThemeColorType.Accent2,
                    ThemeColorType.Accent3,
                    ThemeColorType.Accent4,
                    ThemeColorType.Accent5,
                    ThemeColorType.Accent6,
                    ThemeColorType.Hyperlink,
                    ThemeColorType.FollowedHyperlink,
                    ThemeColorType.Text1,
                    ThemeColorType.Text2,
                    ThemeColorType.Background1,
                    ThemeColorType.Background2
                };

                foreach (var type in themeColorTypes)
                {
                    Color color = workbook.GetThemeColor(type);
                    // Convert color to hex string (e.g., #RRGGBB)
                    string hex = $"#{color.R:X2}{color.G:X2}{color.B:X2}";
                    colors[type.ToString()] = hex;
                }

                // Build the DTO
                ThemeInfo themeInfo = new ThemeInfo
                {
                    Name = themeName,
                    Colors = colors
                };

                // Serialize to JSON (indented for readability)
                var jsonOptions = new JsonSerializerOptions { WriteIndented = true };
                string json = JsonSerializer.Serialize(themeInfo, jsonOptions);

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(jsonOutputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Write JSON to the specified file
                File.WriteAllText(jsonOutputPath, json);
                Console.WriteLine($"Theme exported to '{jsonOutputPath}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred while exporting the theme: {ex.Message}");
            }
        }

        // Example usage
        public static void Main()
        {
            // Path to the source Excel file
            string sourceExcel = "SampleWithTheme.xlsx";

            // Path where the JSON schema will be saved
            string outputJson = "ThemeSchema.json";

            ExportThemeToJson(sourceExcel, outputJson);
        }
    }
}
