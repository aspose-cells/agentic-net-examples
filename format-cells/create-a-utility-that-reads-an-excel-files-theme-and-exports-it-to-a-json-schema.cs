using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

namespace AsposeCellsThemeExport
{
    /// <summary>
    /// Utility class that extracts the theme information from an Excel workbook
    /// and exports it as a JSON schema file.
    /// </summary>
    public static class ThemeExporter
    {
        /// <summary>
        /// Reads the theme from the specified Excel file and writes a JSON file
        /// containing the theme name and its standard colors.
        /// </summary>
        /// <param name="excelPath">Full path to the source Excel file.</param>
        /// <param name="jsonPath">Full path where the JSON schema will be saved.</param>
        public static void ExportThemeToJson(string excelPath, string jsonPath)
        {
            try
            {
                // Verify that the source Excel file exists
                if (!File.Exists(excelPath))
                {
                    Console.Error.WriteLine($"Error: Excel file not found – '{excelPath}'.");
                    return;
                }

                // Load the workbook using Aspose.Cells
                Workbook workbook = new Workbook(excelPath);

                // Retrieve the theme name
                string themeName = workbook.Theme;

                // Define the theme color types to extract
                ThemeColorType[] colorTypes = new ThemeColorType[]
                {
                    ThemeColorType.Accent1,
                    ThemeColorType.Accent2,
                    ThemeColorType.Accent3,
                    ThemeColorType.Accent4,
                    ThemeColorType.Accent5,
                    ThemeColorType.Accent6,
                    ThemeColorType.Hyperlink,
                    ThemeColorType.FollowedHyperlink,
                    ThemeColorType.Background1,
                    ThemeColorType.Text1,
                    ThemeColorType.Background2,
                    ThemeColorType.Text2
                };

                // Collect colors as hex strings
                var colors = new Dictionary<string, string>();
                foreach (ThemeColorType ct in colorTypes)
                {
                    Color c = workbook.GetThemeColor(ct);
                    colors[ct.ToString()] = $"#{c.R:X2}{c.G:X2}{c.B:X2}";
                }

                // Build an object representing the JSON schema
                var themeInfo = new
                {
                    Theme = themeName,
                    Colors = colors
                };

                // Serialize the object to formatted JSON
                var jsonOptions = new JsonSerializerOptions { WriteIndented = true };
                string json = JsonSerializer.Serialize(themeInfo, jsonOptions);

                // Ensure the destination directory exists
                string jsonDir = Path.GetDirectoryName(jsonPath);
                if (!string.IsNullOrEmpty(jsonDir) && !Directory.Exists(jsonDir))
                {
                    Directory.CreateDirectory(jsonDir);
                }

                // Write the JSON to the specified file
                File.WriteAllText(jsonPath, json);
                Console.WriteLine($"Theme exported successfully to '{jsonPath}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred while exporting the theme: {ex.Message}");
            }
        }
    }

    // Example usage
    class Program
    {
        static void Main()
        {
            try
            {
                string excelFile = "SampleWithTheme.xlsx";   // source Excel file
                string jsonFile = "ThemeSchema.json";        // destination JSON file

                ThemeExporter.ExportThemeToJson(excelFile, jsonFile);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}