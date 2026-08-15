// Title: Export Excel Workbook Theme & Colors to JSON with Aspose.Cells for .NET (C#)
// Description: C# utility that opens an Excel workbook via Aspose.Cells, extracts the workbook's theme name and every ThemeColorType value, converts each color to a hex string, and writes a pretty‑printed JSON schema. Handles missing files, password‑protected workbooks, and creates the output folder automatically.
// Keywords: Aspose.Cells theme export | C# export Excel theme to JSON | read Excel theme colors | ThemeColorType Aspose.Cells | save workbook theme as JSON | .NET Excel theme extraction | JSON schema from Excel theme | password protected workbook handling | Aspose.Cells code example | GitHub Aspose.Cells ThemeExporter
// Common Searches: How to extract Excel theme colors using Aspose.Cells C# | Export workbook theme to JSON with Aspose.Cells for .NET | Get theme name and palette from .xlsx programmatically | Save Excel theme information as JSON file in C# | Aspose.Cells ThemeExporter example
// Developer Intent: Read an Excel workbook’s theme and all associated colors with Aspose.Cells and serialize the information to a JSON file.
// Use Cases: Create a design‑system palette by converting Excel theme colors to a JSON file for CSS or UI libraries. | Compare visual styles of multiple workbooks by exporting each theme to JSON and performing a diff. | Automate documentation of workbook styling standards by generating a JSON schema of theme metadata.
// AI Prompts: Write C# code that loads an Excel file with Aspose.Cells, retrieves the theme name and ThemeColorType colors, and outputs a formatted JSON file. | Explain how to detect and handle password‑protected Excel workbooks when exporting theme data using Aspose.Cells. | Generate unit tests for ThemeExporter.ExportThemeToJson covering scenarios: missing file, password protection, and successful export.

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

namespace AsposeCellsThemeExporter
{
    // C# utility that opens an Excel workbook via Aspose.Cells, extracts the workbook's theme name and every ThemeColorType value, converts each color to a hex string, and writes a pretty‑printed JSON schema. Handles missing files, password‑protected workbooks, and creates the output folder automatically.
    public static class ThemeExporter
    {
        /// <param name="excelPath">Full path to the source Excel workbook.</param>
        /// <param name="jsonPath">Full path where the JSON schema will be saved.</param>
        public static void ExportThemeToJson(string excelPath, string jsonPath)
        {
            try
            {
                // Verify that the source Excel file exists
                if (!File.Exists(excelPath))
                {
                    Console.Error.WriteLine($"Error: The file '{excelPath}' was not found.");
                    return;
                }

                // Load the workbook (throws CellsException if the file is password‑protected)
                Workbook workbook = new Workbook(excelPath);

                // Retrieve the theme name
                string themeName = workbook.Theme;

                // Collect all theme colors
                var colors = new Dictionary<string, string>();
                foreach (ThemeColorType colorType in Enum.GetValues(typeof(ThemeColorType)))
                {
                    Color color = workbook.GetThemeColor(colorType);
                    string hex = $"#{color.R:X2}{color.G:X2}{color.B:X2}";
                    colors[colorType.ToString()] = hex;
                }

                // Build an anonymous object representing the JSON schema
                var themeSchema = new
                {
                    Theme = themeName,
                    Colors = colors
                };

                // Serialize the object to formatted JSON
                string json = JsonSerializer.Serialize(themeSchema, new JsonSerializerOptions
                {
                    WriteIndented = true
                });

                // Ensure the output directory exists
                string? outputDir = Path.GetDirectoryName(jsonPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Write the JSON to the specified file
                File.WriteAllText(jsonPath, json);
                Console.WriteLine($"Theme exported successfully to '{jsonPath}'.");
            }
            catch (CellsException ex)
            {
                // Handle password‑protected files or other Aspose.Cells specific errors
                if (ex.Message != null && ex.Message.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    Console.Error.WriteLine("Error: The workbook is password protected. Provide a valid password to load it.");
                }
                else
                {
                    Console.Error.WriteLine($"Aspose.Cells error: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }
    }

    // Example usage
    class Program
    {
        static void Main()
        {
            // Path to the Excel file whose theme will be exported
            string sourceExcel = "SampleWorkbook.xlsx";

            // Destination path for the generated JSON schema
            string outputJson = "WorkbookThemeSchema.json";

            // Export the theme information
            ThemeExporter.ExportThemeToJson(sourceExcel, outputJson);
        }
    }
}
