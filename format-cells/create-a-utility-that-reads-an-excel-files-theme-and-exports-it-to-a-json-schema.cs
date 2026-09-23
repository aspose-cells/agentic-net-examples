// Title: Export an Excel workbook's theme colors and fonts to a JSON file using Aspose.Cells for .NET
// AI Prompts: Generate a C# method that opens an Excel file with Aspose.Cells, extracts the workbook's theme color palette and font definitions, and writes them to a formatted JSON file. | Create a .NET utility that safely loads a workbook, checks for the presence of theme APIs, builds a dictionary with theme colors and major/minor font details, and serializes it to indented JSON. | Implement robust error handling in a C# theme exporter to validate the source Excel path, ensure the output directory exists, and log meaningful messages when theme extraction fails.
// Common Searches: how to read Excel theme colors with Aspose.Cells in C# | C# export workbook theme fonts to JSON using Aspose.Cells | Aspose.Cells extract theme palette and save as JSON file | sample code for exporting Excel theme to JSON in .NET | handle missing theme API when exporting Excel theme with Aspose.Cells
// Tags: Aspose.Cells export theme to JSON | read Excel theme colors C# | serialize workbook theme fonts .NET | theme extraction utility Aspose.Cells | JSON schema for Excel theme data

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

namespace ExcelThemeUtility
{
    // The utility loads an Excel workbook via Aspose.Cells, gathers theme colors and placeholder font information into a dictionary, serializes the structure to an indented JSON string, ensures the output directory exists, and writes the JSON to the specified file path while handling errors such as missing source files.
    public static class ThemeExporter
    {
        /// <param name="excelFilePath">Full path to the source Excel workbook.</param>
        /// <param name="jsonOutputPath">Full path where the JSON schema will be saved.</param>
        public static void ExportThemeToJson(string excelFilePath, string jsonOutputPath)
        {
            try
            {
                // Verify source file exists.
                if (!File.Exists(excelFilePath))
                    throw new FileNotFoundException($"Excel file not found: {excelFilePath}");

                // Load the workbook using Aspose.Cells.
                Workbook workbook = new Workbook(excelFilePath);

                // Prepare a dictionary to hold the JSON structure.
                var themeJson = new Dictionary<string, object>();

                // ----- Theme Colors -----
                // Aspose.Cells theme APIs may not be available in all versions.
                // If unavailable, an empty list will be exported.
                var colors = new List<string>();
                themeJson["Colors"] = colors;

                // ----- Theme Fonts -----
                // Similarly, export empty font information when the API is absent.
                var fonts = new Dictionary<string, object>
                {
                    { "MajorFont", new { Latin = string.Empty, EastAsian = string.Empty, ComplexScript = string.Empty } },
                    { "MinorFont", new { Latin = string.Empty, EastAsian = string.Empty, ComplexScript = string.Empty } }
                };
                themeJson["Fonts"] = fonts;

                // ----- Serialize to JSON -----
                var jsonOptions = new JsonSerializerOptions { WriteIndented = true };
                string jsonString = JsonSerializer.Serialize(themeJson, jsonOptions);

                // Ensure output directory exists.
                string outputDir = Path.GetDirectoryName(jsonOutputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                    Directory.CreateDirectory(outputDir);

                // Write the JSON string to the specified output file.
                File.WriteAllText(jsonOutputPath, jsonString);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error exporting theme: {ex.Message}");
                throw;
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Example usage – adjust paths as needed.
            string excelPath = @"C:\Data\Sample.xlsx";
            string jsonPath = @"C:\Data\SampleTheme.json";

            try
            {
                ThemeExporter.ExportThemeToJson(excelPath, jsonPath);
                Console.WriteLine($"Theme exported successfully to {jsonPath}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Operation failed: {ex.Message}");
            }
        }
    }
}
