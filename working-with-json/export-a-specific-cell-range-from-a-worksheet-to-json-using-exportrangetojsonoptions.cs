// Title: Export a Worksheet Range to Pretty‑Printed JSON with Aspose.Cells (C#)
// Description: Shows how to build a workbook, populate cells A1:B3, create a range, configure ExportRangeToJsonOptions (header row, string output, 4‑space indentation), and call JsonUtility.ExportRangeToJson to obtain a formatted JSON string that is written to the console and saved as a file.
// Keywords: Aspose.Cells | ExportRangeToJsonOptions | C# Excel to JSON | worksheet range export | pretty printed JSON | header row JSON | string values JSON | JsonUtility ExportRangeToJson | sample code | Excel range to JSON C#
// Common Searches: export selected Excel range to JSON using Aspose.Cells C# | how to use ExportRangeToJsonOptions for pretty JSON | Aspose.Cells export range with headers as JSON | C# code to convert worksheet cells to formatted JSON | save Excel range as JSON file Aspose
// Developer Intent: Create a JSON representation of a specific Excel range with optional headers and custom indentation using Aspose.Cells.
// Use Cases: Convert a lookup table in a spreadsheet into a JSON configuration file. | Provide front‑end applications with data extracted from a defined worksheet area. | Generate API payloads directly from Excel sections without manual transformation.
// AI Prompts: Write C# code that uses Aspose.Cells to export the range A1:C10 to a JSON string with headers and 2‑space indentation. | Show how to export a worksheet range to JSON, force all values to strings, and save the result to a file using ExportRangeToJsonOptions. | Explain how to modify ExportRangeToJsonOptions to omit the header row and use tab characters for indentation.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsJsonExport
{
    // Alias to avoid conflict with System.Range (C# 8+)
    using AsposeRange = Aspose.Cells.Range;

    // Shows how to build a workbook, populate cells A1:B3, create a range, configure ExportRangeToJsonOptions (header row, string output, 4‑space indentation), and call JsonUtility.ExportRangeToJson to obtain a formatted JSON string that is written to the console and saved as a file.
    public class ExportRangeToJsonDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data in the worksheet
                worksheet.Cells["A1"].PutValue("Name");
                worksheet.Cells["B1"].PutValue("Age");
                worksheet.Cells["A2"].PutValue("John");
                worksheet.Cells["B2"].PutValue(30);
                worksheet.Cells["A3"].PutValue("Alice");
                worksheet.Cells["B3"].PutValue(25);

                // Define the range that we want to export (A1:B3)
                AsposeRange exportRange = worksheet.Cells.CreateRange("A1:B3");

                // Configure export options
                ExportRangeToJsonOptions options = new ExportRangeToJsonOptions
                {
                    HasHeaderRow = true,      // First row contains column headers
                    ExportAsString = true,    // Export all values as strings
                    Indent = "    "           // Use 4 spaces for pretty‑printed JSON
                };

                // Export the range to a JSON string
                string jsonResult = JsonUtility.ExportRangeToJson(exportRange, options);

                // Output the JSON to console
                Console.WriteLine("Exported JSON:");
                Console.WriteLine(jsonResult);

                // Write the JSON to a file (ensure the directory exists)
                string outputPath = "ExportedRange.json";
                File.WriteAllText(outputPath, jsonResult);
                Console.WriteLine($"JSON written to: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ExportRangeToJsonDemo.Run();
        }
    }
}
