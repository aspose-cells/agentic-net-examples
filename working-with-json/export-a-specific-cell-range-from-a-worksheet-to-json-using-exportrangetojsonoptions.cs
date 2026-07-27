using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsJsonExport
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook in memory
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data (including a header row)
                worksheet.Cells["A1"].PutValue("Name");
                worksheet.Cells["B1"].PutValue("Age");
                worksheet.Cells["A2"].PutValue("John");
                worksheet.Cells["B2"].PutValue(30);
                worksheet.Cells["A3"].PutValue("Alice");
                worksheet.Cells["B3"].PutValue(25);

                // Define the range to export (A1:B3)
                // Use fully qualified type to avoid ambiguity with System.Range
                Aspose.Cells.Range exportRange = worksheet.Cells.CreateRange("A1:B3");

                // Configure export options
                ExportRangeToJsonOptions jsonOptions = new ExportRangeToJsonOptions
                {
                    HasHeaderRow = true,      // Treat first row as header
                    ExportAsString = true,    // Export cell values as strings
                    ExportEmptyCells = false, // Do not include empty cells
                    Indent = "    "           // 4‑space indentation for pretty JSON
                };

                // Export the range to a JSON string
                string jsonResult = JsonUtility.ExportRangeToJson(exportRange, jsonOptions);

                // Output JSON to console
                Console.WriteLine("Exported JSON:");
                Console.WriteLine(jsonResult);

                // Write JSON to a file (ensure the directory exists)
                string outputPath = Path.Combine(Environment.CurrentDirectory, "ExportedRange.json");
                File.WriteAllText(outputPath, jsonResult);
                Console.WriteLine($"\nJSON saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}