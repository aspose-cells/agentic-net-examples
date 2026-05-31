using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsJsonExport
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Populate sample data with a header row (column names)
                cells["A1"].PutValue("Name");
                cells["B1"].PutValue("Age");
                cells["C1"].PutValue("Country");

                cells["A2"].PutValue("John");
                cells["B2"].PutValue(30);
                cells["C2"].PutValue("USA");

                cells["A3"].PutValue("Alice");
                cells["B3"].PutValue(25);
                cells["C3"].PutValue("Canada");

                // Define the range that includes the header and data rows
                // Use fully qualified type to avoid conflict with System.Range
                Aspose.Cells.Range exportRange = cells.CreateRange("A1:C3");

                // Configure export options to treat the first row as header (column names)
                ExportRangeToJsonOptions jsonOptions = new ExportRangeToJsonOptions
                {
                    HasHeaderRow = true,      // Include column names as keys in the JSON output
                    ExportEmptyCells = false, // Do not export empty cells as null
                    Indent = "    "           // Pretty‑print with indentation
                };

                // Export the range to JSON string
                string jsonOutput = JsonUtility.ExportRangeToJson(exportRange, jsonOptions);

                // Output the JSON to console
                Console.WriteLine("Exported JSON with column names as keys:");
                Console.WriteLine(jsonOutput);

                // Optionally, save the JSON to a file
                string jsonFilePath = "ExportedData.json";
                File.WriteAllText(jsonFilePath, jsonOutput);
                Console.WriteLine($"JSON saved to '{Path.GetFullPath(jsonFilePath)}'");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}