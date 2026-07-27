using System;
using Aspose.Cells;
using Aspose.Cells.Utility;
using AsposeRange = Aspose.Cells.Range;

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

                // Populate sample data with a header row
                cells["A1"].PutValue("Name");
                cells["B1"].PutValue("Age");
                cells["A2"].PutValue("John");
                cells["B2"].PutValue(30);
                cells["A3"].PutValue("Alice");
                cells["B3"].PutValue(25);

                // Define the range that includes the header and data rows
                AsposeRange range = cells.CreateRange("A1:B3");

                // Configure export options to treat the first row as header (column names as keys)
                ExportRangeToJsonOptions jsonOptions = new ExportRangeToJsonOptions
                {
                    HasHeaderRow = true,          // Use column names as keys
                    ExportEmptyCells = true,      // Include empty cells as null
                    Indent = "    "               // Pretty‑print with indentation
                };

                // Export the range to JSON using the configured options
                string jsonOutput = JsonUtility.ExportRangeToJson(range, jsonOptions);

                // Output the resulting JSON
                Console.WriteLine("JSON output with column names as keys:");
                Console.WriteLine(jsonOutput);
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}