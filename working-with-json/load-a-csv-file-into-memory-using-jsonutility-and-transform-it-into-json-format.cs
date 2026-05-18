using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsCsvToJson
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the source CSV file
                string csvPath = "input.csv";

                // Verify that the CSV file exists
                if (!File.Exists(csvPath))
                {
                    Console.Error.WriteLine($"CSV file not found: {csvPath}");
                    return;
                }

                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Import CSV data into the worksheet starting at cell A1 (row 0, column 0)
                // Using comma as delimiter and converting numeric strings to numbers
                cells.ImportCSV(csvPath, ",", true, 0, 0);

                // Determine the used range of the worksheet
                // MaxDisplayRange returns the smallest range that contains all non‑empty cells
                Aspose.Cells.Range usedRange = cells.MaxDisplayRange;

                // Configure JSON export options (customize as needed)
                JsonSaveOptions jsonOptions = new JsonSaveOptions
                {
                    ExportNestedStructure = true,
                    SkipEmptyRows = true,
                    ExportEmptyCells = true,
                    HasHeaderRow = true
                };

                // Export the used range to a JSON string
                string jsonOutput = JsonUtility.ExportRangeToJson(usedRange, jsonOptions);

                // Output the resulting JSON
                Console.WriteLine(jsonOutput);
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}