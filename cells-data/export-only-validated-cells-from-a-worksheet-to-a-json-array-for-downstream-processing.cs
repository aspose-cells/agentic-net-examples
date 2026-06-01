using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

class Program
{
    static void Main(string[] args)
    {
        try
        {
            Workbook workbook;

            // If a template file path is supplied, load it only when the file exists.
            if (args.Length > 0 && File.Exists(args[0]))
            {
                workbook = new Workbook(args[0]);
            }
            else
            {
                // Create a new workbook when no valid template is provided.
                workbook = new Workbook();
            }

            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate sample data (some cells are intentionally left empty)
            cells["A1"].PutValue("Name");
            cells["B1"].PutValue("Age");
            cells["A2"].PutValue("John");
            cells["B2"].PutValue(30);
            cells["A3"].PutValue("Alice");
            // B3 is left empty to demonstrate skipping empty cells

            // Determine the used area of the sheet
            int lastRow = cells.MaxDataRow;          // zero‑based index of the last row with data
            int lastColumn = cells.MaxDataColumn;    // zero‑based index of the last column with data

            // Create a range that covers the used area (including header row)
            Aspose.Cells.Range exportRange = cells.CreateRange(0, 0, lastRow + 1, lastColumn + 1);

            // Configure JSON export options
            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                ExportEmptyCells = false,
                HasHeaderRow = true,
                ExportAsString = false
            };

            // Export the range to a JSON string
            string jsonResult = JsonUtility.ExportRangeToJson(exportRange, jsonOptions);

            // Output the JSON for downstream processing
            Console.WriteLine(jsonResult);
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}