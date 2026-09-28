// Title: Export only non‑empty cells from an Excel worksheet to a JSON array with headers using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that uses Aspose.Cells to export the used range of a worksheet to a JSON array, omitting empty cells and treating the first row as column headers. | Show how to configure ExportRangeToJsonOptions to include empty cells as null values while preserving data types in the JSON output. | Create a reusable C# method that accepts a Worksheet object and returns a JSON string of its populated cells with header mapping, using Aspose.Cells.
// Common Searches: asp.net export excel worksheet to json without empty cells using aspose.cells | c# aspose.cells export used range to json array with header row | how to skip blank cells when converting Excel to JSON with Aspose.Cells | export populated cells from workbook to json string aspose.cells .net | json export options for aspose.cells to ignore empty cells
// Tags: Aspose.Cells ExportRangeToJsonOptions configuration | skip empty cells JSON export Aspose.Cells | export used range to JSON C# | header row mapping in JSON output Aspose.Cells | convert worksheet range to JSON string .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Utility;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsExportValidatedJson
{
    // Demonstrates creating a workbook, defining its used range, and using Aspose.Cells ExportRangeToJsonOptions to export only populated cells to a JSON array with the first row treated as headers.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate sample data (some cells are left empty intentionally)
                cells["A1"].PutValue("Name");
                cells["B1"].PutValue("Age");
                cells["C1"].PutValue("City");

                cells["A2"].PutValue("John");
                cells["B2"].PutValue(30);
                // C2 left empty

                cells["A3"].PutValue("Alice");
                // B3 left empty
                cells["C3"].PutValue("London");

                // Determine the used range (from first row/column to the last populated cell)
                int lastRow = cells.MaxDataRow;      // zero‑based index of the last row with data
                int lastCol = cells.MaxDataColumn;   // zero‑based index of the last column with data

                // Create a range that covers the used area
                AsposeRange usedRange = cells.CreateRange(0, 0, lastRow + 1, lastCol + 1);

                // Configure export options:
                // - ExportEmptyCells = false  => skip empty cells (only validated/non‑empty cells are exported)
                // - HasHeaderRow = true       => first row is treated as header
                ExportRangeToJsonOptions jsonOptions = new ExportRangeToJsonOptions
                {
                    ExportEmptyCells = false,
                    HasHeaderRow = true,
                    ExportAsString = false   // keep original data types
                };

                // Export the range to a JSON string
                string jsonResult = JsonUtility.ExportRangeToJson(usedRange, jsonOptions);

                // Output the JSON for downstream processing
                Console.WriteLine(jsonResult);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
