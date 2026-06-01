using System;
using Aspose.Cells;
using Aspose.Cells.Json;

namespace AsposeCellsJsonExport
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data (including header row)
            cells["A1"].PutValue("ID");
            cells["B1"].PutValue("Name");
            cells["C1"].PutValue("Score");

            cells["A2"].PutValue(1);
            cells["B2"].PutValue("Alice");
            cells["C2"].PutValue(85);

            cells["A3"].PutValue(2);
            cells["B3"].PutValue("Bob");
            cells["C3"].PutValue(92);

            cells["A4"].PutValue(3);
            cells["B4"].PutValue("Charlie");
            cells["C4"].PutValue(78);

            // Hide a row (row index 2 -> third row) and a column (column index 1 -> column B)
            sheet.Cells.HideRow(2);      // Hides the row containing Bob's data
            sheet.Cells.HideColumn(1);   // Hides the "Name" column

            // Configure JsonSaveOptions
            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                // Skip empty rows (helps to ignore rows that become empty after hiding)
                SkipEmptyRows = true,

                // Do not export empty cells as null
                ExportEmptyCells = false,

                // Define the export area to cover the used range.
                // Hidden rows/columns are still part of the range, but because we set
                // SkipEmptyRows = true and ExportEmptyCells = false, cells in hidden rows/columns
                // will not appear in the resulting JSON.
                ExportArea = new CellArea
                {
                    StartRow = 0,
                    EndRow = cells.MaxDataRow,
                    StartColumn = 0,
                    EndColumn = cells.MaxDataColumn
                },

                // Indent the JSON for readability
                Indent = "  "
            };

            // Save the workbook as JSON using the configured options
            string outputPath = "ExportedData.json";
            workbook.Save(outputPath, jsonOptions);

            Console.WriteLine($"Workbook exported to JSON at: {outputPath}");
        }
    }
}