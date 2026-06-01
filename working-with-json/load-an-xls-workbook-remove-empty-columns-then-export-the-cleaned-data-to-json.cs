using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

class Program
{
    static void Main()
    {
        // Load the existing XLS workbook
        Workbook workbook = new Workbook("input.xls");

        // Access the first worksheet (adjust index if needed)
        Worksheet sheet = workbook.Worksheets[0];

        // Remove all columns that are completely blank
        sheet.Cells.DeleteBlankColumns();

        // Set up JSON export options
        JsonSaveOptions jsonOptions = new JsonSaveOptions
        {
            // Skip rows that contain no data
            SkipEmptyRows = true,
            // Do not include empty cells in the JSON output
            ExportEmptyCells = false,
            // Treat the first row as header names
            HasHeaderRow = true
        };

        // Export the cleaned workbook to a JSON file
        workbook.Save("output.json", jsonOptions);
    }
}