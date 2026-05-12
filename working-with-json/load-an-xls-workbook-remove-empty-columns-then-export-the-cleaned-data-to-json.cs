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
        Worksheet worksheet = workbook.Worksheets[0];

        // Remove all blank columns that contain no data
        worksheet.Cells.DeleteBlankColumns();

        // Configure JSON export options
        JsonSaveOptions jsonOptions = new JsonSaveOptions
        {
            // Skip rows that are completely empty
            SkipEmptyRows = true,
            // Do not include empty cells in the JSON output
            ExportEmptyCells = false,
            // Treat the first row as header (optional, based on source data)
            HasHeaderRow = true
        };

        // Export the cleaned workbook to a JSON file
        workbook.Save("output.json", jsonOptions);
    }
}