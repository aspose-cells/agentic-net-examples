using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the existing Excel workbook
        Workbook workbook = new Workbook("input.xlsx");   // replace with your source file path

        // Access the first worksheet (or iterate through worksheets as needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Remove all columns that do not contain any data
        worksheet.Cells.DeleteBlankColumns();

        // Set up JSON export options
        JsonSaveOptions jsonOptions = new JsonSaveOptions
        {
            // Skip rows that are completely empty after column cleanup
            SkipEmptyRows = true,
            // Do not include empty cells in the JSON output (they will be omitted)
            ExportEmptyCells = false,
            // Treat the first row as header if your data has headers
            HasHeaderRow = true
        };

        // Export the cleaned workbook to a JSON file
        workbook.Save("cleaned_output.json", jsonOptions);
    }
}