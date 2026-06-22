using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

class Program
{
    static void Main()
    {
        // Load the workbook that contains the data.
        // Replace "input.xlsx" with the actual path to your Excel file.
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet worksheet = workbook.Worksheets[0];

        // Determine the last used row (0‑based index).
        int lastRow = worksheet.Cells.MaxDataRow; // includes header row

        // Set the auto‑filter range.
        // This example assumes the data starts at A1 and spans all columns up to the last used column.
        // Adjust the range as needed for your specific worksheet layout.
        worksheet.AutoFilter.SetRange(0, 0, lastRow);

        // Apply a filter on the "Status" column (assumed to be column B, index 1)
        // Keep only rows where the status equals "Active". Rows with "Inactive" will be hidden.
        worksheet.AutoFilter.Filter(1, "Active");
        worksheet.AutoFilter.Refresh();

        // Configure JSON export options.
        JsonSaveOptions jsonOptions = new JsonSaveOptions
        {
            // Skip rows that become empty after filtering.
            SkipEmptyRows = true,
            // Do not export empty cells as null.
            ExportEmptyCells = false
        };

        // Export the filtered workbook to a JSON file.
        // The hidden rows (those with status "Inactive") will not appear in the output.
        string jsonPath = "output.json";
        workbook.Save(jsonPath, jsonOptions);

        Console.WriteLine($"Filtered JSON has been saved to: {jsonPath}");
    }
}