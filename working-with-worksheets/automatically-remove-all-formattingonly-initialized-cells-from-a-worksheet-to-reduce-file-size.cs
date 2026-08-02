// Title: Remove Formatting‑Only Cells and Unused Styles with Aspose.Cells (C#) to Shrink Excel Files
// Description: This example shows how to load an Excel workbook, scan the used range of the first worksheet, clear formatting from cells that contain no data, purge orphaned styles with Workbook.RemoveUnusedStyles, and save a smaller file. Ideal for .NET developers who need to optimize workbook size automatically.
// Keywords: Aspose.Cells clear formatting C# | remove unused styles Aspose.Cells | optimize Excel file size .NET | clear formats from empty cells | Workbook.RemoveUnusedStyles example | C# Excel performance tuning | reduce workbook size US developers | European .NET Excel optimization
// Common Searches: How to delete formatting from empty cells using Aspose.Cells for .NET | Remove unused styles after clearing formats in an Excel workbook | C# code to reduce Excel file size by clearing cell formats | Aspose.Cells RemoveUnusedStyles usage guide | Automate Excel cleanup with Aspose.Cells
// Developer Intent: Automatically strip formatting from empty cells and eliminate orphaned styles to produce a leaner Excel workbook.
// Use Cases: Clean up auto‑generated reports before archiving to save storage. | Prepare workbooks for high‑throughput server processing by removing unnecessary styles. | Add an optimization step to a batch Excel conversion pipeline for smaller output files.
// AI Prompts: Generate C# code using Aspose.Cells that clears formatting only on empty cells and then calls RemoveUnusedStyles. | Suggest a more performant approach to purge empty‑cell formatting across an entire worksheet with Aspose.Cells. | Explain the purpose of Workbook.RemoveUnusedStyles and why it should be executed after clearing cell formats.

using System;
using Aspose.Cells;

namespace AsposeCellsFormattingCleanup
{
    // This example shows how to load an Excel workbook, scan the used range of the first worksheet, clear formatting from cells that contain no data, purge orphaned styles with Workbook.RemoveUnusedStyles, and save a smaller file. Ideal for .NET developers who need to optimize workbook size automatically.
    class Program
    {
        static void Main()
        {
            // Load the workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Work with the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Determine the used range of the worksheet
            int maxRow = cells.MaxDataRow;
            int maxCol = cells.MaxDataColumn;

            // Iterate through each cell in the used range
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    // Retrieve the cell
                    Cell cell = cells[row, col];

                    // If the cell has no value (i.e., it is empty) but may have formatting,
                    // clear its formatting to avoid keeping unused styles.
                    // Checking for null or empty string covers typical empty cases.
                    if (cell.Value == null || string.IsNullOrEmpty(cell.StringValue))
                    {
                        // Clear formatting for this single cell
                        cells.ClearFormats(row, col, row, col);
                    }
                }
            }

            // After clearing formats from empty cells, remove any styles that are no longer used.
            workbook.RemoveUnusedStyles();

            // Save the optimized workbook (replace with desired output path)
            workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
    }
}
