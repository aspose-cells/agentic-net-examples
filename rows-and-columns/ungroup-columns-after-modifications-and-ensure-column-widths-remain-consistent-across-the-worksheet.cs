// Title: Aspose.Cells C# – Ungroup Columns While Keeping Original Widths
// Description: Shows how to record column widths, group a range, ungroup it, and then restore the saved widths so the worksheet layout remains unchanged. The example uses GroupColumns, UngroupColumns and re‑applies the stored widths before saving the workbook.
// Keywords: Aspose.Cells | C# | UngroupColumns | GroupColumns | column width preservation | restore column width | Excel worksheet layout | column grouping | Aspose.Cells .NET | maintain column dimensions
// Common Searches: how to keep column widths after ungrouping in Aspose.Cells | Aspose.Cells .NET restore column width after grouping | C# ungroup columns preserve layout Aspose.Cells | capture and reapply column widths Excel using Aspose | group then ungroup columns without changing size Aspose.Cells
// Developer Intent: Reapply previously captured column widths after ungrouping a column range in an Excel worksheet with Aspose.Cells.
// Use Cases: Save the visual layout of a sheet before performing structural changes such as grouping. | Undo column grouping while ensuring the original column dimensions are retained. | Automate batch processing of workbooks where column widths must stay consistent after grouping operations.
// AI Prompts: Write C# code using Aspose.Cells that saves column widths, groups columns 0‑2, ungroups them, and restores the saved widths before saving the file. | Create a reusable Aspose.Cells method that accepts a column range, records its widths, performs grouping or ungrouping, and then re‑applies the original widths. | Generate an example that maintains column width consistency across an entire worksheet after multiple group/ungroup actions with Aspose.Cells in .NET.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsColumnUngroupDemo
{
    // Shows how to record column widths, group a range, ungroup it, and then restore the saved widths so the worksheet layout remains unchanged. The example uses GroupColumns, UngroupColumns and re‑applies the stored widths before saving the workbook.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook(); // lifecycle: create
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // ------------------------------------------------------------
            // Step 1: Capture current column widths (to keep them consistent later)
            // ------------------------------------------------------------
            // We'll store widths for the first 10 columns as an example.
            // Adjust the range as needed for your worksheet.
            Dictionary<int, double> originalWidths = new Dictionary<int, double>();
            for (int col = 0; col < 10; col++)
            {
                // Access the Column object; it will be instantiated if not existing.
                Column column = cells.Columns[col];
                originalWidths[col] = column.Width; // default width if never set
            }

            // ------------------------------------------------------------
            // Step 2: Perform some modifications that involve grouping columns
            // ------------------------------------------------------------
            // Example: Group columns 0 to 2
            cells.GroupColumns(0, 2); // using the provided GroupColumns rule

            // (Additional modifications can be placed here)

            // ------------------------------------------------------------
            // Step 3: Ungroup the previously grouped columns
            // ------------------------------------------------------------
            cells.UngroupColumns(0, 2); // using the provided UngroupColumns rule

            // ------------------------------------------------------------
            // Step 4: Reapply the original column widths to maintain consistency
            // ------------------------------------------------------------
            foreach (var kvp in originalWidths)
            {
                int colIndex = kvp.Key;
                double width = kvp.Value;
                cells.Columns[colIndex].Width = width; // free‑form code (no specific rule needed)
            }

            // ------------------------------------------------------------
            // Step 5: Save the workbook
            // ------------------------------------------------------------
            workbook.Save("output.xlsx"); // lifecycle: save
        }
    }
}
