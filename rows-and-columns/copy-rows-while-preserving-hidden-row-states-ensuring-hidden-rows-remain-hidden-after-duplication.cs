// Title: Copy rows while preserving hidden rows – Aspose.Cells for .NET
// Description: Demonstrates how to duplicate a range of rows in a workbook, hide the same rows in the copy, and retain all row settings by using Cells.CopyRows followed by Row.CopySettings. The resulting file keeps hidden rows hidden in both the original and the copied sections.
// Keywords: Aspose.Cells copy rows | preserve hidden rows | Row.CopySettings C# | Cells.CopyRows example | duplicate rows visibility | Aspose.Cells .NET hidden state
// Common Searches: Aspose.Cells copy rows keep hidden rows hidden | How to retain row visibility after copying in Aspose.Cells | CopyRows with hidden row flag in C# | Row.CopySettings after Cells.CopyRows
// Developer Intent: Duplicate a block of rows and maintain the hidden/visible state and other row attributes.
// Use Cases: Copy a table that includes collapsed grouping rows for a multi‑page report. | Create a template section that can be reused without losing hidden row configurations. | Generate repeated worksheet sections where hidden rows control conditional display.
// AI Prompts: Provide C# code that copies rows in Aspose.Cells and keeps hidden rows hidden using Row.CopySettings. | Explain why Row.CopySettings must be invoked after Cells.CopyRows to preserve visibility and formatting. | Show an example that copies rows with height, style, and hidden state intact in Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsRowCopyPreserveHidden
{
    // Demonstrates how to duplicate a range of rows in a workbook, hide the same rows in the copy, and retain all row settings by using Cells.CopyRows followed by Row.CopySettings. The resulting file keeps hidden rows hidden in both the original and the copied sections.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data in rows 0 to 4
            for (int r = 0; r < 5; r++)
            {
                cells[r, 0].PutValue($"Row {r + 1}");
            }

            // Hide rows 1 and 3 (zero‑based indexes)
            cells.HideRow(1); // Row 2
            cells.HideRow(3); // Row 4

            // Destination start row (copy rows 0‑4 to start at row 5)
            int sourceStart = 0;
            int destinationStart = 5;
            int rowCount = 5;

            // Copy rows data and formats
            cells.CopyRows(cells, sourceStart, destinationStart, rowCount);

            // Preserve hidden state (and other row settings) for each copied row
            for (int i = 0; i < rowCount; i++)
            {
                Row srcRow = sheet.Cells.Rows[sourceStart + i];
                Row destRow = sheet.Cells.Rows[destinationStart + i];
                // Copy all settings; checkStyle = false because source and destination are in the same workbook
                destRow.CopySettings(srcRow, false);
            }

            // Save the workbook (output will retain hidden rows in both original and copied sections)
            workbook.Save("RowsCopiedWithHiddenState.xlsx");
        }
    }
}
