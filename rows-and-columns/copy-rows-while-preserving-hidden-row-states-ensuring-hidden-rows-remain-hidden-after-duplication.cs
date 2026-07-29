// Title: Copy rows while preserving hidden rows in Aspose.Cells for .NET (C#)
// Description: Demonstrates how to duplicate a range of rows with Cells.CopyRows and then transfer row settings using Row.CopySettings so that hidden rows stay hidden, along with their height and style, in a C# workbook.
// Keywords: Aspose.Cells copy rows hidden | Row.CopySettings C# | preserve hidden rows Aspose | duplicate rows with formatting | copy rows retain visibility
// Common Searches: Aspose.Cells copy rows keep hidden rows hidden | C# copy rows preserve row height and hidden state | Row.CopySettings example Aspose.Cells | duplicate worksheet rows without losing hidden flags | how to copy rows with formatting in Aspose.Cells
// Developer Intent: Copy a block of rows to another location while ensuring that any hidden rows remain hidden and all row properties are retained.
// Use Cases: Clone a template section that contains hidden header rows for multiple report sheets. | Generate a printable summary that repeats a data block with hidden rows intact. | Batch‑process formulas by copying rows that include hidden rows without altering layout.
// AI Prompts: Show C# code to copy rows in Aspose.Cells and keep hidden rows hidden using Row.CopySettings. | Provide a method that duplicates rows from one index to another while preserving hidden state, row height, and style. | Explain why the checkStyle parameter should be false when copying rows within the same workbook.

using System;
using Aspose.Cells;

// Demonstrates how to duplicate a range of rows with Cells.CopyRows and then transfer row settings using Row.CopySettings so that hidden rows stay hidden, along with their height and style, in a C# workbook.
class CopyRowsPreserveHidden
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate sample data in rows 0‑4
        for (int i = 0; i < 5; i++)
        {
            cells[i, 0].PutValue($"Row {i + 1}");
        }

        // Hide specific rows (row indices are zero‑based)
        cells.HideRow(1); // hide row 2
        cells.HideRow(3); // hide row 4

        // Define the source rows to copy and the destination start row
        int sourceStartRow = 0;   // first row to copy
        int destinationStartRow = 5; // where the copy will be placed
        int rowsToCopy = 5;       // number of rows to copy

        // Copy rows data and formats
        cells.CopyRows(cells, sourceStartRow, destinationStartRow, rowsToCopy);

        // Preserve hidden state (and other row settings) by copying row settings
        for (int i = 0; i < rowsToCopy; i++)
        {
            Row sourceRow = cells.Rows[sourceStartRow + i];
            Row destinationRow = cells.Rows[destinationStartRow + i];
            // copy all settings; set checkStyle to false because rows belong to same workbook
            destinationRow.CopySettings(sourceRow, false);
        }

        // Save the workbook (lifecycle rule: use provided save method)
        workbook.Save("CopyRowsPreserveHidden.xlsx");
    }
}
