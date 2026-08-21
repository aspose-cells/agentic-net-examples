// Title: Aspose.Cells C# – Insert Rows Above Header and Recalculate FreezePanes
// Description: Demonstrates how to add rows before a worksheet header, adjust the frozen row index, and reapply FreezePanes in a .NET workbook using Aspose.Cells.
// Keywords: Aspose.Cells insert rows C# | FreezePanes after row insertion | update frozen rows Aspose.Cells | reapply FreezePanes .NET | worksheet row insertion example
// Common Searches: insert rows above header Aspose.Cells C# | keep freeze panes after adding rows | recalculate frozen row index .NET | Aspose.Cells FreezePanes shift after insert | how to adjust freeze panes when inserting rows
// Developer Intent: Add rows before the header row and automatically reposition the frozen rows by updating FreezePanes settings.
// Use Cases: Add title or spacing rows at the top of a sheet while preserving a frozen header. | Programmatically modify worksheet layout without losing freeze‑pane visibility. | Perform multiple top‑row insertions and keep freeze pane alignment consistent.
// AI Prompts: Write C# code with Aspose.Cells that inserts N rows at the top of a worksheet and updates FreezePanes accordingly. | Explain the steps to retrieve current freeze pane parameters, insert rows, and reapply FreezePanes with the new indices. | Provide a concise tutorial for maintaining frozen header rows when inserting rows before them in a .NET workbook.

using System;
using Aspose.Cells;

namespace AsposeCellsFreezePaneExample
{
    // Demonstrates how to add rows before a worksheet header, adjust the frozen row index, and reapply FreezePanes in a .NET workbook using Aspose.Cells.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data: header at row 2 (index 1) and some data rows
            worksheet.Cells["A2"].PutValue("Header");
            worksheet.Cells["A3"].PutValue("Data 1");
            worksheet.Cells["A4"].PutValue("Data 2");
            worksheet.Cells["A5"].PutValue("Data 3");

            // Freeze the first two rows (including the header)
            // Parameters: row index, column index, number of frozen rows, number of frozen columns
            worksheet.FreezePanes(2, 0, 2, 0);

            // Capture current freeze pane settings
            int frozenRow, frozenColumn, frozenRows, frozenColumns;
            bool hasFreeze = worksheet.GetFreezedPanes(out frozenRow, out frozenColumn, out frozenRows, out frozenColumns);

            // Number of rows to insert before the header
            int rowsToInsert = 2;

            // Insert rows at the top of the sheet (row index 0)
            worksheet.Cells.InsertRows(0, rowsToInsert);

            // Reapply freeze panes with updated row index if the sheet was previously frozen
            if (hasFreeze)
            {
                // The frozen row index shifts down by the number of inserted rows
                int newFrozenRow = frozenRow + rowsToInsert;
                worksheet.FreezePanes(newFrozenRow, frozenColumn, frozenRows, frozenColumns);
            }

            // Save the workbook
            workbook.Save("FreezePaneAfterInsertRows.xlsx");
        }
    }
}
