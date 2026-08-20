// Title: Freeze the top row in worksheets with more than 100 rows using Aspose.Cells for .NET
// Description: C# example that creates or loads a workbook, populates each sheet with data, checks ws.Cells.Rows.Count, and calls ws.FreezePanes(1,0,1,0) to lock the first row only when the sheet exceeds 100 rows, then saves the file.
// Keywords: Aspose.Cells | C# | .NET | FreezePanes | conditional freeze panes | freeze top row | row count | Excel workbook | large worksheet handling | programmatic Excel formatting
// Common Searches: Aspose.Cells freeze first row if rows > 100 | C# conditional FreezePanes based on row count | How to lock header row in large Excel sheets using Aspose.Cells | Check worksheet row count and apply FreezePanes .NET | Freeze panes for worksheets with many rows Aspose
// Developer Intent: Automatically apply a freeze pane to the first row of any worksheet that contains more than one hundred rows.
// Use Cases: Generating Excel reports where header rows stay visible on large data sets. | Processing multi‑sheet workbooks and applying freeze panes only to sheets that exceed a size threshold. | Improving user navigation in exported Excel files by locking the top row for sheets with extensive rows.
// AI Prompts: Create a reusable method that iterates through all worksheets in an Aspose.Cells workbook and freezes the first row when the sheet has over 100 rows. | Explain each parameter of ws.FreezePanes and show how to extend the logic to also freeze the first column when the row count exceeds 200. | Provide sample code that saves the workbook after applying conditional freeze panes and includes proper exception handling.

using System;
using Aspose.Cells;

namespace FreezeRowsExample
{
    // C# example that creates or loads a workbook, populates each sheet with data, checks ws.Cells.Rows.Count, and calls ws.FreezePanes(1,0,1,0) to lock the first row only when the sheet exceeds 100 rows, then saves the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Example: add sample data to demonstrate the logic
            foreach (Worksheet ws in workbook.Worksheets)
            {
                // Populate each worksheet with 150 rows of dummy data
                for (int i = 0; i < 150; i++)
                {
                    ws.Cells[i, 0].PutValue($"Row {i + 1}");
                }
            }

            // Iterate through all worksheets
            foreach (Worksheet ws in workbook.Worksheets)
            {
                // Determine the total number of rows in the worksheet
                // RowCollection.Count gives the total rows (including empty ones)
                int totalRows = ws.Cells.Rows.Count;

                // Apply freeze panes if the worksheet has more than 100 rows
                if (totalRows > 100)
                {
                    // Freeze the top row (row index 1 means the second row, so the first row stays visible)
                    // Parameters: row, column, freezedRows, freezedColumns
                    ws.FreezePanes(1, 0, 1, 0);
                }
            }

            // Save the workbook
            workbook.Save("FrozenRowsWorkbook.xlsx");
        }
    }
}
