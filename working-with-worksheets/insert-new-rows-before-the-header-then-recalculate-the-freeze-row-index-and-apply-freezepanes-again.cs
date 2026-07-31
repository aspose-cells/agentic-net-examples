// Title: C# – Insert Rows Above Header and Update FreezePanes with Aspose.Cells
// Description: Shows how to insert rows at the top of a worksheet, read the current FreezePanes parameters, adjust the frozen row index, and reapply FreezePanes so the header remains frozen using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | InsertRows | FreezePanes | adjust frozen row index | header row freeze | worksheet automation | Excel programmatic freeze pane | top rows insertion | Excel .NET API
// Common Searches: Aspose.Cells insert rows at top of sheet | keep freeze pane after inserting rows C# | update FreezePanes row index Aspose.Cells | reapply FreezePanes after adding rows | how to shift frozen rows in Aspose.Cells
// Developer Intent: Add new rows before the header row, recalculate the frozen row position, and reapply FreezePanes to preserve the original header freeze.
// Use Cases: Prepend blank rows for printing layout while maintaining a frozen header. | Insert summary or title rows at the beginning of a generated report without losing existing freeze settings. | Programmatically add data rows to an existing worksheet and automatically keep the header row frozen. | Adjust freeze panes after bulk row insertion in automated Excel export pipelines.
// AI Prompts: Generate C# code that inserts N rows at index 0 and updates FreezePanes so the original header stays frozen using Aspose.Cells. | Create a reusable method for Aspose.Cells that takes a row count, inserts rows at the top, recalculates frozen row index, and reapplies FreezePanes. | Explain step‑by‑step how to retrieve current FreezePanes parameters, modify them after inserting rows, and reapply them without affecting frozen columns.

using System;
using Aspose.Cells;

namespace AsposeCellsFreezePaneAdjustment
{
    // Shows how to insert rows at the top of a worksheet, read the current FreezePanes parameters, adjust the frozen row index, and reapply FreezePanes so the header remains frozen using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate sample data: header in row 0 and some data rows
            cells["A1"].PutValue("Header");
            cells["A2"].PutValue("Data 1");
            cells["A3"].PutValue("Data 2");
            cells["A4"].PutValue("Data 3");

            // Freeze the first row (header) initially
            // Freeze at row index 1 (second row) with 1 frozen row and 0 frozen columns
            worksheet.FreezePanes(1, 0, 1, 0);

            // Number of rows to insert before the header
            int rowsToInsert = 2;

            // Insert rows at the very top (index 0)
            cells.InsertRows(0, rowsToInsert);

            // Retrieve the original freeze pane information
            int frozenRow, frozenColumn, frozenRows, frozenColumns;
            bool hasFreeze = worksheet.GetFreezedPanes(out frozenRow, out frozenColumn, out frozenRows, out frozenColumns);

            if (hasFreeze)
            {
                // Adjust the freeze row index to account for the inserted rows
                int newFrozenRow = frozenRow + rowsToInsert;

                // Reapply freeze panes with the updated row index
                worksheet.FreezePanes(newFrozenRow, frozenColumn, frozenRows, frozenColumns);
            }

            // Save the workbook
            workbook.Save("Output.xlsx");
        }
    }
}
