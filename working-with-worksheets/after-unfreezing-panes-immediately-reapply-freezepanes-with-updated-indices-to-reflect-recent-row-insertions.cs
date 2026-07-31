// Title: Update FreezePanes After Row Insertion in Aspose.Cells for .NET (C#)
// Description: Demonstrates how to keep frozen rows and columns accurate after inserting rows in an Aspose.Cells worksheet. The example freezes the top rows, records the original pane settings, inserts rows above the frozen area, unfreezes the panes, recalculates the freeze row index, reapplies FreezePanes with the new index, and saves the workbook.
// Keywords: Aspose.Cells FreezePanes C# | adjust frozen rows after insert Aspose.Cells | UnFreezePanes then FreezePanes .NET | recalculate frozen pane indices | programmatic pane update Aspose.Cells | row insertion impact on frozen panes
// Common Searches: How to keep frozen header rows after inserting rows in Aspose.Cells | Aspose.Cells update FreezePanes row index after InsertRows | Unfreeze and reapply FreezePanes in C# | Adjust frozen pane when adding rows before freeze area | Recalculate FreezePanes coordinates Aspose.Cells .NET
// Developer Intent: Refresh the frozen pane position to match newly inserted rows by unfreezing and re‑applying FreezePanes with corrected indices.
// Use Cases: Preserve header rows frozen while programmatically adding rows above them. | Maintain column freeze when bulk inserting rows in generated reports. | Synchronize pane layout after dynamic data insertion in dashboards. | Automate pane adjustments in Excel export utilities.
// AI Prompts: Generate C# code that inserts rows into an Aspose.Cells worksheet and automatically updates the FreezePanes row index. | Create a helper method for Aspose.Cells that takes original freeze settings and insertion details and returns updated FreezePanes parameters. | Explain the step‑by‑step process of using UnFreezePanes and FreezePanes to keep pane layout consistent after row insertions in Aspose.Cells. | Provide a reusable snippet that recalculates frozen pane indices when rows are inserted before the frozen area.

using System;
using Aspose.Cells;

namespace AsposeCellsPaneUpdateDemo
{
    // Demonstrates how to keep frozen rows and columns accurate after inserting rows in an Aspose.Cells worksheet. The example freezes the top rows, records the original pane settings, inserts rows above the frozen area, unfreezes the panes, recalculates the freeze row index, reapplies FreezePanes with the new index, and saves the workbook.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Initial freeze: freeze at row index 4 (5th row) and column index 1 (2nd column)
            // Freeze 4 rows and 1 column (the frozenRows/Columns must not exceed the indices)
            int initialRow = 4;          // zero‑based row index where the split occurs
            int initialColumn = 1;       // zero‑based column index where the split occurs
            int frozenRows = 4;          // number of rows frozen in the top pane
            int frozenColumns = 1;       // number of columns frozen in the left pane
            worksheet.FreezePanes(initialRow, initialColumn, frozenRows, frozenColumns);

            // Store the original freeze information for later recalculation
            int originalRow, originalColumn, originalFrozenRows, originalFrozenColumns;
            worksheet.GetFreezedPanes(out originalRow, out originalColumn, out originalFrozenRows, out originalFrozenColumns);

            // Insert 2 rows at index 2 (i.e., before the original frozen row)
            int insertRowIndex = 2;
            int rowsToInsert = 2;
            worksheet.Cells.InsertRows(insertRowIndex, rowsToInsert);

            // Unfreeze the panes before applying the updated freeze
            worksheet.UnFreezePanes();

            // Recalculate the new freeze row index.
            // If the insertion point is at or above the original freeze row, shift the freeze row down.
            int newRow = originalRow;
            if (insertRowIndex <= originalRow)
            {
                newRow += rowsToInsert;
            }

            // Reapply FreezePanes with the updated row index while keeping other parameters unchanged
            worksheet.FreezePanes(newRow, originalColumn, originalFrozenRows, originalFrozenColumns);

            // Save the workbook
            workbook.Save("PaneUpdateAfterRowInsert.xlsx");
        }
    }
}
