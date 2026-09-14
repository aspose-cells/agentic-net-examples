// Title: How to adjust and reapply FreezePanes after inserting rows in an Aspose.Cells worksheet using C#
// AI Prompts: Write C# code that inserts rows into a worksheet, unfreezes the existing panes, recalculates the frozen row and count, and then calls FreezePanes with the new indices using Aspose.Cells. | Show how to retrieve the current frozen pane settings, modify them after a row insertion, and save the workbook with the corrected freeze‑pane layout.
// Common Searches: Aspose.Cells C# update freeze pane after inserting rows above frozen area | C# preserve frozen header rows when adding rows with Aspose.Cells | How to recalculate FreezePanes row index after row insertion in Aspose.Cells worksheet | Unfreeze and reapply FreezePanes with new row count using Aspose.Cells .NET | Maintain frozen rows after inserting rows in Excel file with Aspose.Cells C#
// Tags: Aspose.Cells FreezePanes after inserting rows | C# adjust frozen row index | unfreeze panes then reapply Aspose.Cells | preserve frozen header rows in worksheet | update frozen rows count programmatically

using System;
using Aspose.Cells;

namespace AsposeCellsFreezePaneUpdate
{
    // The example creates a workbook, freezes the first five rows and first column, inserts two rows above the frozen area, unfreezes the panes, updates the frozen row index and count to reflect the insertion, reapplies FreezePanes with the new values, and saves the file as UpdatedFreezePanes.xlsx.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Initial freeze: freeze first 5 rows and first column (row index 5, column index 0)
            worksheet.FreezePanes(5, 0, 5, 1);

            // Capture the current freeze pane settings
            worksheet.GetFreezedPanes(out int frozenRow, out int frozenColumn,
                                      out int frozenRowsCount, out int frozenColumnsCount);

            // Insert 2 rows at index 2 (before the original frozen row)
            int insertIndex = 2;
            int rowsToInsert = 2;
            worksheet.Cells.InsertRows(insertIndex, rowsToInsert);

            // Unfreeze the panes
            worksheet.UnFreezePanes();

            // Adjust the freeze pane indices to account for the inserted rows
            // If rows are inserted above the original frozen row, shift the row index and frozen rows count
            int newFrozenRow = frozenRow + rowsToInsert;
            int newFrozenRowsCount = frozenRowsCount + rowsToInsert;

            // Reapply freeze panes with the updated indices
            worksheet.FreezePanes(newFrozenRow, frozenColumn, newFrozenRowsCount, frozenColumnsCount);

            // Save the workbook
            workbook.Save("UpdatedFreezePanes.xlsx");
        }
    }
}
