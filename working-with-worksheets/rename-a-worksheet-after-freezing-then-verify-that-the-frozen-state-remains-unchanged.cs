// Title: Aspose.Cells for .NET – Rename a Worksheet After Freezing Panes and Confirm Freeze Settings
// Description: Shows how to apply Freeze Panes at C3 (3 rows × 3 columns) in a new workbook, retrieve the freeze parameters with GetFreezedPanes, rename the sheet, read the parameters again, and verify that the frozen rows and columns stay unchanged before saving the file.
// Keywords: Aspose.Cells | C# | freeze panes | worksheet rename | GetFreezedPanes | .NET | preserve freeze state | Excel workbook | sheet renaming | validation
// Common Searches: Aspose.Cells keep freeze panes after sheet rename | GetFreezedPanes example in C# | verify freeze pane persistence with Aspose.Cells | rename worksheet without losing frozen rows | C# code to check freeze state before and after rename
// Developer Intent: Rename an Excel worksheet while ensuring the existing freeze‑pane configuration remains intact and can be programmatically validated.
// Use Cases: Generate a report, freeze header rows, then rename the sheet to match a dynamic title without altering the view layout. | Run automated checks after batch‑renaming worksheets to confirm that frozen rows/columns are unchanged. | Create unit tests that assert freeze pane settings survive worksheet name changes in CI pipelines.
// AI Prompts: Write C# code using Aspose.Cells that freezes panes, renames the worksheet, and asserts identical GetFreezedPanes results before and after the rename. | Provide an xUnit test verifying that renaming a sheet does not modify its frozen rows or columns with Aspose.Cells for .NET. | Explain the internal storage of freeze pane information in Aspose.Cells and why a worksheet name change does not affect it.

using System;
using Aspose.Cells;

namespace WorksheetRenameFreezeDemo
{
    // Shows how to apply Freeze Panes at C3 (3 rows × 3 columns) in a new workbook, retrieve the freeze parameters with GetFreezedPanes, rename the sheet, read the parameters again, and verify that the frozen rows and columns stay unchanged before saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Freeze panes at cell C3 with 3 frozen rows and 3 frozen columns
            sheet.FreezePanes("C3", 3, 3);

            // Capture freeze pane information before renaming
            int rowBefore, colBefore, frozenRowsBefore, frozenColsBefore;
            bool isFrozenBefore = sheet.GetFreezedPanes(out rowBefore, out colBefore, out frozenRowsBefore, out frozenColsBefore);
            Console.WriteLine($"Before rename - Frozen: {isFrozenBefore}, Row: {rowBefore}, Column: {colBefore}, FrozenRows: {frozenRowsBefore}, FrozenCols: {frozenColsBefore}");

            // Rename the worksheet
            string newName = "RenamedSheet";
            sheet.Name = newName;

            // Capture freeze pane information after renaming
            int rowAfter, colAfter, frozenRowsAfter, frozenColsAfter;
            bool isFrozenAfter = sheet.GetFreezedPanes(out rowAfter, out colAfter, out frozenRowsAfter, out frozenColsAfter);
            Console.WriteLine($"After rename - Frozen: {isFrozenAfter}, Row: {rowAfter}, Column: {colAfter}, FrozenRows: {frozenRowsAfter}, FrozenCols: {frozenColsAfter}");

            // Verify that the freeze state has not changed
            bool freezeStateUnchanged = isFrozenBefore == isFrozenAfter &&
                                        rowBefore == rowAfter &&
                                        colBefore == colAfter &&
                                        frozenRowsBefore == frozenRowsAfter &&
                                        frozenColsBefore == frozenColsAfter;

            Console.WriteLine($"Freeze state unchanged after rename: {freezeStateUnchanged}");

            // Save the workbook to a file
            workbook.Save("WorksheetRenameFreezeDemo.xlsx");
        }
    }
}
