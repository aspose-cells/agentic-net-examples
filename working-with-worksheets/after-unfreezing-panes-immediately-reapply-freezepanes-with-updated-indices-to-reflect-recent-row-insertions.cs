// Title: C# – Update Freeze Panes After Inserting Rows with Aspose.Cells (Unfreeze & Reapply)
// Description: Demonstrates how to freeze the top‑left pane, insert rows above the frozen area, unfreeze, recalculate the new freeze row index, and reapply FreezePanes using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# FreezePanes | update frozen rows after insert | unfreeze and refreeze worksheet | GetFreezedPanes example | InsertRows adjust freeze pane | .NET spreadsheet freeze pane | Aspose.Cells row insertion | Excel freeze pane programmatically
// Common Searches: Aspose.Cells refresh freeze pane after inserting rows | C# unfreeze and reapply FreezePanes | How to shift frozen rows in Aspose.Cells | GetFreezedPanes and FreezePanes usage | Adjust frozen pane indices after row insert
// Developer Intent: Reposition frozen rows/columns to reflect rows inserted above the original freeze point.
// Use Cases: Maintain header rows frozen when new rows are added at the top of a report. | Automatically correct frozen column positions after inserting columns before the frozen area. | Validate and fix freeze settings after bulk row‑insertion operations in generated workbooks.
// AI Prompts: Write C# code with Aspose.Cells that inserts rows and automatically updates the frozen pane coordinates. | Explain the interaction between GetFreezedPanes and FreezePanes for preserving frozen headers after row insertion. | Create a reusable method that accepts an insert index and count, updates freeze row/column values, and reapplies FreezePanes.

using System;
using Aspose.Cells;

namespace AsposeCellsFreezePaneUpdate
{
    // Demonstrates how to freeze the top‑left pane, insert rows above the frozen area, unfreeze, recalculate the new freeze row index, and reapply FreezePanes using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Initial freeze: freeze at row index 5, column index 3 (zero‑based)
            // This also freezes 5 rows and 3 columns in the top‑left pane
            int initialRow = 5;
            int initialColumn = 3;
            worksheet.FreezePanes(initialRow, initialColumn, initialRow, initialColumn);

            // Store the original freeze information
            int frozenRow, frozenColumn, frozenRows, frozenColumns;
            worksheet.GetFreezedPanes(out frozenRow, out frozenColumn, out frozenRows, out frozenColumns);

            // Insert 2 rows at index 2 (before the original frozen row)
            int insertIndex = 2;
            int rowsToInsert = 2;
            worksheet.Cells.InsertRows(insertIndex, rowsToInsert);

            // Unfreeze the panes
            worksheet.UnFreezePanes();

            // Calculate the new freeze row index.
            // If rows were inserted above the original frozen row, shift it down.
            int updatedRow = frozenRow;
            if (insertIndex <= frozenRow)
            {
                updatedRow += rowsToInsert;
            }

            // Reapply freeze with the updated row index while keeping the same column and counts
            worksheet.FreezePanes(updatedRow, frozenColumn, updatedRow, frozenColumn);

            // Optional: verify the new freeze settings
            worksheet.GetFreezedPanes(out frozenRow, out frozenColumn, out frozenRows, out frozenColumns);
            Console.WriteLine($"New freeze position - Row: {frozenRow}, Column: {frozenColumn}");
            Console.WriteLine($"Frozen rows: {frozenRows}, Frozen columns: {frozenColumns}");

            // Save the workbook
            workbook.Save("FreezePaneAfterInsert.xlsx");
        }
    }
}
