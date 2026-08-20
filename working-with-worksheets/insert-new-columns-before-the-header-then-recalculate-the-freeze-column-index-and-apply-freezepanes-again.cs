// Title: Insert Columns Before Header and Update Freeze Panes with Aspose.Cells for .NET
// Description: Load a workbook, insert one or more columns at the very start of the first worksheet, detect existing FreezePanes settings, adjust the frozen column index and count to account for the new columns, reapply FreezePanes, and save the updated file using Aspose.Cells in C#.
// Keywords: Aspose.Cells insert columns | C# FreezePanes update | adjust frozen column index | reapply FreezePanes after column insertion | .NET workbook manipulation | Excel freeze pane recalculation | InsertColumns before header
// Common Searches: Aspose.Cells insert columns before header keep freeze panes | C# update freeze pane column index after inserting columns | How to reapply FreezePanes in Aspose.Cells after adding columns | InsertColumns and preserve frozen rows in .NET | Adjust frozen columns when shifting worksheet data
// Developer Intent: Add leading columns to a worksheet, recalculate any frozen column indices, and reapply FreezePanes so the view remains consistent.
// Use Cases: Extend a report template with extra identifier columns while maintaining the original frozen header row and column. | Programmatically shift existing data right by inserting columns and ensure the freeze pane aligns with the new layout. | Pre‑process incoming workbooks that require column insertion before data import, automatically resetting freeze pane settings.
// AI Prompts: Write C# code using Aspose.Cells to insert N columns at the beginning of a worksheet and automatically update FreezePanes. | Explain how to recalculate freezeRow, freezeColumn, frozenRows, and frozenColumns after inserting columns in an Aspose.Cells workbook. | Provide a step‑by‑step guide for detecting existing freeze panes, adjusting indices, and reapplying FreezePanes when no freeze pane is set.

using System;
using Aspose.Cells;

namespace AsposeCellsInsertColumnsAndReapplyFreeze
{
    // Load a workbook, insert one or more columns at the very start of the first worksheet, detect existing FreezePanes settings, adjust the frozen column index and count to account for the new columns, reapply FreezePanes, and save the updated file using Aspose.Cells in C#.
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with actual path)
            Workbook workbook = new Workbook("Input.xlsx");

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Number of columns to insert before the header (e.g., 2 columns)
            int columnsToInsert = 2;

            // Insert columns at index 0 (before the first column / header)
            worksheet.Cells.InsertColumns(0, columnsToInsert);

            // Retrieve current freeze pane settings
            int freezeRow, freezeColumn, frozenRows, frozenColumns;
            bool hasFreeze = worksheet.GetFreezedPanes(out freezeRow, out freezeColumn, out frozenRows, out frozenColumns);

            if (hasFreeze)
            {
                // If the insertion point is before the frozen column, adjust the indices
                if (0 <= freezeColumn)
                {
                    freezeColumn += columnsToInsert;
                    frozenColumns += columnsToInsert;
                }

                // Reapply freeze panes with the updated indices
                worksheet.FreezePanes(freezeRow, freezeColumn, frozenRows, frozenColumns);
            }

            // Save the modified workbook (replace with desired output path)
            workbook.Save("Output.xlsx");
        }
    }
}
