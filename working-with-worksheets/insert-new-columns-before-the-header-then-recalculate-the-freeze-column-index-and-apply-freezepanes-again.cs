// Title: Insert Columns at the Start and Re‑apply Freeze Panes in Aspose.Cells for .NET
// Description: Loads an existing workbook, inserts one or more columns before the header row, reads any frozen pane configuration, shifts the frozen column index to match the new layout, reapplies FreezePanes, and saves the updated file.
// Keywords: Aspose.Cells C# insert columns | freeze panes after column insertion | Worksheet.FreezePanes C# | GetFreezedPanes Aspose.Cells | Excel automation .NET | US developers Aspose.Cells | UK Excel library examples | GitHub Aspose.Cells sample | coding‑assistant Excel column insert
// Common Searches: how to add columns before header and keep freeze panes in Aspose.Cells | adjust frozen column index after inserting columns C# | Aspose.Cells recalculate FreezePanes when columns are added | C# code to preserve frozen rows and columns after InsertColumns | example of GetFreezedPanes and FreezePanes usage
// Developer Intent: Add new columns at the beginning of a worksheet and automatically update the frozen pane settings to reflect the column shift.
// Use Cases: Expand a template with placeholder columns while retaining the frozen header for easy navigation. | Programmatically insert reporting columns into an existing Excel file without losing the user’s frozen view. | Batch‑process workbooks to add extra data fields and ensure the original freeze pane layout remains intact.
// AI Prompts: Generate C# code that inserts N columns at index 0 in an Aspose.Cells worksheet and updates the frozen pane column index accordingly. | Explain step‑by‑step how GetFreezedPanes and FreezePanes work together to preserve frozen rows and columns after structural changes. | Create a reusable method in Aspose.Cells that inserts columns and automatically recalculates all freeze pane parameters.

using System;
using Aspose.Cells;

namespace AsposeCellsInsertColumnsAndRefreeze
{
    // Loads an existing workbook, inserts one or more columns before the header row, reads any frozen pane configuration, shifts the frozen column index to match the new layout, reapplies FreezePanes, and saves the updated file.
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing workbook (replace with your file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Number of columns to insert before the header (e.g., 2 columns)
            int insertIndex = 0;          // Insert at the very beginning (before column A)
            int insertCount = 2;          // Number of columns to insert

            // Insert the columns
            worksheet.Cells.InsertColumns(insertIndex, insertCount);

            // Retrieve current freeze pane settings
            bool hasFreeze = worksheet.GetFreezedPanes(out int freezeRow, out int freezeColumn,
                                                      out int frozenRows, out int frozenColumns);

            if (hasFreeze)
            {
                // Adjust the freeze column index to account for the newly inserted columns
                // If the freeze column is at or after the insertion point, shift it right
                int newFreezeColumn = freezeColumn >= insertIndex ? freezeColumn + insertCount : freezeColumn;

                // Reapply freeze panes with the updated column index
                worksheet.FreezePanes(freezeRow, newFreezeColumn, frozenRows, frozenColumns);
            }

            // Save the modified workbook
            workbook.Save("output.xlsx");
        }
    }
}
