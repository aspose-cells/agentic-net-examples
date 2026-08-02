// Title: Delete Rows Above a Frozen Pane and Reset FreezePanes with Aspose.Cells for .NET (C#)
// Description: Loads an Excel workbook, reads the current freeze‑pane coordinates, removes every row that sits above the frozen row, and then re‑applies FreezePanes using the updated indices so the worksheet layout stays consistent before saving.
// Keywords: Aspose.Cells C# | .NET Excel freeze pane | DeleteRows above frozen row | Reset FreezePanes programmatically | GetFreezedPanes method | Excel worksheet row deletion | Adjust freeze pane after row removal | Automated Excel layout update
// Common Searches: Aspose.Cells delete rows above frozen pane C# | How to reset FreezePanes after deleting rows in .NET | GetFreezedPanes and FreezePanes usage example | Remove header rows while keeping freeze pane in Excel | Programmatic freeze pane adjustment with Aspose.Cells
// Developer Intent: Remove all rows that precede the current frozen row and re‑apply the freeze pane so it aligns with the new first visible row.
// Use Cases: Cleaning a report by discarding header rows that sit above a frozen pane while preserving the freeze position. | Dynamically trimming worksheets in generated Excel files and automatically updating freeze settings to keep the intended view. | Automating Excel data preparation where rows are removed and the freeze pane must be recalibrated for downstream users.
// AI Prompts: Generate C# code using Aspose.Cells that deletes every row above the frozen row and then calls FreezePanes with the corrected parameters. | Explain the interaction between GetFreezedPanes and FreezePanes when rows are removed from a worksheet in Aspose.Cells. | Provide a step‑by‑step tutorial for safely adjusting freeze pane settings after row deletions in an Aspose.Cells .NET project.

using System;
using Aspose.Cells;

namespace FreezePaneAdjustment
{
    // Loads an Excel workbook, reads the current freeze‑pane coordinates, removes every row that sits above the frozen row, and then re‑applies FreezePanes using the updated indices so the worksheet layout stays consistent before saving.
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Retrieve current freeze pane settings
            int frozenRow, frozenColumn, frozenRows, frozenColumns;
            bool hasFreeze = sheet.GetFreezedPanes(out frozenRow, out frozenColumn, out frozenRows, out frozenColumns);

            if (hasFreeze && frozenRow > 0)
            {
                // Delete all rows that are above the frozen row
                // DeleteRows(startRowIndex, totalRows)
                sheet.Cells.DeleteRows(0, frozenRow);

                // After deletion, the frozen row becomes the first visible row (index 0)
                // Re‑apply freeze panes with updated parameters
                // FreezePanes(row, column, freezedRows, freezedColumns)
                sheet.FreezePanes(0, frozenColumn, 0, frozenColumns);
            }

            // Save the modified workbook (replace with your desired output path)
            workbook.Save("output.xlsx");
        }
    }
}
