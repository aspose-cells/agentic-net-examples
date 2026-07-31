// Title: C# Aspose.Cells: Move a Range While Preserving Frozen Panes
// Description: Demonstrates how to capture FreezePanes settings, relocate a cell range with Range.MoveTo, and re‑apply the original frozen rows/columns so the view stays unchanged, then saves the workbook.
// Keywords: Aspose.Cells C# | .NET Excel automation | move range without affecting freeze panes | Range.MoveTo example | preserve frozen rows and columns | reapply FreezePanes programmatically | Excel freeze pane handling | GitHub Aspose.Cells sample | worksheet range relocation
// Common Searches: Aspose.Cells move range keep frozen panes | C# preserve FreezePanes after moving cells | how to reapply FreezePanes in Aspose.Cells | Range.MoveTo does not shift frozen rows | example code moving range without changing freeze settings
// Developer Intent: The developer needs to shift a specific cell block to a new location while ensuring any existing frozen rows or columns remain exactly where they were.
// Use Cases: Reposition a header block in a report without losing the frozen top rows that aid navigation. | Relocate a data table while keeping the first columns frozen for constant identifier visibility. | Adjust a template section programmatically in a shared workbook without altering the user's pane‑freeze layout.
// AI Prompts: Write C# code using Aspose.Cells to move a range and automatically restore the original FreezePanes configuration. | Explain step‑by‑step how to capture FreezePanes parameters, move cells with Range.MoveTo, and reapply the freeze settings. | Provide a concise guide for preserving frozen rows/columns after relocating data in an Aspose.Cells workbook.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsExamples
{
    // Demonstrates how to capture FreezePanes settings, relocate a cell range with Range.MoveTo, and re‑apply the original frozen rows/columns so the view stays unchanged, then saves the workbook.
    public class MoveRangeWithoutAffectingFrozenPanes
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate some sample data in the source range (A1:B2)
                cells["A1"].PutValue("A1");
                cells["B1"].PutValue("B1");
                cells["A2"].PutValue("A2");
                cells["B2"].PutValue("B2");

                // Freeze panes at cell C3 (row index 2, column index 2) with 2 frozen rows and 2 frozen columns
                sheet.FreezePanes(2, 2, 2, 2);

                // Capture current frozen pane settings
                int frozenRow, frozenColumn, frozenRows, frozenColumns;
                bool hasFreeze = sheet.GetFreezedPanes(out frozenRow, out frozenColumn, out frozenRows, out frozenColumns);

                // Define the range to move (A1:B2)
                AsposeRange rangeToMove = cells.CreateRange("A1", "B2");

                // Destination start cell (move to D4 => row index 3, column index 3)
                int destRow = 3;   // zero‑based index for row 4
                int destColumn = 3; // zero‑based index for column D

                // Move the range using the Range.MoveTo method
                rangeToMove.MoveTo(destRow, destColumn);

                // Re‑apply frozen panes if they existed before the move
                if (hasFreeze)
                {
                    // FreezePanes(row, column, freezedRows, freezedColumns)
                    sheet.FreezePanes(frozenRow, frozenColumn, frozenRows, frozenColumns);
                }

                // Save the workbook
                string outputPath = "MoveRange_NoFreezeShift.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            MoveRangeWithoutAffectingFrozenPanes.Run();
        }
    }
}
