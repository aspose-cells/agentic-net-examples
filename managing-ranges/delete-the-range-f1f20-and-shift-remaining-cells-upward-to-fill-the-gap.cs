// Title: C# – Delete range F1:F20 in Aspose.Cells and shift cells upward
// Description: Shows how to delete the vertical range F1:F20 in a workbook with Aspose.Cells for .NET, shift the remaining cells up, and save the result.
// Keywords: Aspose.Cells DeleteRange | ShiftType.Up | C# delete cells Aspose | remove rows column F | Aspose.Cells .NET example | delete range and shift up | Excel cell deletion Aspose | bulk delete cells Aspose.Cells
// Common Searches: Aspose.Cells delete range F1:F20 | C# Aspose.Cells shift cells up after delete | How to use Cells.DeleteRange with ShiftType.Up | Remove rows 1-20 from column F Aspose.Cells | Aspose.Cells delete vertical range example
// Developer Intent: Delete cells F1 through F20 and move the cells below up to fill the empty space.
// Use Cases: Cleaning a worksheet by removing a block of data in column F while keeping the remaining rows in order. | Generating reports where specific rows in a column must be omitted and the rest of the column should collapse upward. | Automating data preparation that requires stripping header or placeholder rows from a column and preserving layout.
// AI Prompts: Provide C# code that deletes the range F1:F20 in an Aspose.Cells workbook and shifts remaining cells up. | Show an example of using Cells.DeleteRange with ShiftType.Up to remove a vertical range in Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Shows how to delete the vertical range F1:F20 in a workbook with Aspose.Cells for .NET, shift the remaining cells up, and save the result.
class DeleteRangeExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate sample data (optional, just to illustrate the effect)
        for (int i = 0; i < 30; i++)
        {
            cells[i, 0].PutValue($"A{i + 1}");   // Column A
            cells[i, 5].PutValue($"F{i + 1}");   // Column F
        }

        // Delete the range F1:F20 (rows 0‑19, column index 5) and shift remaining cells upward
        cells.DeleteRange(0, 5, 19, 5, ShiftType.Up);

        // Save the modified workbook
        workbook.Save("DeletedRange.xlsx");
    }
}
