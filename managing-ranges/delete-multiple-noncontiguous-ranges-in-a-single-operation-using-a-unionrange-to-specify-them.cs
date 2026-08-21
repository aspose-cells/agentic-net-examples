// Title: C# – Delete Multiple Non‑Contiguous Ranges Using UnionRange in Aspose.Cells
// Description: Creates a workbook, populates sample data, defines a UnionRange (e.g., A2:B3,D5:E6), deletes each sub‑range while shifting cells up, and saves the result as DeletedNonContiguousRanges.xlsx.
// Keywords: Aspose.Cells delete non‑contiguous ranges | UnionRange C# example | delete multiple ranges shift up | .NET spreadsheet cell removal | Aspose.Cells bulk delete cells
// Common Searches: Aspose.Cells delete several non‑adjacent blocks | C# UnionRange delete cells and shift up | remove multiple ranges in one operation Aspose.Cells | how to delete non‑contiguous ranges .NET
// Developer Intent: Remove specified non‑contiguous cell blocks from a worksheet and shift the remaining cells upward in a single workflow.
// Use Cases: Strip out separate header/footer sections from a generated report. | Clean data by deleting scattered rows or columns that are not needed for export. | Implement a bulk delete feature for user‑selected cells across different sheet areas.
// AI Prompts: Write C# code with Aspose.Cells that deletes multiple non‑contiguous ranges defined by a UnionRange and shifts cells up. | Explain the UnionRange class in Aspose.Cells and how to iterate its ranges for deletion. | Show an alternative approach to delete non‑contiguous ranges without explicit loops.

using System;
using Aspose.Cells;

// Creates a workbook, populates sample data, defines a UnionRange (e.g., A2:B3,D5:E6), deletes each sub‑range while shifting cells up, and saves the result as DeletedNonContiguousRanges.xlsx.
class DeleteNonContiguousRanges
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Fill the worksheet with sample data
            for (int row = 0; row < 10; row++)
            {
                for (int col = 0; col < 5; col++)
                {
                    cells[row, col].PutValue($"R{row}C{col}");
                }
            }

            // Create a UnionRange that represents non‑contiguous areas (e.g., A2:B3 and D5:E6)
            UnionRange unionRange = workbook.Worksheets.CreateUnionRange("A2:B3,D5:E6", 0);

            // Delete each range in the union and shift cells up
            foreach (Aspose.Cells.Range range in unionRange.Ranges)
            {
                cells.DeleteRange(
                    range.FirstRow,
                    range.FirstColumn,
                    range.RowCount,
                    range.ColumnCount,
                    ShiftType.Up);
            }

            // Save the modified workbook
            workbook.Save("DeletedNonContiguousRanges.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
