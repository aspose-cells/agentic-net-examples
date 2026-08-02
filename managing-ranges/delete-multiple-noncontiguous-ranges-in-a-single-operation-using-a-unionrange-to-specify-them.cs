// Title: C# – Delete Multiple Non‑Contiguous Ranges Using UnionRange in Aspose.Cells
// Description: Creates a workbook, fills a 10×10 grid, defines a UnionRange covering A2:B3 and D5:E6, iterates each sub‑range, deletes it with Cells.DeleteRange (ShiftType.Up), and saves the result as an XLSX file.
// Keywords: Aspose.Cells C# | UnionRange delete | non‑contiguous range removal | Cells.DeleteRange | ShiftType.Up | Excel range deletion | .NET Excel library | multiple range delete | Aspose.Cells example GitHub | XLSX file manipulation
// Common Searches: Aspose.Cells delete non adjacent ranges C# | How to use UnionRange to remove several areas in a worksheet | Delete multiple cell blocks with ShiftType.Up in Aspose.Cells | C# code for deleting non‑contiguous ranges in Excel | Aspose.Cells UnionRange example GitHub
// Developer Intent: Remove specified non‑contiguous cell blocks from a worksheet in a single operation and shift the remaining cells upward.
// Use Cases: Clear header/footer sections after data processing without disturbing the main table. | Strip out obsolete data blocks before importing a new dataset into the same sheet. | Delete separate merged‑cell areas in a generated report while keeping surrounding content intact.
// AI Prompts: Generate a C# method that accepts a UnionRange address string and deletes all its areas using ShiftType.Left. | Explain best practices for exception handling when deleting multiple ranges with UnionRange in Aspose.Cells. | Show an alternative technique to delete non‑contiguous ranges without looping over each Range object.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsUnionRangeDeleteDemo
{
    // Creates a workbook, fills a 10×10 grid, defines a UnionRange covering A2:B3 and D5:E6, iterates each sub‑range, deletes it with Cells.DeleteRange (ShiftType.Up), and saves the result as an XLSX file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Populate sample data to visualize the ranges before deletion
                for (int row = 0; row < 10; row++)
                {
                    for (int col = 0; col < 10; col++)
                    {
                        cells[row, col].PutValue($"R{row}C{col}");
                    }
                }

                // Define a union range consisting of two non‑contiguous areas:
                //   A2:B3  (rows 1‑2, columns 0‑1)
                //   D5:E6  (rows 4‑5, columns 3‑4)
                UnionRange unionRange = workbook.Worksheets.CreateUnionRange("A2:B3,D5:E6", 0);

                // Delete each area of the union range.
                // Use Aspose.Cells.Range to avoid conflict with System.Range
                foreach (Aspose.Cells.Range r in unionRange.Ranges)
                {
                    int startRow = r.FirstRow;
                    int startColumn = r.FirstColumn;
                    int totalRows = r.RowCount;
                    int totalColumns = r.ColumnCount;

                    // Delete the range and shift cells up
                    cells.DeleteRange(startRow, startColumn, totalRows, totalColumns, ShiftType.Up);
                }

                // Save the workbook
                workbook.Save("UnionRangeDeleteDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
