// Title: C# – Copy rows 5‑10 to rows 20‑25 while preserving formatting with Aspose.Cells
// Description: Shows how to duplicate rows 5‑10 into rows 20‑25 of the same worksheet, keeping values, styles, merged cells and conditional formatting. The sample creates a workbook, optionally fills the source rows, uses Cells.CopyRows(source, 4, 19, 6), and saves the file as RowCopyResult.xlsx.
// Keywords: Aspose.Cells C# copy rows | Cells.CopyRows example | preserve formatting Aspose.Cells | copy multiple rows .NET | duplicate worksheet rows
// Common Searches: Aspose.Cells copy rows with formatting | C# copy rows 5 to 10 to 20 to 25 | How to use Cells.CopyRows in .NET | Preserve merged cells when copying rows Aspose | Copy rows between non‑adjacent ranges Aspose.Cells
// Developer Intent: Copy a block of rows to another location in the same worksheet without losing any formatting or formulas.
// Use Cases: Clone a template section to a new area for report generation while retaining styles. | Create a backup of specific rows before applying calculations or transformations. | Insert a pre‑filled data block at a precise position, preserving merged cells and conditional formatting.
// AI Prompts: Generate C# code using Aspose.Cells to copy rows 5‑10 to rows 20‑25 and keep all cell styles and formulas. | Explain each parameter of Cells.CopyRows and how formatting is handled during the copy operation. | Provide an Aspose.Cells example that copies rows containing merged cells and conditional formatting in .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsRowCopyExample
{
    // Shows how to duplicate rows 5‑10 into rows 20‑25 of the same worksheet, keeping values, styles, merged cells and conditional formatting. The sample creates a workbook, optionally fills the source rows, uses Cells.CopyRows(source, 4, 19, 6), and saves the file as RowCopyResult.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Example: fill source rows 5-10 with sample data (optional)
            for (int row = 4; row <= 9; row++)          // rows 5 to 10 (zero‑based index 4‑9)
            {
                for (int col = 0; col < 5; col++)
                {
                    cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Copy rows 5‑10 (indices 4‑9) to rows 20‑25 (indices 19‑24)
            // rowNumber = number of rows to copy = 6
            cells.CopyRows(cells, 4, 19, 6);

            // Save the workbook
            workbook.Save("RowCopyResult.xlsx");
        }
    }
}
