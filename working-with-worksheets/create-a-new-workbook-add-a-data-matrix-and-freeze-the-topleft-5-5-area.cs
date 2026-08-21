// Title: C# – Create a Workbook, Fill a 5×5 Matrix, and Freeze the Top‑Left Pane with Aspose.Cells
// Description: This example shows how to instantiate a new Aspose.Cells Workbook, populate cells A1:E5 with sample data, freeze the first five rows and five columns using the FreezePanes(5,5,5,5) method, and save the result as TopLeft5x5_Frozen.xlsx.
// Keywords: Aspose.Cells | C# | FreezePanes | freeze top left pane | 5x5 matrix | create workbook | populate cells | Excel automation
// Common Searches: Aspose.Cells freeze first 5 rows and columns C# | how to freeze top left area in Excel using Aspose.Cells | C# example to create 5x5 data table with Aspose.Cells | FreezePanes method parameters Aspose.Cells .NET
// Developer Intent: Generate a new workbook, insert a 5×5 data block, and lock that block in place with frozen panes.
// Use Cases: Static header rows and columns for large spreadsheets | Dashboard sheet with a fixed summary matrix | Printable report where the top‑left block stays visible on every page
// AI Prompts: Write a C# snippet that freezes N rows and M columns with Aspose.Cells, using variables for N and M. | Explain how the four parameters of FreezePanes map to zero‑based row/column indices. | Show how to apply bold font and background shading to the frozen 5×5 area before saving.

using System;
using Aspose.Cells;

namespace AsposeCellsFreezeTopLeft5x5
{
    // This example shows how to instantiate a new Aspose.Cells Workbook, populate cells A1:E5 with sample data, freeze the first five rows and five columns using the FreezePanes(5,5,5,5) method, and save the result as TopLeft5x5_Frozen.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate a 5 × 5 data matrix (rows 0‑4, columns 0‑4)
            for (int row = 0; row < 5; row++)
            {
                for (int col = 0; col < 5; col++)
                {
                    // Example data: "R{row}C{col}"
                    worksheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Freeze the top‑left 5 × 5 area.
            // Parameters: row index, column index, number of frozen rows, number of frozen columns.
            // Row/column indices are zero‑based, so use 5 to start the scrollable area after the frozen pane.
            worksheet.FreezePanes(5, 5, 5, 5); // freeze panes rule

            // Save the workbook (lifecycle rule)
            workbook.Save("TopLeft5x5_Frozen.xlsx");
        }
    }
}
