using System;
using Aspose.Cells;

class FreezePaneDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Fill the worksheet with sample data to make scrolling visible
        for (int row = 0; row < 30; row++)
        {
            for (int col = 0; col < 10; col++)
            {
                sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
            }
        }

        // Freeze the top 3 rows and the left 2 columns.
        // Parameters: row index, column index, number of frozen rows, number of frozen columns.
        // Row index 3 and column index 2 correspond to cell C4 (zero‑based).
        sheet.FreezePanes(3, 2, 3, 2);

        // Save the workbook to a file
        workbook.Save("FreezePaneDemo.xlsx");
    }
}