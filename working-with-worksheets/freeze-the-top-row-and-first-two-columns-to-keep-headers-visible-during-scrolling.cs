using System;
using Aspose.Cells;

class FreezeTopRowAndFirstTwoColumns
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Populate some sample data (optional, just to visualize the freeze)
        for (int row = 0; row < 20; row++)
        {
            for (int col = 0; col < 5; col++)
            {
                sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
            }
        }

        // Freeze the top row (row index 1) and the first two columns (column index 2)
        // Parameters: row index, column index, number of frozen rows, number of frozen columns
        sheet.FreezePanes(1, 2, 1, 2);

        // Save the workbook to a file
        workbook.Save("FreezeTopRowFirstTwoColumns.xlsx");
    }
}