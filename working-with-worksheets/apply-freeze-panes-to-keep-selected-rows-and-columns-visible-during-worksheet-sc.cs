using System;
using Aspose.Cells;

class FreezePanesExample
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample data to make scrolling noticeable
        for (int i = 0; i < 100; i++)
        {
            for (int j = 0; j < 20; j++)
            {
                sheet.Cells[i, j].PutValue($"R{i + 1}C{j + 1}");
            }
        }

        // Freeze the top 3 rows and the leftmost 2 columns.
        // row = 3, column = 2 specify the cell just below/right of the frozen area.
        // freezedRows = 3, freezedColumns = 2 define how many rows/columns stay visible.
        sheet.FreezePanes(3, 2, 3, 2);

        // Save the workbook
        workbook.Save("FreezePanesExample.xlsx");
    }
}