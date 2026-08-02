using System;
using Aspose.Cells;

public class SetColumnWidthPixelExample
{
    public static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Set the width of column B (index 1) to 150 pixels
        cells.SetColumnWidthPixel(1, 150);

        // Add numeric data to column B
        cells[0, 1].PutValue(123);      // Row 1, Column B
        cells[1, 1].PutValue(4567);    // Row 2, Column B
        cells[2, 1].PutValue(89.01);   // Row 3, Column B

        // Save the workbook to a file
        workbook.Save("ColumnWidthPixelDemo.xlsx");
    }
}