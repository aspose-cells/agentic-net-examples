using System;
using Aspose.Cells;

// Author: Aspose.Cells .NET example – sets column width in pixels

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Set the width of column 3 (zero‑based index 2) to 150 pixels
        cells.SetColumnWidthPixel(2, 150);

        // Save the workbook
        workbook.Save("ColumnWidthPixelDemo.xlsx");
    }
}