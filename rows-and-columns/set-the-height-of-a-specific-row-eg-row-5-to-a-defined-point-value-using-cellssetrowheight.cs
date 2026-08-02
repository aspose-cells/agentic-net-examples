using System;
using Aspose.Cells;

// Author: Aspose.Cells .NET example – sets row height using Cells.SetRowHeight
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Set the height of row 5 (zero‑based index) to 30 points
        cells.SetRowHeight(5, 30.0);

        // Verify the height (optional)
        Console.WriteLine("Row 5 height: " + cells.GetRowHeight(5));

        // Save the workbook to a file
        workbook.Save("RowHeightDemo.xlsx");
    }
}