using System;
using Aspose.Cells;

class SetRowHeightsAndFreeze
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Set custom heights for the first two rows (in points)
        cells.SetRowHeight(0, 30); // Row 1 height = 30 points
        cells.SetRowHeight(1, 40); // Row 2 height = 40 points

        // Add sample data to visualize the rows
        sheet.Cells["A1"].PutValue("First row");
        sheet.Cells["A2"].PutValue("Second row");

        // Freeze the first two rows while preserving their custom heights
        // Freeze at cell A3 (row index 2) with 2 frozen rows and 0 frozen columns
        sheet.FreezePanes("A3", 2, 0);

        // Save the workbook
        workbook.Save("RowHeightsAndFreeze.xlsx");
    }
}