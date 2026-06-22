using System;
using Aspose.Cells;

namespace AsposeCellsColumnWidthAndFreezeDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Set specific column widths (in character units)
            // Column A (index 0) width = 20 characters
            cells.SetColumnWidth(0, 20);
            // Column B (index 1) width = 30 characters
            cells.SetColumnWidth(1, 30);
            // Column C (index 2) width = 15 characters
            cells.SetColumnWidth(2, 15);

            // Optionally add some data to visualize the widths
            cells["A1"].PutValue("Column A");
            cells["B1"].PutValue("Column B with longer header");
            cells["C1"].PutValue("Col C");

            // Freeze the first three columns.
            // Freeze at cell D1 (column index 3) with 0 frozen rows and 3 frozen columns.
            worksheet.FreezePanes("D1", 0, 3);

            // Save the workbook
            workbook.Save("ColumnWidthsAndFreeze.xlsx");
        }
    }
}