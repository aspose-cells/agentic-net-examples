using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsStyleSizeDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook wb = new Workbook();

            // Access the first worksheet
            Worksheet sheet = wb.Worksheets[0];

            // Add sample data with different styles to generate unused styles later
            for (int i = 0; i < 10; i++)
            {
                Cell cell = sheet.Cells[i, 0];
                cell.PutValue($"Item {i + 1}");

                // Create a unique style for each cell
                Style style = wb.CreateStyle();
                style.Font.Name = "Arial";
                style.Font.Size = 10 + i;
                style.Font.IsBold = i % 2 == 0;
                cell.SetStyle(style);
            }

            // Delete some rows to leave unused styles in the workbook
            sheet.Cells.DeleteRows(5, 5);

            // Save the workbook before removing unused styles
            string beforePath = "BeforeRemoveUnusedStyles.xlsx";
            wb.Save(beforePath);

            // Measure file size before removal
            long sizeBefore = new FileInfo(beforePath).Length;
            Console.WriteLine($"File size before removing unused styles: {sizeBefore} bytes");

            // Remove all unused styles
            wb.RemoveUnusedStyles();

            // Save the workbook after removing unused styles
            string afterPath = "AfterRemoveUnusedStyles.xlsx";
            wb.Save(afterPath);

            // Measure file size after removal
            long sizeAfter = new FileInfo(afterPath).Length;
            Console.WriteLine($"File size after removing unused styles: {sizeAfter} bytes");

            // Display reduction impact
            long reduction = sizeBefore - sizeAfter;
            Console.WriteLine($"Size reduction: {reduction} bytes ({(reduction * 100.0 / sizeBefore):F2}% decrease)");
        }
    }
}