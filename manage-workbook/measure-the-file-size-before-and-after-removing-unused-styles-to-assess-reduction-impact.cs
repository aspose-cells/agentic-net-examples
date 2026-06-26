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
            Cells cells = sheet.Cells;

            // Add sample data with different styles to generate unused styles later
            for (int i = 0; i < 20; i++)
            {
                Cell cell = cells[i, 0];
                cell.PutValue($"Item {i + 1}");

                // Create a unique style for each cell
                Style style = wb.CreateStyle();
                style.Font.Name = "Arial";
                style.Font.Size = 10 + i;               // Vary font size
                style.Font.IsBold = i % 2 == 0;         // Alternate bold
                style.Font.Color = i % 3 == 0 ? System.Drawing.Color.Red : System.Drawing.Color.Blue;
                cell.SetStyle(style);
            }

            // Delete a range of rows to leave some styles unused
            sheet.Cells.DeleteRows(10, 5); // Rows 10-14 removed, their styles become unused

            // Save workbook before removing unused styles
            string beforePath = "BeforeRemoveUnusedStyles.xlsx";
            wb.Save(beforePath);

            // Measure file size before removal
            long sizeBefore = new FileInfo(beforePath).Length;
            Console.WriteLine($"File size before removing unused styles: {sizeBefore} bytes");

            // Remove all unused styles
            wb.RemoveUnusedStyles();

            // Save workbook after removing unused styles
            string afterPath = "AfterRemoveUnusedStyles.xlsx";
            wb.Save(afterPath);

            // Measure file size after removal
            long sizeAfter = new FileInfo(afterPath).Length;
            Console.WriteLine($"File size after removing unused styles: {sizeAfter} bytes");

            // Display reduction information
            long reduction = sizeBefore - sizeAfter;
            double percent = sizeBefore > 0 ? (double)reduction / sizeBefore * 100 : 0;
            Console.WriteLine($"Size reduction: {reduction} bytes ({percent:F2}%)");
        }
    }
}