using System;
using Aspose.Cells;

namespace AsposeCellsPrintFitDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data to demonstrate scaling
            for (int row = 0; row < 50; row++)
            {
                for (int col = 0; col < 10; col++)
                {
                    sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Get the PageSetup object for the worksheet
            PageSetup pageSetup = sheet.PageSetup;

            // Define the desired printable area (optional)
            pageSetup.PrintArea = "A1:J50";

            // Fit the worksheet to a specific number of pages wide and tall.
            // For example, fit to 1 page wide and 2 pages tall.
            pageSetup.SetFitToPages(1, 2);

            // Alternatively, you can set the properties directly:
            // pageSetup.FitToPagesWide = 1;   // fit all columns on one page
            // pageSetup.FitToPagesTall = 2;   // allow up to two pages in height

            // Optionally set a custom paper size (width and height in inches)
            // Here we set a standard A5 size as an example (5.83 x 8.27 inches)
            pageSetup.CustomPaperSize(5.83, 8.27);

            // Save the workbook to a file
            workbook.Save("PrintFitDemo.xlsx");
        }
    }
}