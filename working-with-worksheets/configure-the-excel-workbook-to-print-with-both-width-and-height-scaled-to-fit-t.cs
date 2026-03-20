using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsFitToPagesDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add some sample data (optional, just to have content)
            for (int row = 0; row < 20; row++)
            {
                for (int col = 0; col < 5; col++)
                {
                    worksheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Get the PageSetup object for the worksheet
            PageSetup pageSetup = worksheet.PageSetup;

            // Ensure scaling is based on FitToPages rather than percent scale
            pageSetup.IsPercentScale = false;

            // Set the worksheet to fit to a specific number of pages (both width and height)
            // Example: fit to 2 pages wide and 3 pages tall
            pageSetup.SetFitToPages(2, 3);

            // Optionally, you can also set the properties directly:
            // pageSetup.FitToPagesWide = 2;
            // pageSetup.FitToPagesTall = 3;

            // Save the workbook to a file
            workbook.Save("FitToPagesDemo.xlsx", SaveFormat.Xlsx);

            // Demonstrate that the settings have been applied
            Console.WriteLine($"IsPercentScale: {pageSetup.IsPercentScale}");
            Console.WriteLine($"FitToPagesWide: {pageSetup.FitToPagesWide}");
            Console.WriteLine($"FitToPagesTall: {pageSetup.FitToPagesTall}");
        }
    }
}