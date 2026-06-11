using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsTests
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate the worksheet with data spanning many columns (e.g., 50 columns) and a few rows
                int totalColumns = 50;
                int totalRows = 5;
                for (int row = 0; row < totalRows; row++)
                {
                    for (int col = 0; col < totalColumns; col++)
                    {
                        worksheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                    }
                }

                // Define the print area to include all populated cells
                string lastColumn = CellsHelper.ColumnIndexToName(totalColumns - 1);
                worksheet.PageSetup.PrintArea = $"A1:{lastColumn}{totalRows}";

                // Set FitToPagesWide to 2 pages and let height adjust automatically (FitToPagesTall = 0)
                worksheet.PageSetup.FitToPagesWide = 2;
                worksheet.PageSetup.FitToPagesTall = 0;

                // Create rendering options (default options are sufficient for this test)
                ImageOrPrintOptions options = new ImageOrPrintOptions();

                // Render the worksheet
                SheetRender render = new SheetRender(worksheet, options);
                int actualPageCount = render.PageCount;

                // Output the result
                Console.WriteLine($"Actual page count: {actualPageCount}");
                Console.WriteLine(actualPageCount == 2
                    ? "Test passed: Expected 2 pages wide."
                    : $"Test failed: Expected 2 pages wide, but got {actualPageCount} pages.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}