// Title: C# unit test to verify that Worksheet.PageSetup.FitToPagesWide limits printed columns to two pages with Aspose.Cells
// Common Searches: how to write a C# unit test for FitToPagesWide in Aspose.Cells | Aspose.Cells verify page count when fitting worksheet to 2 pages wide | C# test workbook rendering page count with FitToPagesWide set to 2 | unit testing Aspose.Cells page setup FitToPagesWide property | check printed column pagination using Aspose.Cells FitToPagesWide in C#
// Tags: Worksheet.PageSetup.FitToPagesWide verification | Aspose.Cells pagination using WorkbookRender | FitToPagesTall zero unlimited height setting | C# workbook rendering page count test | printed columns limited to two pages Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsTests
{
    // The example creates a workbook, fills the first worksheet with 200 columns and 10 rows, sets PageSetup.FitToPagesWide to 2 and FitToPagesTall to 0 (unlimited height), renders the workbook with WorkbookRender, and asserts that the resulting page count is exactly 2, confirming that FitToPagesWide correctly restricts printed columns to two pages.
    public class FitToPagesWideDemo
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                var workbook = new Workbook();
                var sheet = workbook.Worksheets[0];

                // Populate the worksheet with a large number of columns (200) and a few rows (10)
                const int totalColumns = 200;
                const int totalRows = 10;
                for (int row = 0; row < totalRows; row++)
                {
                    for (int col = 0; col < totalColumns; col++)
                    {
                        // Fill each cell with a simple value to ensure content exists
                        sheet.Cells[row, col].PutValue($"R{row}C{col}");
                    }
                }

                // Configure the page setup:
                // - Fit the printed output to 2 pages wide.
                // - Do not limit the height (set to 0, which means unlimited pages tall).
                sheet.PageSetup.FitToPagesWide = 2;
                sheet.PageSetup.FitToPagesTall = 0; // 0 = unlimited height

                // Use the WorkbookRender to calculate the number of pages that will be generated.
                var renderer = new WorkbookRender(workbook, new ImageOrPrintOptions());

                // The renderer calculates the total page count based on the FitToPagesWide/Tall settings.
                int totalPageCount = renderer.PageCount;

                // Verify the expected page count (2 pages wide, unlimited height should result in 2 pages total).
                if (totalPageCount == 2)
                {
                    Console.WriteLine("Test passed: FitToPagesWide limited columns to the expected number of pages.");
                }
                else
                {
                    Console.WriteLine($"Test failed: Expected 2 pages, but got {totalPageCount} pages.");
                }
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors and display them.
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
