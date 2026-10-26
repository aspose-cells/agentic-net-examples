// Title: C# unit test to confirm Worksheet.PageSetup.FitToPagesWide limits printed columns to two pages with Aspose.Cells
// Common Searches: how to write a unit test for FitToPagesWide in Aspose.Cells C# | verify printed page count after setting FitToPagesWide to 2 using Aspose.Cells | Aspose.Cells C# example asserting page count with WorkbookRender | unit testing worksheet page setup column limit in Aspose.Cells | using WorkbookRender in a test to check FitToPagesWide behavior
// Tags: Aspose.Cells FitToPagesWide page count test | Worksheet PageSetup column limit verification | C# WorkbookRender page count assertion | unit testing printed pages wide setting | FitToPagesWide property validation Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsExamples
{
    // The example creates a workbook, fills the first row with 100 columns, sets PageSetup.FitToPagesWide to 2 and FitToPagesTall to 0, renders the workbook with WorkbookRender, and asserts that the resulting PageCount equals 2, demonstrating how to unit‑test the FitToPagesWide setting in C#.
    public class FitToPagesWideDemo
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                var workbook = new Workbook();
                var worksheet = workbook.Worksheets[0];

                // Populate 100 columns with sample data in the first row
                for (int col = 0; col < 100; col++)
                {
                    worksheet.Cells[0, col].PutValue($"Col{col}");
                }

                // Add a few rows of data to keep the vertical page count to 1
                for (int row = 1; row < 10; row++)
                {
                    worksheet.Cells[row, 0].PutValue($"Row{row}");
                }

                // Configure the page setup:
                // - Limit the printed output to 2 pages wide
                // - Allow unlimited pages tall (0 means no limit)
                worksheet.PageSetup.FitToPagesWide = 2;
                worksheet.PageSetup.FitToPagesTall = 0;

                // Render the workbook to determine the total number of pages
                var renderOptions = new ImageOrPrintOptions();
                var renderer = new WorkbookRender(workbook, renderOptions);

                // Verify that the total page count matches the expected value (2)
                if (renderer.PageCount == 2)
                {
                    Console.WriteLine("Test passed: Page count is 2 as expected.");
                }
                else
                {
                    Console.WriteLine($"Test failed: Expected page count 2, but got {renderer.PageCount}.");
                }
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors and display them
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
