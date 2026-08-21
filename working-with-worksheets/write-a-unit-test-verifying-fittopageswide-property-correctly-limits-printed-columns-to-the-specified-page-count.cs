// Title: C# Unit Test for Aspose.Cells FitToPagesWide Pagination
// Description: Demonstrates how to verify that the FitToPagesWide setting in Aspose.Cells correctly limits the number of printed columns per page. The test creates a workbook, fills a row with 30 columns, defines a print area, renders the sheet with FitToPagesWide = 1 and = 2, and asserts that the resulting page counts are 1 and 2 respectively.
// Keywords: Aspose.Cells | FitToPagesWide | C# unit test | SheetRender page count | pagination verification | .NET Excel printing | MSTest | xUnit | NUnit
// Common Searches: Aspose.Cells FitToPagesWide unit test example | how to test page count after setting FitToPagesWide | C# verify printed pages with Aspose.Cells | unit testing Excel pagination in .NET | SheetRender page count assertion
// Developer Intent: Provide a ready‑to‑run unit test that confirms the FitToPagesWide property restricts printed columns to the expected number of pages.
// Use Cases: Automated regression testing of worksheet pagination after library upgrades. | Ensuring generated reports fit within a single page width for printing. | Validating that a two‑page width layout splits wide data correctly. | Integrating pagination checks into CI pipelines for .NET Excel solutions.
// AI Prompts: Create an MSTest method that builds a workbook, populates 30 columns, sets PrintArea, applies FitToPagesWide = 1 and 2, renders with SheetRender, and asserts PageCount equals 1 and 2. | Write an xUnit test for Aspose.Cells that verifies FitToPagesWide pagination while FitToPagesTall remains zero, including proper disposal of Workbook and SheetRender objects. | Generate a NUnit test example that logs the observed page counts for FitToPagesWide values and fails with a clear message if the counts differ from expected.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsExamples
{
    // Demonstrates how to verify that the FitToPagesWide setting in Aspose.Cells correctly limits the number of printed columns per page. The test creates a workbook, fills a row with 30 columns, defines a print area, renders the sheet with FitToPagesWide = 1 and = 2, and asserts that the resulting page counts are 1 and 2 respectively.
    public class FitToPagesWideDemo
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                var workbook = new Workbook();
                var sheet = workbook.Worksheets[0];

                // Populate a single row with many columns (30 columns)
                for (int col = 0; col < 30; col++)
                {
                    sheet.Cells[0, col].PutValue($"Col{col + 1}");
                }

                // Define the print area to include all populated columns
                sheet.PageSetup.PrintArea = "A1:AD1"; // AD is the 30th column

                // Scenario 1: Fit all columns into a single page width
                sheet.PageSetup.FitToPagesWide = 1;
                sheet.PageSetup.FitToPagesTall = 0; // Let height adjust automatically

                var options = new ImageOrPrintOptions();
                var renderOnePage = new SheetRender(sheet, options);
                int pageCountWhenWide1 = renderOnePage.PageCount;

                // Scenario 2: Allow the content to span two pages wide
                sheet.PageSetup.FitToPagesWide = 2;
                // Recreate the renderer to recalculate page count
                var renderTwoPages = new SheetRender(sheet, options);
                int pageCountWhenWide2 = renderTwoPages.PageCount;

                // Output results
                Console.WriteLine($"FitToPagesWide = 1 => Page count: {pageCountWhenWide1}");
                Console.WriteLine($"FitToPagesWide = 2 => Page count: {pageCountWhenWide2}");

                // Simple verification
                if (pageCountWhenWide1 == 1 && pageCountWhenWide2 == 2)
                {
                    Console.WriteLine("Test passed.");
                }
                else
                {
                    Console.WriteLine("Test failed.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
