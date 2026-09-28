// Title: C# unit test for Aspose.Cells PageSetup.FitToPagesWide to ensure printed columns are limited to a specific number of pages
// AI Prompts: Write a C# NUnit test that populates a worksheet with many columns, sets PageSetup.FitToPagesWide, renders the sheet with SheetRender, and asserts that the total page count is a multiple of the FitToPagesWide value. | Create a C# xUnit test that builds a single‑row worksheet, applies PageSetup.FitToPagesWide, renders it, and verifies the rendered page count equals the expected horizontal page count. | Generate a C# test method that catches an InvalidOperationException when the rendered page count does not satisfy the FitToPagesWide constraint and reports the failure.
// Common Searches: how to assert Aspose.Cells FitToPagesWide pagination in a C# test | C# Aspose.Cells verify horizontal page limit using SheetRender | unit testing PageSetup.FitToPagesWide with multiple columns in .NET | Aspose.Cells render worksheet and check printed pages count programmatically
// Tags: Aspose.Cells FitToPagesWide pagination test | C# SheetRender rendering verification | Aspose.Cells horizontal page limit | C# workbook unit testing Aspose.Cells | FitToPagesTall zero setting behavior

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsExamples
{
    // The example creates a workbook, fills it with 100 columns and 50 rows, sets PageSetup.FitToPagesWide to 2 and FitToPagesTall to 0, renders the sheet with SheetRender to obtain the page count, and throws an exception if the total pages are not divisible by 2. It also renders a single‑row worksheet with the same FitToPagesWide setting and confirms that exactly two pages are produced.
    public class FitToPagesWideDemo
    {
        public static void Main()
        {
            try
            {
                RunFitToPagesWideTest();
                Console.WriteLine("Test passed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Test failed: {ex.Message}");
            }
        }

        private static void RunFitToPagesWideTest()
        {
            // Create a new workbook and get the first worksheet
            var workbook = new Workbook();
            var sheet = workbook.Worksheets[0];

            // Populate the sheet with data across many columns (100 columns, 50 rows)
            for (int col = 0; col < 100; col++)
            {
                for (int row = 0; row < 50; row++)
                {
                    sheet.Cells[row, col].PutValue($"R{row}C{col}");
                }
            }

            // Set FitToPagesWide to 2 (limit horizontal pages to 2)
            sheet.PageSetup.FitToPagesWide = 2;
            // Allow unlimited vertical pages
            sheet.PageSetup.FitToPagesTall = 0;

            // Render the sheet to determine the total number of printed pages
            var renderOptions = new ImageOrPrintOptions { OnePagePerSheet = false };
            var sheetRender = new SheetRender(sheet, renderOptions);
            int totalPages = sheetRender.PageCount;

            // Verify total pages divisible by FitToPagesWide
            if (totalPages % 2 != 0)
                throw new InvalidOperationException("Total pages should be divisible by FitToPagesWide value.");

            // Additional verification: render a sheet that contains only a single row.
            var tempWorkbook = new Workbook();
            var tempSheet = tempWorkbook.Worksheets[0];
            for (int col = 0; col < 100; col++)
            {
                tempSheet.Cells[0, col].PutValue($"C{col}");
            }
            tempSheet.PageSetup.FitToPagesWide = 2;
            tempSheet.PageSetup.FitToPagesTall = 0;
            var tempRender = new SheetRender(tempSheet, renderOptions);
            int tempPages = tempRender.PageCount;

            if (tempPages != 2)
                throw new InvalidOperationException("FitToPagesWide should limit horizontal pages to 2.");
        }
    }
}
