// Title: C# Unit Test for Worksheet.PageSetup.FitToPagesWide Pagination in Aspose.Cells
// Description: This example creates a workbook, fills cells A1:J100, defines a print area, sets FitToPagesWide = 2 and FitToPagesTall = 1, renders the sheet with SheetRender, and asserts that the resulting PageCount is two and that PageScale is a valid positive percentage.
// Keywords: Aspose.Cells | FitToPagesWide | unit test | C# | SheetRender pagination | PageCount verification | PrintArea scaling
// Common Searches: Aspose.Cells unit test FitToPagesWide | verify printed page count C# Aspose | how to test pagination with SheetRender | check page scale after setting FitToPagesTall | C# Aspose.Cells pagination example
// Developer Intent: Confirm that the FitToPagesWide setting restricts the rendered worksheet to the specified number of pages.
// Use Cases: Automated regression testing of pagination after library upgrades. | Ensuring generated reports fit within a predefined column width per printed page. | Validating that page scaling remains within acceptable limits when using fit‑to‑pages options.
// AI Prompts: Generate an MSTest method that asserts Worksheet.PageSetup.FitToPagesWide limits the page count to two using Aspose.Cells. | Create a NUnit test that populates a workbook, applies FitToPagesWide/Tall, renders with SheetRender, and checks PageCount and PageScale. | Write an xUnit test verifying pagination and scaling when FitToPagesWide is set to 3 and FitToPagesTall to 1.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsTests
{
    // This example creates a workbook, fills cells A1:J100, defines a print area, sets FitToPagesWide = 2 and FitToPagesTall = 1, renders the sheet with SheetRender, and asserts that the resulting PageCount is two and that PageScale is a valid positive percentage.
    public class FitToPagesWideTests
    {
        public static void Main()
        {
            try
            {
                RunTest();
                Console.WriteLine("Test completed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void RunTest()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate the worksheet with data spanning many columns and rows (100 rows, 10 columns)
                for (int row = 0; row < 100; row++)
                {
                    for (int col = 0; col < 10; col++)
                    {
                        worksheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                    }
                }

                // Define the print area covering all populated cells
                worksheet.PageSetup.PrintArea = "A1:J100";

                // Set FitToPagesWide to 2 pages and FitToPagesTall to 1 page
                worksheet.PageSetup.FitToPagesWide = 2;
                worksheet.PageSetup.FitToPagesTall = 1;

                // Create rendering options (default options are sufficient for this test)
                ImageOrPrintOptions options = new ImageOrPrintOptions();

                // Render the worksheet to evaluate pagination
                SheetRender render = new SheetRender(worksheet, options);
                // Verify that the total page count equals 2
                if (render.PageCount != 2)
                {
                    throw new InvalidOperationException($"FitToPagesWide did not limit the worksheet to the expected number of pages. Expected 2, got {render.PageCount}.");
                }

                // Verify that the calculated page scale reflects the fit-to-pages setting
                double pageScale = render.PageScale;
                if (pageScale <= 0 || pageScale > 100)
                {
                    throw new InvalidOperationException($"PageScale should be a positive percentage when using FitToPages settings. Actual: {pageScale}.");
                }

                Console.WriteLine($"PageCount: {render.PageCount}, PageScale: {pageScale}%");
            }
            catch (Exception ex)
            {
                // Propagate any errors to the caller for unified handling
                throw new ApplicationException("RunTest failed.", ex);
            }
        }
    }
}
