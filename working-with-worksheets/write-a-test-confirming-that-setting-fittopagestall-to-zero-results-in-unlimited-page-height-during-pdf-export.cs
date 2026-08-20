// Title: C# Test: FitToPagesTall = 0 Gives Unlimited Height on PDF Export with Aspose.Cells
// Description: Creates a workbook, fills 200 rows, sets PageSetup.FitToPagesWide = 1 and FitToPagesTall = 0, renders the sheet, checks that SheetRender.PageCount > 1, and optionally saves the result as a PDF to confirm unlimited vertical pagination.
// Keywords: Aspose.Cells | FitToPagesTall | C# | .NET | PDF export | page count verification | unlimited height | SheetRender | PageSetup | automated test
// Common Searches: Aspose.Cells FitToPagesTall zero test | PDF export unlimited page height Aspose.Cells | verify page count after setting FitToPagesTall to 0 | C# Aspose.Cells pagination test | how to check unlimited height in PDF with Aspose.Cells
// Developer Intent: Validate that setting PageSetup.FitToPagesTall to 0 produces multiple PDF pages (unlimited vertical length) instead of a single page.
// Use Cases: Automated regression test for pagination behavior when unlimited height is required. | Ensuring large worksheets split across several PDF pages while keeping all columns on one page width. | Generating PDFs where column width is fixed but rows flow across as many pages as needed.
// AI Prompts: Generate an NUnit test that asserts SheetRender.PageCount > 1 when PageSetup.FitToPagesTall = 0. | Provide a C# snippet that programmatically confirms a PDF created with FitToPagesTall = 0 contains more than one page. | Explain the impact of FitToPagesTall = 0 on Aspose.Cells pagination and describe a method to verify it via code.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsFitToPagesTallTest
{
    // Creates a workbook, fills 200 rows, sets PageSetup.FitToPagesWide = 1 and FitToPagesTall = 0, renders the sheet, checks that SheetRender.PageCount > 1, and optionally saves the result as a PDF to confirm unlimited vertical pagination.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate many rows to ensure the content spans multiple pages vertically
                for (int row = 0; row < 200; row++)
                {
                    for (int col = 0; col < 5; col++)
                    {
                        sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                    }
                }

                // Configure page setup:
                // Fit all columns on one page (wide = 1)
                // Unlimited page height (tall = 0)
                PageSetup pageSetup = sheet.PageSetup;
                pageSetup.FitToPagesWide = 1;
                pageSetup.FitToPagesTall = 0; // Unlimited height

                // Create rendering options (default options are sufficient)
                ImageOrPrintOptions renderOptions = new ImageOrPrintOptions();

                // Render the sheet to obtain page count
                SheetRender renderer = new SheetRender(sheet, renderOptions);
                int pageCount = renderer.PageCount;

                // Output the page count for verification
                Console.WriteLine($"Page count with FitToPagesTall = 0: {pageCount}");

                // Verify that more than one page is generated (unlimited height)
                if (pageCount > 1)
                {
                    Console.WriteLine("Test passed: FitToPagesTall = 0 results in unlimited page height.");
                }
                else
                {
                    Console.WriteLine("Test failed: Expected multiple pages but got a single page.");
                }

                // Save the workbook as PDF to visually confirm the result (optional)
                PdfSaveOptions pdfOptions = new PdfSaveOptions();
                workbook.Save("FitToPagesTall_Unlimited.pdf", pdfOptions);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
