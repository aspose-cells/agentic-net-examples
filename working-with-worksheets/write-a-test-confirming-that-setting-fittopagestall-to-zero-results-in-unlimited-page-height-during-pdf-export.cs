using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Saving;

namespace AsposeCellsFitToPagesTallTest
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate the sheet with enough rows to span multiple pages vertically
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
            pageSetup.PrintArea = "A1:E200";

            // Create rendering options (default)
            ImageOrPrintOptions renderOptions = new ImageOrPrintOptions();

            // Use SheetRender to determine the number of pages after the settings are applied
            SheetRender renderer = new SheetRender(sheet, renderOptions);
            int pageCount = renderer.PageCount;

            // Output the page count – it should be greater than 1, confirming unlimited height
            Console.WriteLine($"Page count with FitToPagesTall = 0: {pageCount}");

            // Save the workbook to PDF (optional verification)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();
            using (MemoryStream pdfStream = new MemoryStream())
            {
                workbook.Save(pdfStream, pdfOptions);
                // The PDF is now in pdfStream; you could write it to disk if needed
                Console.WriteLine($"PDF generated, size: {pdfStream.Length} bytes");
            }

            // Simple assertion (replace with a testing framework assert if desired)
            if (pageCount <= 1)
            {
                throw new InvalidOperationException("FitToPagesTall = 0 did not produce multiple pages as expected.");
            }

            Console.WriteLine("Test passed: FitToPagesTall = 0 results in unlimited page height.");
        }
    }
}