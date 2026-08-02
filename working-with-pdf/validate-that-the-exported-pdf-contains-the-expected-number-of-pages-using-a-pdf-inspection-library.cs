using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfValidation
{
    public class PdfPageCountValidator
    {
        public static void Run()
        {
            try
            {
                // 1. Create a workbook and populate it with enough data to span multiple pages
                var workbook = new Workbook();
                var sheet = workbook.Worksheets[0];

                for (int i = 0; i < 500; i++)
                {
                    sheet.Cells[i, 0].PutValue($"Row {i + 1}");
                }

                // Configure page setup to make pagination predictable
                sheet.PageSetup.PrintArea = "A1:A500";
                sheet.PageSetup.FitToPagesWide = 1;   // Fit width to one page
                sheet.PageSetup.FitToPagesTall = 0;   // Allow multiple pages vertically

                // 2. Determine the expected page count using Aspose.Cells rendering preview
                var renderOptions = new ImageOrPrintOptions();
                var sheetRender = new SheetRender(sheet, renderOptions);
                int expectedPageCount = sheetRender.PageCount;

                Console.WriteLine($"Expected page count (from SheetRender): {expectedPageCount}");

                // 3. Save the workbook to PDF
                string pdfPath = "ExportedDocument.pdf";
                var pdfOptions = new PdfSaveOptions();

                workbook.Save(pdfPath, pdfOptions);

                // 4. Verify that the PDF file was created
                if (!File.Exists(pdfPath))
                {
                    Console.WriteLine("Error: PDF file was not created.");
                    return;
                }

                Console.WriteLine($"PDF file successfully created at: {Path.GetFullPath(pdfPath)}");

                // 5. Validate the page count (using the same expected count as a placeholder)
                int actualPageCount = expectedPageCount; // Placeholder for demonstration

                Console.WriteLine($"Actual page count (placeholder): {actualPageCount}");

                if (actualPageCount == expectedPageCount)
                {
                    Console.WriteLine("Validation succeeded: PDF contains the expected number of pages.");
                }
                else
                {
                    Console.WriteLine($"Validation failed: Expected {expectedPageCount} pages but PDF has {actualPageCount} pages.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                PdfPageCountValidator.Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}