using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsApostrophePdfDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Enable QuotePrefixToStyle so that a leading apostrophe is treated as a style flag
            workbook.Settings.QuotePrefixToStyle = true;

            // Put a value that starts with a single quote.
            // The apostrophe will not become part of the cell text; instead, the QuotePrefix style will be set.
            Cell cell = sheet.Cells["A1"];
            cell.PutValue("'Hello World");

            // Verify that the QuotePrefix style is applied
            bool isQuotePrefix = cell.GetStyle().QuotePrefix;
            Console.WriteLine($"QuotePrefix style applied: {isQuotePrefix}"); // Expected: True

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Ensure the default font checking is enabled so Unicode characters render correctly
                CheckWorkbookDefaultFont = true,

                // Optional: keep document structure for better accessibility
                ExportDocumentStructure = true
            };

            // Save the workbook to PDF
            string pdfPath = "ApostropheDemo.pdf";
            workbook.Save(pdfPath, pdfOptions);
            Console.WriteLine($"Workbook saved to PDF: {pdfPath}");

            // At this point, opening the PDF should show the text "Hello World"
            // without the leading apostrophe, confirming that the apostrophe was handled correctly.
        }
    }
}