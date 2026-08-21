// Title: Export Aspose.Cells Workbook to PDF with Leading Apostrophe Preserved (QuotePrefixToStyle)
// Description: C# example that creates a workbook, enables QuotePrefixToStyle so a leading apostrophe is stored as a style flag, writes "'SampleText" to A1, verifies the style, configures PdfSaveOptions (CheckWorkbookDefaultFont), and saves the file as a PDF where the apostrophe remains visible.
// Keywords: Aspose.Cells PDF export leading apostrophe | QuotePrefixToStyle .NET | PdfSaveOptions CheckWorkbookDefaultFont | preserve apostrophe in PDF | Aspose.Cells C# PDF conversion
// Common Searches: how to keep a leading apostrophe when converting Aspose.Cells to PDF | QuotePrefixToStyle PDF export Aspose.Cells example | Aspose.Cells PDF output missing leading apostrophe | C# save workbook as PDF with apostrophe displayed
// Developer Intent: Generate a PDF from an Aspose.Cells workbook that accurately displays any leading apostrophe in cell values.
// Use Cases: Export financial reports where account numbers start with an apostrophe. | Create product catalogs that include SKU codes prefixed by a single quote. | Validate data entry forms that require visible leading apostrophes in the final PDF.
// AI Prompts: Provide a step‑by‑step guide to enable QuotePrefixToStyle and export a workbook to PDF while preserving leading apostrophes. | Write a unit test in C# that opens the generated PDF and asserts that cell A1 contains the leading apostrophe. | Explain how to apply the same apostrophe‑preserving settings across multiple worksheets in a single PDF export.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsApostrophePdfDemo
{
    // C# example that creates a workbook, enables QuotePrefixToStyle so a leading apostrophe is stored as a style flag, writes "'SampleText" to A1, verifies the style, configures PdfSaveOptions (CheckWorkbookDefaultFont), and saves the file as a PDF where the apostrophe remains visible.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Enable QuotePrefixToStyle so that a leading apostrophe is stored as a style flag
            workbook.Settings.QuotePrefixToStyle = true;

            // Put a string that starts with a single quote (apostrophe) into cell A1
            Cell cell = sheet.Cells["A1"];
            cell.PutValue("'SampleText");

            // Verify that the QuotePrefix style is applied (optional, for debugging)
            bool isQuotePrefix = cell.GetStyle().QuotePrefix;
            Console.WriteLine($"QuotePrefix applied to A1: {isQuotePrefix}");

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Ensure the default font is checked to render Unicode correctly
                CheckWorkbookDefaultFont = true
            };

            // Save the workbook as PDF
            string pdfPath = "ApostropheDemo.pdf";
            workbook.Save(pdfPath, pdfOptions);
            Console.WriteLine($"Workbook saved to PDF: {pdfPath}");

            // Confirmation message – the leading apostrophe should appear in the PDF output
            Console.WriteLine("Please open the PDF to verify that the leading apostrophe is displayed as intended.");
        }
    }
}
