// Title: Export Excel to PDF with Aspose.Cells (.NET) – Limit Output to First 10 Pages (C#)
// Description: This C# snippet shows how to open an Excel workbook using Aspose.Cells, set PdfSaveOptions.PageCount to 10, and save the file as a PDF. The resulting PDF contains only the initial ten pages of the original workbook, useful for previews or size‑reduced exports.
// Keywords: Aspose.Cells | PdfSaveOptions | PageCount | C# | .NET | limit PDF pages | Excel to PDF | first 10 pages | PDF preview | reduce PDF size
// Common Searches: Aspose.Cells limit PDF to first 10 pages C# | PdfSaveOptions.PageCount example | How to export only first ten pages of Excel as PDF using Aspose.Cells | Create PDF preview of Excel workbook with Aspose.Cells
// Developer Intent: Apply PdfSaveOptions.PageCount to restrict the page range when converting an Excel workbook to PDF with Aspose.Cells.
// Use Cases: Generate a quick preview PDF for a large report without exporting the whole document. | Provide a sample PDF to clients while keeping the remaining pages confidential. | Cut down file size for email distribution by exporting only the opening pages.
// AI Prompts: Show me C# code to export an Excel file to PDF using Aspose.Cells and limit the output to ten pages. | How does PdfSaveOptions.PageCount affect PDF generation in Aspose.Cells for .NET? | Give an example of creating a PDF preview of the first 10 pages from an Excel workbook with Aspose.Cells.

using System;
using Aspose.Cells;

// This C# snippet shows how to open an Excel workbook using Aspose.Cells, set PdfSaveOptions.PageCount to 10, and save the file as a PDF. The resulting PDF contains only the initial ten pages of the original workbook, useful for previews or size‑reduced exports.
class LimitPdfPages
{
    static void Main()
    {
        // Load the Excel workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Create PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Limit the output to the first ten pages
        pdfOptions.PageCount = 10;

        // Save the workbook as PDF using the configured options
        workbook.Save("output.pdf", pdfOptions);
    }
}
