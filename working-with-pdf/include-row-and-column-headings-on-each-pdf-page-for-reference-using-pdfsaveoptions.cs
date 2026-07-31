// Title: Add Row and Column Headings to Every PDF Page with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to enable the worksheet PrintHeadings property, set an optional print area, and use PdfSaveOptions to export an Excel workbook to PDF so that each page displays the original row and column headings for easy reference.
// Keywords: Aspose.Cells PDF headings | PrintHeadings C# | PdfSaveOptions export Excel to PDF | row and column headings PDF | .NET Excel to PDF with headings | Aspose.Cells print area
// Common Searches: Aspose.Cells show row and column headings in PDF | C# export Excel to PDF with PrintHeadings | How to keep Excel headings on each PDF page | Set PrintArea and PrintHeadings before PDF conversion Aspose | Default PdfSaveOptions for headings Aspose.Cells
// Developer Intent: Generate a PDF from an Excel workbook that retains the worksheet’s row and column headings on every printed page.
// Use Cases: Create a product catalog PDF that includes Excel’s row/column labels for quick lookup. | Export a multi‑page financial statement while preserving sheet headings across all pages. | Produce a concise PDF report with a defined print area and visible headings for documentation.
// AI Prompts: Write C# code using Aspose.Cells to export a worksheet to PDF with row and column headings on each page and a custom print area. | Explain the interaction between the PrintHeadings property and PdfSaveOptions when converting Excel to PDF with Aspose.Cells. | Show how to change page orientation while keeping PrintHeadings enabled for PDF output in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfHeadingsExample
{
    // Demonstrates how to enable the worksheet PrintHeadings property, set an optional print area, and use PdfSaveOptions to export an Excel workbook to PDF so that each page displays the original row and column headings for easy reference.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data with headings
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Price");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(2.5);
            sheet.Cells["A3"].PutValue("Orange");
            sheet.Cells["B3"].PutValue(1.8);

            // Enable printing of row and column headings on each page
            sheet.PageSetup.PrintHeadings = true;

            // (Optional) Define the print area to limit what is exported
            sheet.PageSetup.PrintArea = "A1:B3";

            // Create PDF save options (default options are sufficient for headings)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as PDF with the specified options
            workbook.Save("ProductsWithHeadings.pdf", pdfOptions);
        }
    }
}
