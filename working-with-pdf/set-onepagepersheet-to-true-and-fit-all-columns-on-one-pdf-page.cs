// Title: Export a Wide Worksheet to a Single PDF Page with All Columns Using Aspose.Cells for .NET (C#)
// Description: This C# example creates a workbook, adds many columns, sets the worksheet page setup to fit the width to one page (FitToPagesWide = 1, FitToPagesTall = 0), enables PdfSaveOptions.OnePagePerSheet and AllColumnsInOnePagePerSheet, and saves the file as a PDF where each sheet appears on a single page containing all columns.
// Keywords: Aspose.Cells | C# | .NET | PdfSaveOptions | OnePagePerSheet | AllColumnsInOnePagePerSheet | FitToPagesWide | fit columns to one PDF page | export worksheet to PDF | single page PDF | wide worksheet PDF | page setup
// Common Searches: Aspose.Cells set OnePagePerSheet true | fit all columns on one PDF page Aspose.Cells | PdfSaveOptions AllColumnsInOnePagePerSheet example | C# export Excel to single-page PDF | force worksheet columns onto one PDF page | Aspose.Cells PDF page setup width
// Developer Intent: Generate a PDF where each worksheet is rendered on a single page and all columns are compressed to fit the page width.
// Use Cases: Printing wide reports without horizontal scrolling | Creating compact PDFs for catalogs or dashboards | Archiving spreadsheets so each sheet occupies one page | Producing invoices or purchase orders with many line‑item columns | Distributing spreadsheet data to stakeholders with limited screen space
// AI Prompts: Show C# code that uses Aspose.Cells to export a worksheet to a one‑page‑per‑sheet PDF with all columns forced onto the page. | Write an example that sets worksheet.PageSetup.FitToPagesWide = 1, FitToPagesTall = 0 and enables PdfSaveOptions.OnePagePerSheet and AllColumnsInOnePagePerSheet. | Explain how to combine page‑setup settings and PdfSaveOptions to produce a single‑page PDF for a wide worksheet in Aspose.Cells.

using System;
using Aspose.Cells;

// This C# example creates a workbook, adds many columns, sets the worksheet page setup to fit the width to one page (FitToPagesWide = 1, FitToPagesTall = 0), enables PdfSaveOptions.OnePagePerSheet and AllColumnsInOnePagePerSheet, and saves the file as a PDF where each sheet appears on a single page containing all columns.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate the sheet with sample data (optional, demonstrates many columns)
        for (int col = 0; col < 50; col++)
        {
            worksheet.Cells[0, col].PutValue("Column " + (col + 1));
            worksheet.Cells[1, col].PutValue("Sample data " + (col + 1));
        }

        // Configure page setup to fit all columns on one page (height adjusts automatically)
        worksheet.PageSetup.FitToPagesWide = 1;   // one page wide
        worksheet.PageSetup.FitToPagesTall = 0;   // unlimited height

        // Create PDF save options and enable the required properties
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.OnePagePerSheet = true;                // all content on one page per sheet
        pdfOptions.AllColumnsInOnePagePerSheet = true;    // force all columns onto that page

        // Save the workbook as PDF using the configured options
        workbook.Save("output.pdf", pdfOptions);
    }
}
