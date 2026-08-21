// Title: C# – Convert Excel Workbook to PDF with Fit‑to‑Page Scaling using Aspose.Cells
// Description: Shows how to load or create a workbook, set the page setup to fit all columns on one page (rows auto‑adjust), configure PdfSaveOptions (OnePagePerSheet), and save the file as a PDF for efficient paper usage.
// Keywords: Aspose.Cells | C# | .NET | PDF conversion | Fit to page | SetFitToPages | OnePagePerSheet | page scaling | printable area | Excel to PDF | PdfSaveOptions
// Common Searches: Aspose.Cells fit to page PDF export C# | SetFitToPages example Aspose.Cells | OnePagePerSheet PDF option Aspose.Cells .NET | How to scale Excel to PDF width using Aspose | Configure page setup for PDF conversion Aspose.Cells
// Developer Intent: Generate a PDF from an Excel workbook while scaling the printable area to fit the page width, reducing the number of printed pages.
// Use Cases: Printing invoices where all columns must appear on a single PDF page. | Exporting multi‑sheet reports to PDF with each sheet confined to one page for concise documentation. | Batch converting Excel files to PDF with consistent fit‑to‑page settings to standardize print layouts.
// AI Prompts: Provide C# code that loads an existing Excel file, applies SetFitToPages(1,0), enables OnePagePerSheet, and saves it as a PDF with Aspose.Cells. | Explain the impact of SetFitToPages(1,0) on PDF output and how OnePagePerSheet further optimizes paper usage. | Show an example of configuring PdfSaveOptions for single‑page‑per‑sheet PDF export in Aspose.Cells.

using System;
using Aspose.Cells;

// Shows how to load or create a workbook, set the page setup to fit all columns on one page (rows auto‑adjust), configure PdfSaveOptions (OnePagePerSheet), and save the file as a PDF for efficient paper usage.
class ConvertWorkbookToPdf
{
    static void Main()
    {
        // Create a new workbook (or load an existing one with new Workbook("input.xlsx"))
        Workbook workbook = new Workbook();

        // Populate some sample data
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Item");
        sheet.Cells["B1"].PutValue("Quantity");
        sheet.Cells["A2"].PutValue("Apples");
        sheet.Cells["B2"].PutValue(50);
        sheet.Cells["A3"].PutValue("Bananas");
        sheet.Cells["B3"].PutValue(30);

        // Set page scaling to fit the printable area:
        // Fit all columns on one page (wide = 1) and let rows expand automatically (tall = 0)
        sheet.PageSetup.SetFitToPages(1, 0);
        // Alternatively:
        // sheet.PageSetup.FitToPagesWide = 1;
        // sheet.PageSetup.FitToPagesTall = 0;

        // Configure PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Ensure each worksheet is rendered on a single page (optional, improves paper usage)
            OnePagePerSheet = true
        };

        // Save the workbook as PDF using the configured options
        workbook.Save("output.pdf", pdfOptions);
    }
}
