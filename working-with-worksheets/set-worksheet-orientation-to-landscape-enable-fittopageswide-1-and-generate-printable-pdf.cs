// Title: Aspose.Cells .NET – Landscape worksheet, Fit‑to‑Width = 1, Export Printable PDF
// Description: Learn how to set a worksheet’s page orientation to landscape, configure PageSetup.FitToPagesWide = 1 (height auto‑adjusts), apply PdfSaveOptions.OnePagePerSheet, and save the workbook as a print‑ready PDF using Aspose.Cells for .NET.
// Keywords: Aspose.Cells landscape orientation | FitToPagesWide 1 | export worksheet to PDF | PageSetup FitToPagesWide example | PdfSaveOptions OnePagePerSheet | .NET spreadsheet PDF export
// Common Searches: Aspose.Cells set worksheet to landscape | Fit worksheet width to one page Aspose .NET | How to export PDF with one page per sheet Aspose.Cells | PageSetup FitToPagesWide = 1 example | Printable PDF from Excel using Aspose.Cells
// Developer Intent: Configure a worksheet for landscape printing, fit its width to a single page, and generate a printable PDF.
// Use Cases: Create landscape reports that print on a single page per sheet | Export financial statements with consistent one‑page width formatting | Generate printable invoices or receipts where each sheet occupies one PDF page
// AI Prompts: Show code to set PageSetup.Orientation to Landscape and FitToPagesWide to 1 before saving as PDF with Aspose.Cells. | Provide an Aspose.Cells .NET example that exports a workbook to PDF with OnePagePerSheet enabled and automatic height scaling. | Explain how to adjust FitToPagesTall while keeping FitToPagesWide = 1 for PDF output using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Learn how to set a worksheet’s page orientation to landscape, configure PageSetup.FitToPagesWide = 1 (height auto‑adjusts), apply PdfSaveOptions.OnePagePerSheet, and save the workbook as a print‑ready PDF using Aspose.Cells for .NET.
class GeneratePdf
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // (Optional) Add some sample data
        sheet.Cells["A1"].PutValue("Sample Data");
        sheet.Cells["B1"].PutValue(123);
        sheet.Cells["A2"].PutValue("More Data");
        sheet.Cells["B2"].PutValue(456);

        // Configure page setup
        PageSetup setup = sheet.PageSetup;
        setup.Orientation = PageOrientationType.Landscape; // Landscape orientation
        setup.FitToPagesWide = 1;   // Fit to one page wide
        setup.FitToPagesTall = 0;   // Height adjusts automatically

        // Set PDF save options (optional: force one page per sheet)
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.OnePagePerSheet = true;

        // Save the workbook as a printable PDF
        workbook.Save("output.pdf", pdfOptions);
    }
}
