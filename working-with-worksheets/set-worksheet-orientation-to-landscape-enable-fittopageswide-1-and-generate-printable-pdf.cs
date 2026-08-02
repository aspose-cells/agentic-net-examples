// Title: C# – Set Landscape Orientation, Fit Worksheet to One Page Wide, and Export to PDF with Aspose.Cells
// Description: Create a workbook, set the first worksheet to landscape, configure FitToPagesWide = 1 (auto height), apply PdfSaveOptions, and save a printable single‑page PDF using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# | landscape orientation | FitToPagesWide | FitToPagesTall | PDF export | single page PDF | PageSetup | PdfSaveOptions | .NET worksheet printing | global developers
// Common Searches: Aspose.Cells set worksheet to landscape and fit to one page wide | export worksheet as single‑page PDF C# Aspose.Cells | FitToPagesWide 1 PDF output Aspose.Cells .NET | how to force one page per sheet in Aspose.Cells PDF | C# code for printable PDF with landscape orientation
// Developer Intent: Configure a worksheet for landscape printing, fit it horizontally to one page, and generate a PDF.
// Use Cases: Produce printable reports that must appear in landscape on a single PDF page. | Create invoices or data sheets that automatically adjust to one‑page width for any printer. | Generate compact dashboards for distribution where horizontal layout is critical.
// AI Prompts: Write C# code with Aspose.Cells to set landscape orientation, FitToPagesWide = 1, and export the sheet to a PDF. | Explain the effect of FitToPagesTall = 0 when exporting a worksheet to PDF using Aspose.Cells. | Show how to use PdfSaveOptions.OnePagePerSheet to force a single‑page PDF per worksheet.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Create a workbook, set the first worksheet to landscape, configure FitToPagesWide = 1 (auto height), apply PdfSaveOptions, and save a printable single‑page PDF using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // (Optional) Add some data to demonstrate the layout
        worksheet.Cells["A1"].PutValue("Sample Data");
        worksheet.Cells["B1"].PutValue(123);
        worksheet.Cells["A2"].PutValue("More Data");
        worksheet.Cells["B2"].PutValue(456);

        // Set page orientation to Landscape
        worksheet.PageSetup.Orientation = PageOrientationType.Landscape;

        // Fit the worksheet to one page wide; height will adjust automatically
        worksheet.PageSetup.FitToPagesWide = 1;
        worksheet.PageSetup.FitToPagesTall = 0; // 0 means auto

        // Prepare PDF save options (optional settings)
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.OnePagePerSheet = true; // ensure the whole sheet is on a single PDF page

        // Save the workbook as a printable PDF
        workbook.Save("output.pdf", pdfOptions);
    }
}
