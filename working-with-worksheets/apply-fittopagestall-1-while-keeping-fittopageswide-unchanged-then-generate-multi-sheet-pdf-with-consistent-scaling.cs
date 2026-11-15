// Title: Configure FitToPagesTall = 1 while preserving FitToPagesWide and export a multi‑sheet PDF with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set page setup FitToPagesTall to 1 without changing FitToPagesWide | Export all worksheets to a single PDF while keeping vertical scaling consistent using Aspose.Cells | How to retain per‑sheet page scaling when saving a workbook as PDF with Aspose.Cells .NET | C# generate multi‑sheet PDF with FitToPagesTall = 1 and OnePagePerSheet false in Aspose.Cells | Aspose.Cells PdfSaveOptions keep worksheet scaling across pages
// Tags: FitToPagesTall scaling Aspose.Cells C# | preserve FitToPagesWide Aspose.Cells | multi-sheet PDF export Aspose.Cells | PdfSaveOptions OnePagePerSheet false | per-sheet page scaling PDF Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Rendering;

// Load an existing workbook (replace with your source file)
Workbook workbook = new Workbook("input.xlsx");

// Iterate through all worksheets to set page scaling
foreach (Worksheet sheet in workbook.Worksheets)
{
    PageSetup pageSetup = sheet.PageSetup;

    // Keep the current FitToPagesWide value unchanged.
    // (If it was 0, it remains automatic; otherwise the existing value is preserved.)

    // Set FitToPagesTall to 1 to force each sheet to fit vertically on a single page.
    pageSetup.FitToPagesTall = 1;
}

// Configure PDF save options to preserve the per‑sheet scaling set above.
PdfSaveOptions pdfOptions = new PdfSaveOptions
{
    // Do not force each sheet onto a single PDF page; keep the scaling defined by FitToPagesTall.
    OnePagePerSheet = false,

    // Do not combine all columns onto one page; retain the original layout.
    AllColumnsInOnePagePerSheet = false
};

// Save the workbook as a multi‑sheet PDF with consistent scaling across sheets.
workbook.Save("output.pdf", pdfOptions);
