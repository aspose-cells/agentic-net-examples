using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class PdfExportDemo
{
    static void Main()
    {
        // Create a new workbook and add some sample data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Aspose.Cells PDF Export Demo");

        // Instantiate PdfSaveOptions (rule: PdfSaveOptions constructor)
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Configure desired PDF export settings
        pdfOptions.Compliance = PdfCompliance.PdfA1b;                 // PDF/A-1b compliance
        pdfOptions.PdfCompression = PdfCompressionCore.Flate;        // Use Flate compression
        pdfOptions.OptimizationType = PdfOptimizationType.MinimumSize; // Optimize for minimum file size
        pdfOptions.ExportDocumentStructure = true;                   // Export document structure
        pdfOptions.DisplayDocTitle = true;                           // Show document title in PDF viewer
        pdfOptions.DefaultFont = "Arial";                            // Default font for Unicode characters
        pdfOptions.OnePagePerSheet = true;                           // Render each sheet on a single page
        pdfOptions.AllColumnsInOnePagePerSheet = true;               // Fit all columns on one page per sheet
        pdfOptions.CheckWorkbookDefaultFont = true;                  // Use workbook's default font when needed
        pdfOptions.CheckFontCompatibility = true;                    // Verify font compatibility for each character
        pdfOptions.IsFontSubstitutionCharGranularity = true;         // Substitute fonts per character if needed
        pdfOptions.IgnoreError = true;                               // Hide rendering errors
        pdfOptions.OutputBlankPageWhenNothingToPrint = true;         // Output a blank page if nothing to print

        // Save the workbook as a PDF file using the configured options (rule: workbook.Save with options)
        workbook.Save("ExportedDocument.pdf", pdfOptions);
    }
}