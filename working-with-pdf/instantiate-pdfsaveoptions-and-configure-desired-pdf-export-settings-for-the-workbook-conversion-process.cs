// Title: Configure PdfSaveOptions for PDF/A‑1b Export with Compression, Optimization & Security (C# Aspose.Cells)
// Description: Demonstrates how to create a Workbook, populate cells, instantiate PdfSaveOptions, and set properties such as PDF/A‑1b compliance, minimum‑size optimization, Flate compression, document structure preservation, title display, font fallback, font compatibility checks, error handling, page range, and blank‑page behavior. Also shows how to apply PdfSecurityOptions (owner/user passwords, printing permissions) before saving the workbook as a PDF.
// Keywords: Aspose.Cells PdfSaveOptions | PDF/A-1b C# | Flate compression Aspose.Cells | PDF optimization minimum size | Aspose.Cells PDF security | font fallback Aspose.Cells | export Excel to PDF C# | Aspose.Cells PDF/A compliance | Aspose.Cells page range PDF
// Common Searches: Aspose.Cells set PDF/A-1b compliance | How to use Flate compression with PdfSaveOptions | Configure PDF security passwords in Aspose.Cells | Prevent blank pages when converting Excel to PDF | Enable document structure and title in PDF export
// Developer Intent: Configure PdfSaveOptions to control PDF compliance, compression, optimization, font handling, page output, and security before saving an Aspose.Cells workbook as a PDF.
// Use Cases: Create an archival‑ready PDF/A‑1b report with the smallest possible file size. | Apply owner and user passwords with printing permissions to protect confidential data. | Export a multi‑sheet workbook while preserving the workbook’s default font and avoiding unnecessary blank pages.
// AI Prompts: Generate C# code that sets PdfSaveOptions for PDF/A‑2b compliance with JPEG compression and custom margins using Aspose.Cells. | Explain how to enable per‑character font substitution and verify font compatibility during PDF export in Aspose.Cells. | Show how to change the owner and user passwords of a PDF created by Aspose.Cells after the file has been saved.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Rendering.PdfSecurity;

namespace AsposeCellsPdfExport
{
    // Demonstrates how to create a Workbook, populate cells, instantiate PdfSaveOptions, and set properties such as PDF/A‑1b compliance, minimum‑size optimization, Flate compression, document structure preservation, title display, font fallback, font compatibility checks, error handling, page range, and blank‑page behavior. Also shows how to apply PdfSecurityOptions (owner/user passwords, printing permissions) before saving the workbook as a PDF.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add some sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Aspose.Cells PDF Export Example");
            sheet.Cells["A2"].PutValue(DateTime.Now);
            sheet.Cells["B1"].PutValue(12345);
            sheet.Cells["B2"].PutValue(67890);

            // Instantiate PdfSaveOptions
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Configure desired PDF export settings
            pdfOptions.Compliance = PdfCompliance.PdfA1b;                     // PDF/A-1b compliance
            pdfOptions.OptimizationType = PdfOptimizationType.MinimumSize;   // Optimize for minimum file size
            pdfOptions.PdfCompression = PdfCompressionCore.Flate;            // Use Flate compression
            pdfOptions.ExportDocumentStructure = true;                       // Preserve document structure
            pdfOptions.DisplayDocTitle = true;                               // Show document title in viewer
            pdfOptions.DefaultFont = "Arial";                                // Fallback font
            pdfOptions.OnePagePerSheet = false;                              // Allow multiple pages per sheet
            pdfOptions.AllColumnsInOnePagePerSheet = false;                 // Do not force all columns onto one page
            pdfOptions.CheckWorkbookDefaultFont = true;                      // Use workbook default font when needed
            pdfOptions.CheckFontCompatibility = true;                       // Verify font compatibility
            pdfOptions.IsFontSubstitutionCharGranularity = true;            // Substitute fonts per character if needed
            pdfOptions.IgnoreError = true;                                   // Hide rendering errors
            pdfOptions.OutputBlankPageWhenNothingToPrint = false;            // Do not output blank pages
            pdfOptions.PageIndex = 0;                                        // Start from first page
            pdfOptions.PageCount = 0;                                        // 0 means all pages
            pdfOptions.PrintingPageType = PrintingPageType.IgnoreBlank;      // Ignore blank pages when printing

            // Optional: set security options (e.g., password protection)
            PdfSecurityOptions security = new PdfSecurityOptions
            {
                OwnerPassword = "ownerPwd",
                UserPassword = "userPwd",
                PrintPermission = true,
                FullQualityPrintPermission = true
            };
            pdfOptions.SecurityOptions = security;

            // Save the workbook as PDF using the configured options
            workbook.Save("ExportedDocument.pdf", pdfOptions);
        }
    }
}
