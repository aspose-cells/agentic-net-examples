using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Rendering.PdfSecurity;

class PdfA3uAccessibilityExample
{
    static void Main()
    {
        // Create a new workbook and add sample data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sample content for PDF/A‑3u with accessibility metadata");

        // Configure PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Set compliance level to PDF/A‑3u
            Compliance = PdfCompliance.PdfA3u,
            // Export document structure to improve accessibility (tags, outline, etc.)
            ExportDocumentStructure = true
        };

        // Enable accessibility features via security options
        PdfSecurityOptions securityOptions = new PdfSecurityOptions
        {
            // Allow extraction of text/graphics for assistive technologies
            AccessibilityExtractContent = true
        };
        pdfOptions.SecurityOptions = securityOptions;

        // Save the workbook as a PDF with the specified options
        workbook.Save("AccessiblePdfA3u.pdf", pdfOptions);
    }
}