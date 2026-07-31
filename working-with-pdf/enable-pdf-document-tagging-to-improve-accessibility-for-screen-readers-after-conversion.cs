// Title: Create an Accessible Tagged PDF from Excel with Aspose.Cells for .NET
// Description: This example builds a workbook, populates it with sample data, and saves it as a PDF with document‑structure tags (ExportDocumentStructure = true). It also configures PdfSecurityOptions so screen readers can extract content while other extraction permissions are disabled, producing a PDF that meets accessibility standards.
// Keywords: Aspose.Cells PDF tagging | ExportDocumentStructure | PDF accessibility .NET | screen reader friendly PDF | PdfSecurityOptions AccessibilityExtractContent | tagged PDF generation | accessible PDF from Excel
// Common Searches: how to enable PDF tags with Aspose.Cells | Aspose.Cells ExportDocumentStructure true example | set AccessibilityExtractContent in Aspose.Cells PDF | create accessible PDF from workbook .NET | Aspose.Cells PDF security options for screen readers
// Developer Intent: Add PDF document tags so the generated file is accessible to screen‑reading software.
// Use Cases: Produce compliance‑ready PDFs for reports that require WCAG‑AA accessibility. | Export Excel tables to PDFs that preserve logical structure for assistive technologies. | Distribute PDFs that allow text extraction for screen readers while restricting other content copying.
// AI Prompts: Show how to add custom PDF tags for table headings using Aspose.Cells. | Demonstrate setting additional PdfSecurityOptions (e.g., printing or editing permissions) while keeping document tagging enabled. | Explain how to validate the presence of accessibility tags in the output PDF with common PDF accessibility checkers.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Rendering.PdfSecurity;

namespace AsposeCellsPdfTaggingExample
{
    // This example builds a workbook, populates it with sample data, and saves it as a PDF with document‑structure tags (ExportDocumentStructure = true). It also configures PdfSecurityOptions so screen readers can extract content while other extraction permissions are disabled, producing a PDF that meets accessibility standards.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add some sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Employee");
            sheet.Cells["B1"].PutValue("Salary");
            sheet.Cells["A2"].PutValue("John");
            sheet.Cells["B2"].PutValue(50000);
            sheet.Cells["A3"].PutValue("Alice");
            sheet.Cells["B3"].PutValue(62000);

            // Configure PDF save options to enable document tagging (structure)
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // ExportDocumentStructure adds PDF tags for accessibility
                ExportDocumentStructure = true
            };

            // Configure security options to allow accessibility extraction
            PdfSecurityOptions security = new PdfSecurityOptions
            {
                // Enables screen readers to extract text/graphics
                AccessibilityExtractContent = true,
                // Optional: prevent other content extraction if desired
                ExtractContentPermission = false
            };
            pdfOptions.SecurityOptions = security;

            // Save the workbook as a tagged PDF
            workbook.Save("TaggedOutput.pdf", pdfOptions);
        }
    }
}
