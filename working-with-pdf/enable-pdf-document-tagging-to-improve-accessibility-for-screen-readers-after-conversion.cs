// Title: Create a Tagged PDF with Accessibility for Screen Readers using Aspose.Cells .NET (C#)
// Description: C# example that generates a PDF from an Excel workbook with PDF tags enabled (ExportDocumentStructure), sets the document title, and configures security to allow accessibility extraction while restricting other content copying.
// Keywords: Aspose.Cells | C# | PDF tagging | ExportDocumentStructure | PDF accessibility | screen reader | PdfSaveOptions | PdfSecurityOptions | PDF/UA | document title
// Common Searches: Aspose.Cells enable PDF tags | How to create accessible PDF from Excel C# | ExportDocumentStructure Aspose.Cells example | Allow screen readers in PDF saved with Aspose.Cells | Set PDF document title using Aspose.Cells | Restrict PDF content extraction while keeping accessibility
// Developer Intent: Generate a PDF with structural tags and accessibility settings so screen readers can read the content.
// Use Cases: Produce PDF/UA‑compliant documents from Excel for compliance audits | Distribute accessible reports to visually impaired users | Create PDFs with a visible title bar while preventing unauthorized content copying | Automate generation of tagged PDFs in enterprise reporting pipelines
// AI Prompts: Give a step‑by‑step guide to add custom PDF tags for tables in Aspose.Cells. | Show how to combine PDF/A‑2u compliance with ExportDocumentStructure in Aspose.Cells. | Demonstrate how to validate PDF accessibility tags using open‑source tools after saving with Aspose.Cells. | Explain how to set different security permissions for accessibility versus editing in Aspose.Cells PDF export.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Rendering.PdfSecurity;

// C# example that generates a PDF from an Excel workbook with PDF tags enabled (ExportDocumentStructure), sets the document title, and configures security to allow accessibility extraction while restricting other content copying.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Add some sample data to demonstrate accessibility tagging
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Title");
        sheet.Cells["A2"].PutValue("This text will be accessible to screen readers after PDF conversion.");

        // Configure PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Enable export of document structure (PDF tags) for accessibility
        pdfOptions.ExportDocumentStructure = true;

        // Optional: display the document title in the PDF viewer's title bar
        pdfOptions.DisplayDocTitle = true;

        // Configure security options to allow accessibility extraction
        PdfSecurityOptions security = new PdfSecurityOptions
        {
            // Allow screen readers to extract text and graphics
            AccessibilityExtractContent = true,
            // Prevent other content extraction if desired
            ExtractContentPermission = false
        };

        // Assign the security options to the PDF save options
        pdfOptions.SecurityOptions = security;

        // Save the workbook as a tagged PDF
        workbook.Save("TaggedDocument.pdf", pdfOptions);
    }
}
