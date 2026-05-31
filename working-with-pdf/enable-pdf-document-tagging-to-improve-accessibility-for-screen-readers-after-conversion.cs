using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Rendering.PdfSecurity;

class EnablePdfTagging
{
    static void Main()
    {
        // Load the source Excel workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Create PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Enable export of document structure (PDF tagging) for accessibility
        pdfOptions.ExportDocumentStructure = true;

        // Configure security options to allow accessibility extraction
        PdfSecurityOptions security = new PdfSecurityOptions
        {
            // No passwords are set; only the accessibility flag is needed
            AccessibilityExtractContent = true,
            // Prevent other content extraction if desired
            ExtractContentPermission = false
        };
        pdfOptions.SecurityOptions = security;

        // Save the workbook as a tagged PDF
        workbook.Save("output_tagged.pdf", pdfOptions);
    }
}