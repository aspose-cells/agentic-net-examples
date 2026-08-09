// Title: C# – Convert Excel to a Secure Non‑Editable PDF with Aspose.Cells (remove form fields)
// Description: Load an Excel workbook, strip personal information, and save it as a PDF using Aspose.Cells. PdfSaveOptions together with PdfSecurityOptions apply owner/user passwords, allow printing, and block editing, form filling, annotations, and content extraction, delivering a tamper‑proof, read‑only PDF.
// Keywords: Aspose.Cells PDF conversion .NET | secure PDF from Excel C# | disable form fields Aspose.Cells | remove personal information PDF | PdfSecurityOptions example | non editable PDF Aspose | Excel to PDF with permissions
// Common Searches: how to create a read‑only PDF from Excel using Aspose.Cells | Aspose.Cells C# disable editing and form filling in PDF | remove personal data and protect PDF generated from workbook | set owner and user passwords for PDF in Aspose.Cells | export Excel to secured PDF .NET
// Developer Intent: Produce a PDF from an Excel file that is protected against editing, form filling, annotations, and content extraction.
// Use Cases: Distribute client‑ready reports that must remain unchanged. | Archive regulatory spreadsheets as tamper‑proof PDFs for compliance audits. | Convert filled Excel forms into locked PDFs to prevent further modifications.
// AI Prompts: Generate C# code that uses Aspose.Cells to convert an Excel workbook to PDF, removes personal information, and applies PdfSecurityOptions to block editing, form filling, annotations, and extraction while allowing printing. | Show an example of protecting a PDF with owner and user passwords in Aspose.Cells, ensuring the document is non‑editable and form fields are disabled.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering.PdfSecurity;

// Load an Excel workbook, strip personal information, and save it as a PDF using Aspose.Cells. PdfSaveOptions together with PdfSecurityOptions apply owner/user passwords, allow printing, and block editing, form filling, annotations, and content extraction, delivering a tamper‑proof, read‑only PDF.
class ConvertToPdfNonEditable
{
    static void Main()
    {
        // Load the source Excel workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Remove any personal information (author names, comments, etc.)
        workbook.RemovePersonalInformation();

        // Create PDF save options
        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();

        // Configure security options to make the PDF non‑editable
        PdfSecurityOptions securityOptions = new PdfSecurityOptions
        {
            OwnerPassword = "owner123",          // Owner password (can be any value)
            UserPassword = "user123",            // User password (can be any value)
            PrintPermission = true,              // Allow printing
            ModifyDocumentPermission = false,    // Disallow document modifications
            FillFormsPermission = false,         // Disallow filling form fields
            AnnotationsPermission = false,       // Disallow adding/modifying annotations
            ExtractContentPermission = false     // Disallow content extraction
        };

        // Assign the security options to the PDF save options
        pdfSaveOptions.SecurityOptions = securityOptions;

        // Save the workbook as a PDF with the specified options
        workbook.Save("output.pdf", pdfSaveOptions);
    }
}
