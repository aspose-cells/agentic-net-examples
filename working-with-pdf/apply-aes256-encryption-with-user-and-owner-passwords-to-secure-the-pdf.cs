// Title: Encrypt PDF with AES‑256 using owner & user passwords in Aspose.Cells for .NET (C#)
// Description: Creates a workbook, adds text, and saves it as a password‑protected PDF. Setting PdfSecurityOptions with owner and user passwords triggers AES‑256 encryption and lets you define permissions such as printing and editing.
// Keywords: Aspose.Cells | C# | .NET | PDF encryption | AES-256 | owner password | user password | PdfSecurityOptions | secure PDF | PDF permissions
// Common Searches: AES‑256 PDF encryption Aspose.Cells C# | set owner and user passwords PDF Aspose.Cells .NET | restrict printing in PDF saved from workbook | Aspose.Cells PDF security options example | C# encrypt PDF with password using Aspose.Cells
// Developer Intent: Add AES‑256 encryption and password protection to a PDF generated from a workbook.
// Use Cases: Protect confidential financial reports before distribution. | Share read‑only PDFs with external stakeholders while allowing printing. | Create PDFs that require a password to open, with the owner retaining full rights.
// AI Prompts: Generate code to set copy and annotation permissions in PdfSecurityOptions. | Show how to use a custom encryption level (e.g., 128‑bit) with Aspose.Cells. | Explain how to validate the owner password of an encrypted PDF programmatically.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering.PdfSecurity;

// Creates a workbook, adds text, and saves it as a password‑protected PDF. Setting PdfSecurityOptions with owner and user passwords triggers AES‑256 encryption and lets you define permissions such as printing and editing.
class SecurePdfExample
{
    static void Main()
    {
        // Create a new workbook and add some content
        Workbook workbook = new Workbook();
        workbook.Worksheets[0].Cells["A1"].Value = "Secure PDF with AES‑256 encryption";

        // Configure PDF save options
        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();

        // Set up PDF security options with owner and user passwords
        PdfSecurityOptions securityOptions = new PdfSecurityOptions
        {
            OwnerPassword = "OwnerPass123!",
            UserPassword = "UserPass123!",
            // Example permissions (adjust as needed)
            PrintPermission = true,
            ModifyDocumentPermission = false
        };

        // Attach the security options to the PDF save options
        pdfSaveOptions.SecurityOptions = securityOptions;

        // Save the workbook as a PDF; Aspose.Cells applies AES‑256 encryption by default when passwords are set
        workbook.Save("SecureDocument.pdf", pdfSaveOptions);
    }
}
