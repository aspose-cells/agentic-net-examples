// Title: C# – Convert HTML to Password‑Protected PDF using Aspose.Cells (User & Owner Passwords)
// Description: Load an HTML file into an Aspose.Cells Workbook, configure PdfSaveOptions with PdfSecurityOptions (user password, owner password, permissions), and save the workbook as an encrypted PDF.
// Keywords: Aspose.Cells | HTML to PDF C# | PDF encryption Aspose.Cells | PdfSecurityOptions | user password PDF | owner password PDF | set PDF permissions | Aspose.Cells .NET | encrypt PDF C# | secure PDF generation
// Common Searches: Aspose.Cells convert html to pdf c# | password protect pdf using Aspose.Cells | set user password for pdf Aspose.Cells | pdf security options c# Aspose | encrypt pdf from html workbook
// Developer Intent: Create a PDF from an HTML file and protect it with a user password (optionally an owner password) using Aspose.Cells for .NET.
// Use Cases: Generate a confidential report from an HTML template and distribute it as a password‑protected PDF. | Produce an invoice PDF from an HTML invoice that requires a user password to open while still allowing printing. | Create internal documentation PDFs with restricted permissions and a user password for secure access.
// AI Prompts: Show how to add copy and modify permissions to PdfSecurityOptions in Aspose.Cells. | Provide C# code that encrypts a PDF with only a user password and disables all other permissions using Aspose.Cells. | Explain how to load HTML from a memory stream instead of a file and save an encrypted PDF with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering.PdfSecurity;

// Load an HTML file into an Aspose.Cells Workbook, configure PdfSaveOptions with PdfSecurityOptions (user password, owner password, permissions), and save the workbook as an encrypted PDF.
class HtmlToPdfEncrypted
{
    static void Main()
    {
        // Load the HTML file into a workbook
        Workbook workbook = new Workbook("input.html");

        // Create PDF save options
        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();

        // Configure PDF security options
        PdfSecurityOptions securityOptions = new PdfSecurityOptions
        {
            // Password required for opening the PDF (user password)
            UserPassword = "UserPassword123",
            // Owner password provides full access without restrictions
            OwnerPassword = "OwnerPassword123",
            // Example permission: allow printing
            PrintPermission = true
        };

        // Assign the security options to the PDF save options
        pdfSaveOptions.SecurityOptions = securityOptions;

        // Save the workbook as an encrypted PDF
        workbook.Save("EncryptedOutput.pdf", pdfSaveOptions);
    }
}
