// Title: C# – Convert Aspose.Cells Workbook to Encrypted PDF with User & Owner Passwords
// Description: Creates a workbook, adds data, configures PdfSaveOptions with PdfSecurityOptions (user password, optional owner password, printing permission) and saves the file as a password‑protected PDF using Aspose.Cells for .NET.
// Keywords: Aspose.Cells PDF encryption | C# encrypt PDF Aspose | PdfSecurityOptions user password | owner password Aspose.Cells | Aspose.Cells save as PDF with security | .NET PDF password protection | secure PDF generation C# | Aspose.Cells PDF conversion
// Common Searches: Aspose.Cells add password to PDF | How to encrypt PDF using Aspose.Cells C# | Set user and owner password when saving workbook to PDF | Enable printing permission in Aspose.Cells PDF | C# code sample for PDF security Aspose.Cells
// Developer Intent: Generate a PDF from a workbook and protect it with a user password (optionally an owner password) via Aspose.Cells for .NET.
// Use Cases: Distribute confidential spreadsheet reports as PDFs that require a password to open. | Provide secure, printable PDFs for regulatory compliance while restricting editing. | Automate batch conversion of workbooks to encrypted PDFs with custom permissions.
// AI Prompts: Show C# code to convert an Aspose.Cells workbook to a PDF protected by a user password and disable editing. | Explain how to set owner password and control permissions (printing, copying) when saving a PDF with Aspose.Cells. | Provide a step‑by‑step guide to modify PDF security settings after a workbook has been saved as an encrypted PDF.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering.PdfSecurity;

// Creates a workbook, adds data, configures PdfSaveOptions with PdfSecurityOptions (user password, optional owner password, printing permission) and saves the file as a password‑protected PDF using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook and add some data
        Workbook workbook = new Workbook();
        workbook.Worksheets[0].Cells["A1"].Value = "Secure PDF Content";

        // Prepare PDF save options
        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();

        // Configure PDF security options with a user password
        PdfSecurityOptions securityOptions = new PdfSecurityOptions();
        securityOptions.UserPassword = "userPassword123";   // password required to open the PDF
        securityOptions.OwnerPassword = "ownerPassword123"; // optional owner password
        securityOptions.PrintPermission = true;            // allow printing (optional)

        // Assign the security options to the PDF save options
        pdfSaveOptions.SecurityOptions = securityOptions;

        // Save the workbook as an encrypted PDF
        workbook.Save("EncryptedWorkbook.pdf", pdfSaveOptions);
    }
}
