// Title: Encrypt a Workbook to PDF with a User Password using Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, apply PdfSecurityOptions (user password, optional owner password, printing permission) via PdfSaveOptions, and save the file as a password‑protected PDF with Aspose.Cells for .NET.
// Keywords: Aspose.Cells PDF encryption | C# password protected PDF | PdfSecurityOptions example | Aspose.Cells save as encrypted PDF | user password PDF Aspose | owner password PDF Aspose | protect PDF with Aspose.Cells
// Common Searches: Aspose.Cells set user password when saving PDF C# | How to encrypt PDF output from Aspose.Cells | PdfSecurityOptions C# Aspose.Cells tutorial | Save workbook as protected PDF using Aspose.Cells
// Developer Intent: Create a PDF from a workbook and secure it with a user password (optional owner password) using Aspose.Cells for .NET.
// Use Cases: Distribute confidential reports as password‑protected PDFs to clients. | Automate generation of secure invoices that require a password to open. | Provide printable PDFs that allow printing but restrict unauthorized access.
// AI Prompts: Show me a C# snippet that adds an owner password and disables editing when exporting a workbook to PDF with Aspose.Cells. | Give an example of encrypting a PDF with a user password and preventing content copying using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering.PdfSecurity;

// Demonstrates how to create a workbook, apply PdfSecurityOptions (user password, optional owner password, printing permission) via PdfSaveOptions, and save the file as a password‑protected PDF with Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook and add some data
        Workbook workbook = new Workbook();
        workbook.Worksheets[0].Cells["A1"].PutValue("Encrypted PDF Demo");

        // Create PDF save options
        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();

        // Configure PDF security options
        PdfSecurityOptions securityOptions = new PdfSecurityOptions();
        securityOptions.UserPassword = "user123";   // Password required to open the PDF
        securityOptions.OwnerPassword = "owner123"; // Owner password (optional)
        securityOptions.PrintPermission = true;    // Allow printing

        // Assign the security options to the PDF save options
        pdfSaveOptions.SecurityOptions = securityOptions;

        // Save the workbook as an encrypted PDF
        workbook.Save("EncryptedOutput.pdf", pdfSaveOptions);
    }
}
