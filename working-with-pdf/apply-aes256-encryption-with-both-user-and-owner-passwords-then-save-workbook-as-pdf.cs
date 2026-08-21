// Title: AES‑256 PDF encryption with user & owner passwords using Aspose.Cells for .NET (C#)
// Description: The sample creates a Workbook, writes a value to cell A1, configures PdfSaveOptions with PdfSecurityOptions (owner password, user password, print permission), and saves the file as an AES‑256 encrypted PDF called EncryptedOutput.pdf.
// Keywords: Aspose.Cells PDF encryption C# | AES-256 PDF Aspose | PdfSecurityOptions example | owner password PDF .NET | user password Aspose.Cells | protect PDF with Aspose
// Common Searches: How to encrypt PDF with AES‑256 in Aspose.Cells C# | Set owner and user passwords for PDF output using Aspose.Cells | Aspose.Cells PDF security options tutorial
// Developer Intent: Generate a PDF from a workbook and secure it with AES‑256 encryption, requiring both a user password to open and an owner password to control permissions.
// Use Cases: Deliver confidential reports that require a password to view while allowing printing. | Send secure invoices where only authorized recipients can open the document. | Create internal documentation with restricted access and specific permissions such as printing only.
// AI Prompts: Show how to disable text copying in the encrypted PDF. | Add editing and annotation permissions while keeping AES‑256 protection. | Explain how to switch the encryption level to 128‑bit in PdfSecurityOptions.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering.PdfSecurity;

// The sample creates a Workbook, writes a value to cell A1, configures PdfSaveOptions with PdfSecurityOptions (owner password, user password, print permission), and saves the file as an AES‑256 encrypted PDF called EncryptedOutput.pdf.
class Program
{
    static void Main()
    {
        // Create a new workbook and add some data
        Workbook workbook = new Workbook();
        workbook.Worksheets[0].Cells["A1"].Value = "AES‑256 Protected PDF";

        // Create PDF save options
        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();

        // Configure PDF security options (owner and user passwords)
        PdfSecurityOptions pdfSecurityOptions = new PdfSecurityOptions
        {
            OwnerPassword = "OwnerPass123!",
            UserPassword = "UserPass123!",
            PrintPermission = true // example permission
        };

        // Assign the security options to the PDF save options
        pdfSaveOptions.SecurityOptions = pdfSecurityOptions;

        // Save the workbook as a PDF with the specified security settings
        workbook.Save("EncryptedOutput.pdf", pdfSaveOptions);
    }
}
