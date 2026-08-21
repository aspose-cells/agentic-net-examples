// Title: Encrypt PDF with strong owner and user passwords using Aspose.Cells PdfSaveOptions (C#)
// Description: Creates a Workbook, applies workbook‑level encryption, configures PdfSecurityOptions with robust owner and user passwords and specific permissions (print, modify, extract), assigns the options to PdfSaveOptions.SecurityOptions, and saves the workbook as a password‑protected PDF.
// Keywords: Aspose.Cells | PdfSaveOptions | PDF encryption C# | owner password PDF | user password PDF | strong encryption | PDF permissions | .NET | Aspose.Cells PDF security | EncryptionType 256‑bit
// Common Searches: Aspose.Cells set owner password for PDF | C# PdfSaveOptions password protection | How to apply strong encryption to PDF with Aspose.Cells | Configure PDF permissions (print, modify) using Aspose.Cells | Encrypt exported workbook as PDF in .NET
// Developer Intent: Add password protection and granular security settings to a PDF generated from an Aspose.Cells workbook.
// Use Cases: Distribute confidential reports that require a password to open and restrict editing. | Provide partners with printable PDFs while preventing content extraction or modification. | Meet regulatory compliance by exporting workbooks as 256‑bit encrypted PDFs with owner/user passwords.
// AI Prompts: Show C# code to encrypt a PDF with 256‑bit owner and user passwords using Aspose.Cells PdfSaveOptions. | How can I disable document modification but allow printing in a password‑protected PDF with Aspose.Cells? | Explain the steps to combine workbook encryption and PDF security options in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering.PdfSecurity;

// Creates a Workbook, applies workbook‑level encryption, configures PdfSecurityOptions with robust owner and user passwords and specific permissions (print, modify, extract), assigns the options to PdfSaveOptions.SecurityOptions, and saves the workbook as a password‑protected PDF.
class Program
{
    static void Main()
    {
        // Create a new workbook and add some data
        Workbook workbook = new Workbook();
        workbook.Worksheets[0].Cells["A1"].Value = "Secure PDF Example";

        // Apply strong encryption to the workbook (Excel file encryption)
        workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 256);
        workbook.Settings.Password = "StrongWorkbookPassword!@#";

        // Configure PDF security options with strong passwords
        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();
        PdfSecurityOptions pdfSecurityOptions = new PdfSecurityOptions
        {
            OwnerPassword = "StrongOwnerPassword!@#",
            UserPassword = "StrongUserPassword!@#",
            PrintPermission = true,
            ModifyDocumentPermission = false,
            ExtractContentPermission = false,
            FullQualityPrintPermission = true
        };
        pdfSaveOptions.SecurityOptions = pdfSecurityOptions;

        // Save the workbook as a password‑protected PDF
        workbook.Save("SecureDocument.pdf", pdfSaveOptions);
    }
}
