using System;
using Aspose.Cells;
using Aspose.Cells.Rendering.PdfSecurity;

// Author: Aspose.Cells .NET example – AES‑256 workbook encryption + PDF user/owner passwords
class Program
{
    static void Main()
    {
        // Create a new workbook and add some sample data
        Workbook workbook = new Workbook();
        workbook.Worksheets[0].Cells["A1"].Value = "Secure PDF Demo";

        // ---------- Workbook encryption (AES‑256) ----------
        // Set encryption type to Enhanced Cryptographic Provider (supports 256‑bit keys)
        workbook.SetEncryptionOptions(EncryptionType.EnhancedCryptographicProviderV1, 256);
        // Protect the workbook with an owner password (required to activate encryption)
        workbook.Protect(ProtectionType.All, "ownerPassword");

        // ---------- PDF security options ----------
        // Create PDF save options and configure security (owner & user passwords)
        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();
        PdfSecurityOptions pdfSecurity = new PdfSecurityOptions
        {
            OwnerPassword = "ownerPassword",   // Owner password – full permissions
            UserPassword = "userPassword",     // User password – limited permissions
            PrintPermission = true,            // Allow printing
            FullQualityPrintPermission = true // Allow high‑quality printing
        };
        pdfSaveOptions.SecurityOptions = pdfSecurity;

        // Save the workbook as a PDF with the defined security settings
        workbook.Save("EncryptedOutput.pdf", pdfSaveOptions);
    }
}