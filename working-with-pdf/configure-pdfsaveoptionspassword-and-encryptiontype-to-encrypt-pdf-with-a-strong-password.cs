using System;
using Aspose.Cells;
using Aspose.Cells.Rendering.PdfSecurity;

class Program
{
    static void Main()
    {
        // Create a new workbook and add some content
        Workbook workbook = new Workbook();
        workbook.Worksheets[0].Cells["A1"].Value = "Secure PDF Demo";

        // Set a strong password for the workbook (Excel encryption)
        workbook.Settings.Password = "StrongWorkbookPwd!";

        // Apply strong encryption options to the workbook
        // Using EnhancedCryptographicProviderV1 with a 256‑bit key for maximum security
        workbook.SetEncryptionOptions(EncryptionType.EnhancedCryptographicProviderV1, 256);

        // Configure PDF security options
        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();
        PdfSecurityOptions pdfSecurity = new PdfSecurityOptions
        {
            OwnerPassword = "OwnerStrongPwd!",
            UserPassword = "UserStrongPwd!",
            PrintPermission = true,               // Allow printing
            ModifyDocumentPermission = false,    // Disallow modifications
            ExtractContentPermission = false     // Disallow content extraction
        };
        pdfSaveOptions.SecurityOptions = pdfSecurity;

        // Save the workbook as a password‑protected PDF
        workbook.Save("SecureDocument.pdf", pdfSaveOptions);
    }
}
// Author: Aspose.Cells .NET example – demonstrates PDF encryption with strong passwords and encryption type.