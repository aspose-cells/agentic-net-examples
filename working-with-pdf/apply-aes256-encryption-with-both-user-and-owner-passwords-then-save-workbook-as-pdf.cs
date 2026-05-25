using System;
using Aspose.Cells;
using Aspose.Cells.Rendering.PdfSecurity;

class Program
{
    static void Main()
    {
        // Create a new workbook and add some data
        Workbook workbook = new Workbook();
        workbook.Worksheets[0].Cells["A1"].Value = "AES‑256 Encrypted PDF";

        // Create PDF save options
        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();

        // Configure PDF security (AES‑256 is the default encryption algorithm)
        PdfSecurityOptions securityOptions = new PdfSecurityOptions
        {
            OwnerPassword = "OwnerSecret123!",
            UserPassword = "UserSecret123!",
            PrintPermission = true,
            ModifyDocumentPermission = false,
            ExtractContentPermission = false
        };

        // Assign security options to the PDF save options
        pdfSaveOptions.SecurityOptions = securityOptions;

        // Save the workbook as a PDF with the specified encryption
        workbook.Save("EncryptedOutput.pdf", pdfSaveOptions);
    }
}