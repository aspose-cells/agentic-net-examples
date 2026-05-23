using System;
using Aspose.Cells;
using Aspose.Cells.Rendering.PdfSecurity;

class Program
{
    static void Main()
    {
        // Create a new workbook and add some sample data
        Workbook workbook = new Workbook();
        workbook.Worksheets[0].Cells["A1"].Value = "Encrypted PDF Demo";

        // Prepare PDF save options
        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();

        // Configure PDF security options
        PdfSecurityOptions securityOptions = new PdfSecurityOptions();
        securityOptions.UserPassword = "user123";   // password required to open the PDF
        securityOptions.OwnerPassword = "owner123"; // optional owner password with full rights
        securityOptions.PrintPermission = true;    // allow printing (adjust as needed)

        // Attach the security options to the PDF save options
        pdfSaveOptions.SecurityOptions = securityOptions;

        // Save the workbook as an encrypted PDF
        workbook.Save("EncryptedDocument.pdf", pdfSaveOptions);
    }
}