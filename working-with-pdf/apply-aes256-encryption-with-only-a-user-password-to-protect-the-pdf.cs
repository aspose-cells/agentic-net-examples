using System;
using Aspose.Cells;
using Aspose.Cells.Rendering.PdfSecurity;

class Program
{
    static void Main()
    {
        // Create a new workbook and add some data
        Workbook workbook = new Workbook();
        workbook.Worksheets[0].Cells["A1"].Value = "PDF protected with user password only";

        // Prepare PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Configure PDF security: set only the user password (owner password left null)
        PdfSecurityOptions security = new PdfSecurityOptions();
        security.UserPassword = "user123"; // password required to open the PDF
        // OwnerPassword is not set, so only user password is used
        pdfOptions.SecurityOptions = security;

        // Save the workbook as a PDF with the specified security settings
        workbook.Save("ProtectedUserPassword.pdf", pdfOptions);
    }
}