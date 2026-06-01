using System;
using Aspose.Cells;
using Aspose.Cells.Rendering.PdfSecurity;

class Program
{
    static void Main()
    {
        // Create a new workbook and add some content
        Workbook workbook = new Workbook();
        workbook.Worksheets[0].Cells["A1"].Value = "PDF with user password and printing disabled";

        // Initialize PDF save options
        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();

        // Set up security options
        PdfSecurityOptions securityOptions = new PdfSecurityOptions
        {
            // Password required to open the PDF
            UserPassword = "user123",
            // Owner password grants full access
            OwnerPassword = "owner123",
            // Disallow printing
            PrintPermission = false,
            // Additional restrictions (optional)
            ModifyDocumentPermission = false,
            ExtractContentPermission = false
        };

        // Attach security options to the PDF save options
        pdfSaveOptions.SecurityOptions = securityOptions;

        // Save the workbook as a secured PDF
        workbook.Save("SecuredNoPrint.pdf", pdfSaveOptions);
    }
}