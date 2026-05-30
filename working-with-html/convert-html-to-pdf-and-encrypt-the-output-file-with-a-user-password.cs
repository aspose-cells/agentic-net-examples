using System;
using Aspose.Cells;
using Aspose.Cells.Rendering.PdfSecurity;

class Program
{
    static void Main()
    {
        // Load the HTML file into a workbook
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Html);
        Workbook workbook = new Workbook("input.html", loadOptions);

        // Create PDF save options
        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();

        // Configure PDF security (user password protection)
        PdfSecurityOptions securityOptions = new PdfSecurityOptions
        {
            UserPassword = "UserPassword123",   // password required to open the PDF
            OwnerPassword = "OwnerPassword123", // owner password (optional, gives full rights)
            PrintPermission = true,             // allow printing
            ModifyDocumentPermission = false    // disallow modifications
        };

        // Assign the security options to the PDF save options
        pdfSaveOptions.SecurityOptions = securityOptions;

        // Save the workbook as an encrypted PDF
        workbook.Save("output.pdf", pdfSaveOptions);
    }
}