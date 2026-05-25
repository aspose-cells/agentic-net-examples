using Aspose.Cells;
using Aspose.Cells.Rendering.PdfSecurity;

class Program
{
    static void Main()
    {
        // Create a new workbook and add some content
        Workbook workbook = new Workbook();
        workbook.Worksheets[0].Cells["A1"].Value = "Secure PDF with user password";

        // Prepare PDF save options
        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();

        // Configure PDF security options
        PdfSecurityOptions securityOptions = new PdfSecurityOptions();
        securityOptions.UserPassword = "user123";          // Password required to open the PDF
        securityOptions.PrintPermission = false;          // Disallow printing
        securityOptions.ModifyDocumentPermission = false; // Disallow modifications
        securityOptions.ExtractContentPermission = false; // Disallow content extraction
        securityOptions.FullQualityPrintPermission = false; // Ensure low‑quality printing is also disabled

        // Attach security options to the save options
        pdfSaveOptions.SecurityOptions = securityOptions;

        // Save the workbook as a secured PDF
        workbook.Save("SecuredDocument.pdf", pdfSaveOptions);
    }
}