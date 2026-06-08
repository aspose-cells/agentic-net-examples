using Aspose.Cells;
using Aspose.Cells.Rendering.PdfSecurity;

class ConvertProtectedWorkbookToPdf
{
    static void Main()
    {
        // Input protected workbook and output PDF paths
        string inputPath = "ProtectedWorkbook.xlsx";
        string outputPath = "ProtectedWorkbook.pdf";

        // Load the workbook with its password (if any)
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.Password = "workbookPwd"; // replace with actual password
        Workbook workbook = new Workbook(inputPath, loadOptions);

        // Prepare PDF save options
        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();

        // Create PDF security options to carry over protection metadata
        PdfSecurityOptions pdfSecurity = new PdfSecurityOptions();

        // Preserve workbook structure/window protection as PDF owner password
        if (workbook.IsWorkbookProtectedWithPassword)
        {
            pdfSecurity.OwnerPassword = "workbookPwd"; // same as workbook password
        }

        // Preserve write‑protection password as PDF user password
        if (workbook.Settings.WriteProtection.IsWriteProtected)
        {
            pdfSecurity.UserPassword = workbook.Settings.WriteProtection.Password;
        }

        // Assign the security options to the PDF save options
        pdfSaveOptions.SecurityOptions = pdfSecurity;

        // Save the workbook as a PDF while keeping protection metadata
        workbook.Save(outputPath, pdfSaveOptions);
    }
}