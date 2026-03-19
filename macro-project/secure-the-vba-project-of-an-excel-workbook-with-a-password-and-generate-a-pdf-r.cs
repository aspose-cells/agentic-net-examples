using System;
using Aspose.Cells;
using Aspose.Cells.Vba;
using Aspose.Cells.Rendering.PdfSecurity;

class Program
{
    static void Main()
    {
        // Create a new workbook (macro-enabled)
        Workbook workbook = new Workbook();

        // Ensure there is at least one worksheet (required for VBA project)
        workbook.Worksheets.Add();

        // Protect the VBA project and lock it for viewing
        string vbaPassword = "VbaSecret123";
        workbook.VbaProject.Protect(true, vbaPassword);

        // Save the workbook as a macro‑enabled file
        workbook.Save("ProtectedWorkbook.xlsm", SaveFormat.Xlsm);

        // Configure PDF security options
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        PdfSecurityOptions security = new PdfSecurityOptions
        {
            OwnerPassword = "OwnerPass!",
            UserPassword = "UserPass!",
            PrintPermission = true,
            ModifyDocumentPermission = false,
            ExtractContentPermission = false
        };
        pdfOptions.SecurityOptions = security;

        // Save the workbook as a secured PDF
        workbook.Save("ProtectedWorkbook.pdf", pdfOptions);
    }
}