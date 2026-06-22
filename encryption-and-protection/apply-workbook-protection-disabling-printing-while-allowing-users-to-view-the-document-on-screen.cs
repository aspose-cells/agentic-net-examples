using System;
using Aspose.Cells;
using Aspose.Cells.Rendering.PdfSecurity;

namespace AsposeCellsProtectionDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add some sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].Value = "Confidential Data";

            // Protect the workbook structure with a password (prevents adding/removing sheets)
            workbook.Protect(ProtectionType.Structure, "workbookPwd");

            // Prepare PDF save options with security settings
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Configure PDF security: allow viewing but disable printing
            PdfSecurityOptions security = new PdfSecurityOptions
            {
                OwnerPassword = "ownerPwd",
                UserPassword = "userPwd",
                PrintPermission = false   // Disable printing
            };

            pdfOptions.SecurityOptions = security;

            // Save the workbook as a PDF with the defined security options
            workbook.Save("ProtectedDocument.pdf", pdfOptions);
        }
    }
}