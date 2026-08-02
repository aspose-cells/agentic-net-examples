using System;
using Aspose.Cells;
using Aspose.Cells.Rendering.PdfSecurity;

namespace AsposeCellsPdfSecurityDemo
{
    // Author: Aspose.Cells .NET example
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add sample content
            workbook.Worksheets[0].Cells["A1"].Value = "Secure PDF Example";

            // Configure PDF save options
            PdfSaveOptions saveOptions = new PdfSaveOptions();

            // Set security options
            PdfSecurityOptions security = new PdfSecurityOptions();
            security.UserPassword = "user123";          // Password required to open the PDF
            security.OwnerPassword = "owner123";        // Owner password for full access
            security.PrintPermission = false;           // Restrict printing
            security.ModifyDocumentPermission = false; // Optional: restrict modifications
            security.ExtractContentPermission = false; // Optional: restrict content extraction

            // Assign security options to the PDF save options
            saveOptions.SecurityOptions = security;

            // Save the workbook as a secured PDF
            workbook.Save("SecuredOutput.pdf", saveOptions);
        }
    }
}