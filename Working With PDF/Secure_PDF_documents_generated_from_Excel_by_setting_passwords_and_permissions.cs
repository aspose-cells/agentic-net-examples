using System;
using Aspose.Cells;
using Aspose.Cells.Rendering.PdfSecurity;

namespace AsposeCellsPdfSecurityDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add sample data
            Workbook workbook = new Workbook();
            workbook.Worksheets[0].Cells["A1"].Value = "Secure PDF generated from Excel";

            // Configure PDF save options
            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();

            // Create and configure PDF security options
            PdfSecurityOptions securityOptions = new PdfSecurityOptions();
            securityOptions.OwnerPassword = "ownerPass123";   // Owner password (full access)
            securityOptions.UserPassword = "userPass123";     // User password (restricted access)

            // Set desired permissions
            securityOptions.PrintPermission = true;               // Allow printing
            securityOptions.FullQualityPrintPermission = true;   // Allow high‑quality printing
            securityOptions.ModifyDocumentPermission = false;    // Disallow modifications
            securityOptions.ExtractContentPermission = false;    // Disallow content extraction
            securityOptions.FillFormsPermission = true;          // Allow filling form fields
            securityOptions.AssembleDocumentPermission = false; // Disallow assembling pages

            // Assign security options to PDF save options
            pdfSaveOptions.SecurityOptions = securityOptions;

            // Save the workbook as a secured PDF
            workbook.Save("SecuredOutput.pdf", pdfSaveOptions);

            Console.WriteLine("Secured PDF created successfully.");
        }
    }
}